from __future__ import annotations

from datetime import date, datetime, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from socketserver import ThreadingMixIn
from urllib.parse import parse_qs, urlparse
import csv
import json
import logging
import mimetypes
import os
import threading
import time
import webbrowser

import pandas as pd

from src.data_loader import (de_duplicate, list_matrix_files, read_aliases, read_manual_csv,
                             read_matrix, read_responsibles, stable_fingerprint)
from src.rules import (SERVICES, add_responsible, calculate_pending, classify_indicator_records,
                       is_active, is_completed, is_card_or_partnership, make_quality,
                       norm_text, pending_records, service_summary)

BASE = Path(__file__).resolve().parent
DATA = BASE / "dados"
APPDATA = BASE / "appdata"
APPDATA.mkdir(exist_ok=True)
LOG = APPDATA / "atualizacoes.log"
GOALS = APPDATA / "metas_manuais.csv"
MARKERS = APPDATA / "marcacoes_manuais.csv"
ALIASES = BASE / "aliases.csv"
CONFIG = BASE / "config.json"

logger = logging.getLogger("duarte_produtividade")
logger.setLevel(logging.INFO)
if not logger.handlers:
    h = logging.FileHandler(LOG, encoding="utf-8")
    h.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
    logger.addHandler(h)

DEFAULT_CONFIG = {
    "matrices_dir": "dados/matrizes",
    "responsibles_file": "dados/clientes_credenciamento_responsaveis.xlsx",
    "refresh_seconds": 60,
    "pending_days": 7,
    "stale_contact_days": 7,
    "adjustment_period_months": 12,
    "adjustment_warning_days": 60,
    "listen_on_lan": False,
    "port": 8765,
}

def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")))
    except Exception as exc:
        logger.warning("Configuração padrão utilizada: %s", exc)
    cfg["refresh_seconds"] = max(15, int(cfg.get("refresh_seconds", 60)))
    cfg["pending_days"] = max(0, int(cfg.get("pending_days", 7)))
    cfg["stale_contact_days"] = max(0, int(cfg.get("stale_contact_days", 7)))
    cfg["port"] = max(1024, min(65535, int(cfg.get("port", 8765))))
    return cfg

def resolve(value):
    p = Path(str(value).strip().strip('"'))
    return p if p.is_absolute() else (BASE / p).resolve()

def iso(value):
    if value is None or value == "" or pd.isna(value):
        return None
    try:
        return pd.Timestamp(value).date().isoformat()
    except Exception:
        return None

def empty_records():
    return pd.DataFrame(columns=["record_id","client_code","client","insurance_id","operator","category","priority","source","source_row","service","status","service_operator_source","last_contact","detail","notes","marker_status","marker_evidence","contract_date","last_adjustment","responsible","client_match"])

class DashboardStore:
    def __init__(self):
        self.lock = threading.RLock()
        self.matrix_cache = {}
        self.records = empty_records()
        self.roster = pd.DataFrame()
        self.roster_fp = None
        self.file_status = []
        self.last_success = None
        self.last_message = "Aguardando a primeira leitura."
        self.last_error = ""
        self.duplicate_count = 0
        self.last_date_rebuild = None
        self.refreshing = False

    def refresh(self, force=False):
        with self.lock:
            if self.refreshing:
                return
            self.refreshing = True
        cfg = load_config()
        root = resolve(cfg["matrices_dir"])
        new_status = []
        any_ok = False
        errors = False
        try:
            if not root.is_dir():
                raise FileNotFoundError(f"Pasta de matrizes indisponível: {root}")
            paths = list_matrix_files(root)
            present = {str(p.resolve()) for p in paths}
            with self.lock:
                for removed in set(self.matrix_cache) - present:
                    del self.matrix_cache[removed]
                    logger.info("REMOVIDO | %s", removed)
            for p in paths:
                key = str(p.resolve())
                prev = self.matrix_cache.get(key)
                try:
                    st = p.stat()
                    quick = (st.st_size, st.st_mtime_ns)
                    if prev and prev["fingerprint"] == quick and not force:
                        new_status.append({"name":p.name,"state":"Sem alteração","rows":prev["rows"],"error":"","warnings":prev["warnings"]})
                        any_ok = True
                        continue
                    fp = stable_fingerprint(p, settle_seconds=.2, retries=2)
                    if prev and prev["fingerprint"] == fp and not force:
                        new_status.append({"name":p.name,"state":"Estável — sem alteração","rows":prev["rows"],"error":"","warnings":prev["warnings"]})
                        any_ok = True
                        continue
                    parsed = read_matrix(p)
                    self.matrix_cache[key] = {"fingerprint":fp,"records":parsed["records"],"rows":parsed["rows"],"warnings":parsed["warnings"]}
                    new_status.append({"name":p.name,"state":"Atualizado","rows":parsed["rows"],"error":"","warnings":parsed["warnings"]})
                    logger.info("LIDO/ALTERADO | %s | linhas=%s | registros=%s | avisos=%s",p.name,parsed["rows"],len(parsed["records"]),len(parsed["warnings"]))
                    any_ok=True
                except Exception as exc:
                    errors=True
                    msg=f"{type(exc).__name__}: {exc}"
                    new_status.append({"name":p.name,"state":"Falha — dados anteriores preservados" if prev else "Falha de leitura","rows":prev["rows"] if prev else 0,"error":msg,"warnings":prev["warnings"] if prev else []})
                    logger.exception("FALHA | %s",p.name)
            # The responsible roster is checked every cycle; a failed read never replaces a valid roster.
            roster_path=resolve(cfg["responsibles_file"])
            try:
                fp=stable_fingerprint(roster_path,settle_seconds=.15,retries=2)
                if fp!=self.roster_fp or self.roster.empty:
                    roster,_=read_responsibles(roster_path)
                    self.roster=roster
                    self.roster_fp=fp
                    logger.info("CADASTRO ATUALIZADO | %s | registros=%s",roster_path.name,len(roster))
            except Exception as exc:
                self.last_error=f"Cadastro de responsáveis: {exc}"
                logger.warning("CADASTRO PRESERVADO | %s",exc)
            self.file_status=new_status
            if not paths:
                self.last_message="Nenhuma matriz .xlsm encontrada; confira a pasta configurada."
            elif any_ok and not errors:
                self.last_success=datetime.now().astimezone()
                self.last_message=f"Leitura concluída: {len(paths)} matriz(es), {sum(x['rows'] for x in new_status)} linhas."
                self.last_error=""
            elif any_ok:
                self.last_message="Leitura parcial; os últimos dados válidos foram mantidos nos arquivos com falha."
            else:
                self.last_message="Nenhum arquivo pôde ser lido; os dados anteriores permanecem disponíveis."
            if self.last_error:
                self.last_message += " " + self.last_error
            self.rebuild(cfg)
        except Exception as exc:
            self.last_error=f"{type(exc).__name__}: {exc}"
            self.last_message="Atualização indisponível; os últimos dados carregados foram preservados."
            self.file_status=[{"name":str(root),"state":"Falha de leitura","rows":0,"error":self.last_error,"warnings":[]}]
            logger.exception("ATUALIZAÇÃO INDISPONÍVEL")
        finally:
            with self.lock:
                self.refreshing=False

    def rebuild(self,cfg=None):
        cfg=cfg or load_config()
        parts=[pd.DataFrame(v["records"]) for v in self.matrix_cache.values() if v.get("records")]
        records=pd.concat(parts,ignore_index=True) if parts else empty_records()
        if not records.empty:
            records,self.duplicate_count=de_duplicate(records)
            records=add_responsible(records,self.roster,read_aliases(ALIASES))
            overrides=read_manual_csv(MARKERS,["record_id","marcado"])
            if not overrides.empty:
                ov=dict(zip(overrides.record_id,overrides.marcado))
                records["marker_status"]=[ov.get(x,m) for x,m in zip(records.record_id,records.marker_status)]
            records["marker_status"]=records.marker_status.fillna("Revisar")
        else:
            self.duplicate_count=0
        # Recalculate date-dependent classifications without reparsing workbooks.
        now=date.today()
        if not records.empty:
            records["pending_class"]=[calculate_pending(m,d,now,cfg["pending_days"]) for m,d in zip(records.marker_status,records.last_contact)]
        with self.lock:
            self.records=records
            self.last_date_rebuild=now

    def mark(self, record_id, marked):
        if marked not in {"Sim","Não"}:
            raise ValueError("Marcação precisa ser Sim ou Não.")
        if self.records.empty or record_id not in set(self.records.record_id):
            raise KeyError("Registro não encontrado na carga atual.")
        rows=read_manual_csv(MARKERS,["record_id","marcado"])
        rows=rows[rows.record_id.astype(str)!=str(record_id)]
        rows=pd.concat([rows,pd.DataFrame([{"record_id":record_id,"marcado":marked}])],ignore_index=True)
        rows.to_csv(MARKERS,index=False,encoding="utf-8-sig")
        self.rebuild()

    def save_goal(self,service,target,period):
        if service not in SERVICES: raise ValueError("Serviço inválido.")
        if target is not None and (int(target)<0 or int(target)>10_000_000): raise ValueError("Meta fora do intervalo permitido.")
        rows=read_manual_csv(GOALS,["service","target","period"])
        rows=rows[rows.service.astype(str)!=service]
        if target is not None: rows=pd.concat([rows,pd.DataFrame([{"service":service,"target":int(target),"period":str(period)[:50]}])],ignore_index=True)
        rows.to_csv(GOALS,index=False,encoding="utf-8-sig")
        self.rebuild()

    def snapshot(self, query):
        with self.lock:
            records=self.records.copy()
            file_status=list(self.file_status)
            last_success=self.last_success
            message=self.last_message
            refreshing=self.refreshing
            duplicate_count=self.duplicate_count
            last_error=self.last_error
        cfg=load_config()
        all_records=records.copy()
        records=filter_records(records,query)
        now=date.today()
        active=records[records.status.map(is_active)] if not records.empty else records
        completed=active[active.status.map(is_completed)] if not active.empty else active
        pending=pending_records(records,cfg["pending_days"],now) if not records.empty else records
        services=[]
        goals=read_manual_csv(GOALS,["service","target","period"])
        goal_map={r.service:r for _,r in goals.iterrows()} if not goals.empty else {}
        for service in SERVICES:
            g=records[records.service==service] if not records.empty else records
            a=g[g.status.map(is_active)] if not g.empty else g
            c=a[a.status.map(is_completed)] if not a.empty else a
            p=pending[pending.service==service] if not pending.empty else pending
            goal=goal_map.get(service)
            target=int(float(goal.target)) if goal is not None and str(goal.target).strip() else None
            services.append({"key":service,"name":service,"active":len(a),"completed":len(c),"inProgress":len(a)-len(c),"pending":len(p),"achieved":len(c),"target":target,"targetPeriod":str(goal.period) if goal is not None else None})
        indicators=classify_indicator_records(records)
        indicator_counts={k:{"count":len(v),"clients":int(v.client.nunique()) if not v.empty else 0} for k,v in indicators.items()}
        quality=make_quality(records,file_status,now,cfg["stale_contact_days"]) if not records.empty else pd.DataFrame()
        # Every record whose icon is not explicitly represented remains reviewable.
        if not records.empty:
            marker_review=records[records.marker_status=="Revisar"]
            for _,r in marker_review.iterrows():
                quality.loc[len(quality)]={"tipo":"Bolinha verde não confirmada","cliente":r.client,"operadora":r.operator,"serviço":r.service,"status":r.status,"último_contato":r.last_contact,"fonte":r.source,"detalhe":"A marca deve ser revisada manualmente; cor de status não é tomada como bolinha de meta."}
        qrecords=[]
        if not quality.empty:
            group_cols=["tipo"]
            for typ,g in quality.groupby("tipo",dropna=False):
                note=str(g.iloc[0].get("detalhe") or "Verifique os registros detalhados.")
                qrecords.append({"type":str(typ),"count":len(g),"note":note})
        owner_groups=[]
        if not records.empty:
            for owner,g in records.groupby("responsible",dropna=False):
                a=g[g.status.map(is_active)]; c=a[a.status.map(is_completed)]; p=pending[pending.responsible==owner] if not pending.empty else pending
                owner_groups.append({"name":str(owner),"processes":len(a),"completed":len(c),"inProgress":len(a)-len(c),"pending":len(p),"clients":int(g.client.nunique())})
            owner_groups.sort(key=lambda x:(x["name"]=="Responsável não identificado",x["name"].casefold()))
        priorities=[]
        cred=records[records.service=="CREDENCIAMENTO"] if not records.empty else records
        if not cred.empty:
            for priority,g in cred.groupby("priority",dropna=False): priorities.append({"name":str(priority or "Sem prioridade"),"processes":len(g)})
            priorities.sort(key=lambda x:-x["processes"])
        mapped=[]
        for _,r in records.iterrows():
            age=(now-r.last_contact).days if r.get("last_contact") is not None and not pd.isna(r.get("last_contact")) else None
            mapped.append({"recordId":str(r.get("record_id","")),"code":str(r.get("client_code") or r.get("insurance_id") or ""),"client":str(r.get("client") or ""),"operator":str(r.get("operator") or ""),"service":str(r.get("service") or ""),"category":str(r.get("category") or ""),"priority":str(r.get("priority") or ""),"status":str(r.get("status") or ""),"owner":str(r.get("responsible") or "Responsável não identificado"),"lastContact":iso(r.get("last_contact")),"days":age,"marker":str(r.get("marker_status") or "Revisar"),"pendingClass":str(r.get("pending_class") or ""),"source":str(r.get("source") or ""),"sourceRow":int(r.get("source_row") or 0),"notes":str(r.get("notes") or r.get("detail") or ""),"clientMatch":str(r.get("client_match") or "")})
        pending_ids={str(x) for x in pending.record_id} if not pending.empty else set()
        pending_mapped=[p for p in mapped if p["recordId"] in pending_ids]
        all_opts={"owners":sorted({str(x) for x in all_records.get("responsible",[]) if str(x).strip()}),"priorities":sorted({str(x) for x in all_records.get("priority",[]) if str(x).strip()}),"categories":sorted({str(x) for x in all_records.get("category",[]) if str(x).strip()}),"services":list(SERVICES),"operators":sorted({str(x) for x in all_records.get("operator",[]) if str(x).strip()})}
        quality_rows=[]
        if not quality.empty:
            for _,r in quality.iterrows():
                quality_rows.append({"type":str(r.get("tipo") or "Alerta"),"client":str(r.get("cliente") or ""),"operator":str(r.get("operadora") or ""),"service":str(r.get("serviço") or ""),"status":str(r.get("status") or ""),"lastContact":iso(r.get("último_contato")),"source":str(r.get("fonte") or ""),"detail":str(r.get("detalhe") or "")})
        updated=last_success.isoformat(timespec="seconds") if last_success else None
        response={"mode":"live","statusLabel":"ATUALIZAÇÃO EM ANDAMENTO" if refreshing else ("FONTE LOCAL · COM AVISOS" if last_error or any(f.get("error") for f in file_status) else "FONTE LOCAL · MATRIZES"),"statusMessage":message,"updatedAt":updated,"referenceDate":now.isoformat(),"kpis":{"active":len(active),"completed":len(completed),"inProgress":len(active)-len(completed),"pending":len(pending),"owners":len({x["name"] for x in owner_groups if x["name"]!="Responsável não identificado"}),"unassigned":sum(x["name"]=="Responsável não identificado" for x in owner_groups),"completionRate":round(len(completed)/len(active)*100) if len(active) else 0,"markerReview":sum(x["marker"]=="Revisar" for x in mapped)},"services":services,"processes":mapped,"pendingRecords":pending_mapped,"owners":owner_groups,"priorities":priorities,"indicators":indicator_counts,"quality":quality_rows,"qualitySummary":qrecords,"fileStatus":file_status,"filterOptions":all_opts,"config":{"pendingDays":cfg["pending_days"],"refreshSeconds":cfg["refresh_seconds"],"staleContactDays":cfg["stale_contact_days"]},"duplicates":duplicate_count}
        return response

def filter_records(df,query):
    if df.empty:return df
    out=df
    for qkey,col in [("owner","responsible"),("priority","priority"),("category","category"),("service","service")]:
        val=(query.get(qkey) or [""])[0]
        if val:out=out[out[col].astype(str)==val]
    operator=(query.get("operator") or [""])[0].strip().casefold()
    search=(query.get("search") or [""])[0].strip().casefold()
    for term,cols in [(operator,["operator"]),(search,["client","operator","service","category","priority","status","responsible","client_code"])]:
        if term:
            mask=False
            for col in cols:mask=mask|out[col].fillna("").astype(str).str.casefold().str.contains(term,regex=False)
            out=out[mask]
    start=(query.get("date_from") or [""])[0]; end=(query.get("date_to") or [""])[0]
    if start:out=out[out.last_contact.map(lambda x:x is not None and pd.notna(x) and x.isoformat()>=start)]
    if end:out=out[out.last_contact.map(lambda x:x is not None and pd.notna(x) and x.isoformat()<=end)]
    period=(query.get("period") or ["all"])[0]
    if period in {"7","30","90"}:
        cutoff=date.today()-timedelta(days=int(period))
        out=out[out.last_contact.map(lambda x:x is not None and pd.notna(x) and x>=cutoff)]
    return out

STORE=DashboardStore()

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs):super().__init__(*args,directory=str(BASE),**kwargs)
    def log_message(self,fmt,*args):logger.info("HTTP | %s",fmt%args)
    def send_json(self,obj,status=200):
        data=json.dumps(obj,ensure_ascii=False,default=str).encode("utf-8")
        self.send_response(status);self.send_header("Content-Type","application/json; charset=utf-8");self.send_header("Content-Length",str(len(data)));self.send_header("Cache-Control","no-store");self.end_headers();self.wfile.write(data)
    def do_GET(self):
        parsed=urlparse(self.path)
        if parsed.path=="/api/health":return self.send_json({"ok":True,"refreshing":STORE.refreshing,"updatedAt":STORE.last_success.isoformat(timespec="seconds") if STORE.last_success else None})
        if parsed.path=="/api/dashboard":return self.send_json(STORE.snapshot(parse_qs(parsed.query)))
        relative="index.html" if parsed.path=="/" else parsed.path.lstrip("/")
        allowed={"index.html","assets/styles.css","assets/app.css","assets/data-service.js","assets/app.js","assets/brand-logo.png"}
        if relative not in allowed:return self.send_json({"error":"Recurso não encontrado."},404)
        path=(BASE/relative).resolve()
        if not path.is_file() or BASE not in path.parents:return self.send_json({"error":"Recurso não encontrado."},404)
        payload=path.read_bytes(); mime=mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(200);self.send_header("Content-Type",mime+ ("; charset=utf-8" if mime.startswith("text/") or "javascript" in mime else ""));self.send_header("Content-Length",str(len(payload)));self.send_header("X-Content-Type-Options","nosniff");self.end_headers();self.wfile.write(payload)
    def do_POST(self):
        parsed=urlparse(self.path)
        try:
            n=int(self.headers.get("Content-Length","0"));body=json.loads(self.rfile.read(min(n,20000)) or b"{}")
            if parsed.path=="/api/refresh":
                threading.Thread(target=STORE.refresh,kwargs={"force":True},daemon=True).start()
                return self.send_json({"ok":True,"message":"Atualização iniciada."},202)
            if parsed.path=="/api/marker":
                STORE.mark(str(body.get("record_id","")),str(body.get("marked","")))
                return self.send_json({"ok":True})
            if parsed.path=="/api/goals":
                STORE.save_goal(str(body.get("service","")),body.get("target"),str(body.get("period","manual")))
                return self.send_json({"ok":True})
            return self.send_json({"error":"Endpoint não encontrado."},404)
        except KeyError as exc:return self.send_json({"error":str(exc)},404)
        except Exception as exc:return self.send_json({"error":str(exc)},400)

class Server(ThreadingHTTPServer):
    daemon_threads=True
    allow_reuse_address=True

def refresh_loop():
    while True:
        time.sleep(load_config()["refresh_seconds"])
        if not STORE.refreshing:STORE.refresh()

def main():
    cfg = load_config()
    railway_port = os.environ.get("PORT")

    # Railway supplies PORT and requires the process to bind on all interfaces.
    if railway_port:
        host = "0.0.0.0"
        port = int(railway_port)
        if not 1 <= port <= 65535:
            raise ValueError("A variável PORT precisa estar entre 1 e 65535.")
        url = None
    else:
        host = "0.0.0.0" if cfg.get("listen_on_lan", False) else "127.0.0.1"
        port = cfg["port"]
        url = f"http://127.0.0.1:{port}"

    server = Server((host, port), Handler)
    threading.Thread(target=STORE.refresh, daemon=True).start()
    threading.Thread(target=refresh_loop, daemon=True).start()

    if railway_port:
        print(f"Painel Duarte iniciado em {host}:{port} (Railway).")
        print("A leitura inicial das matrizes está em andamento.")
    else:
        print(f"Painel Duarte disponível em {url}")
        print("A leitura inicial das matrizes está em andamento. Deixe esta janela aberta.")
        print("Para encerrar, pressione Ctrl+C.")
        if os.environ.get("DUARTE_NO_BROWSER") != "1":
            threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando o painel.")
    finally:
        server.server_close()

if __name__=="__main__":main()
