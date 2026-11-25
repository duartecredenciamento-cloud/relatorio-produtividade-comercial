from __future__ import annotations

from datetime import date, datetime, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from socketserver import ThreadingMixIn
from urllib.parse import parse_qs, urlparse
import base64
import csv
import gzip
import hmac
import io
import json
import logging
import mimetypes
import os
import threading
import time
import webbrowser
import zipfile

import pandas as pd

from src.data_loader import (de_duplicate, list_matrix_files, read_aliases, read_manual_csv,
                             read_matrix, read_responsibles, stable_fingerprint)
from src.rules import (SERVICES, add_responsible, calculate_pending, classify_indicator_records,
                       is_active, is_inactive, is_completed, is_card_or_partnership, make_quality,
                       norm_text, pending_records, service_summary)

BASE = Path(__file__).resolve().parent
DATA = BASE / "dados"
APPDATA = Path(os.environ.get("APPDATA_DIR", str(BASE / "appdata")))
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
    "pending_red_days": 10,
    "stale_contact_days": 7,
    "adjustment_period_months": 12,
    "adjustment_warning_days": 60,
    "listen_on_lan": False,
    "port": 8765,
}


_DRIVE_SCOPE = ["https://www.googleapis.com/auth/drive.readonly"]
_drive_sync_lock = threading.Lock()
_last_drive_sync = 0.0


def sync_drive_matrices(force=False):
    """Mirror .xlsm files from one Google Drive folder into MATRICES_DIR.

    Drive sync is opt-in: without both environment variables, local behavior is unchanged.
    """
    global _last_drive_sync
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
    service_account_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not folder_id and not service_account_json:
        return None
    if not folder_id or not service_account_json:
        raise RuntimeError("Configure GOOGLE_DRIVE_FOLDER_ID e GOOGLE_SERVICE_ACCOUNT_JSON juntos.")

    try:
        interval = max(60, int(os.environ.get("DRIVE_SYNC_SECONDS", "900")))
    except ValueError as exc:
        raise RuntimeError("DRIVE_SYNC_SECONDS precisa ser um número de segundos.") from exc

    with _drive_sync_lock:
        now = time.monotonic()
        if not force and _last_drive_sync and now - _last_drive_sync < interval:
            return {"enabled": True, "synced": False, "count": None}

        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaIoBaseDownload
        except ImportError as exc:
            raise RuntimeError("Dependências Google ausentes; instale o requirements.txt atualizado.") from exc

        try:
            credentials_info = json.loads(service_account_json)
            credentials = service_account.Credentials.from_service_account_info(
                credentials_info, scopes=_DRIVE_SCOPE
            )
        except Exception as exc:
            raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON inválido ou incompleto.") from exc

        drive = build("drive", "v3", credentials=credentials, cache_discovery=False)
        remote_files = []
        page_token = None
        query = f"'{folder_id}' in parents and trashed = false"
        while True:
            response = drive.files().list(
                q=query,
                pageSize=1000,
                pageToken=page_token,
                fields="nextPageToken,files(id,name,mimeType,modifiedTime,size,md5Checksum)",
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            ).execute()
            for item in response.get("files", []):
                name = str(item.get("name", ""))
                if item.get("mimeType") == "application/vnd.google-apps.folder":
                    continue
                if name.startswith("~$") or Path(name).suffix.casefold() != ".xlsm":
                    continue
                # Files are intentionally kept directly inside the shared folder.
                if Path(name).name != name or name in {".", ".."} or "/" in name or "\\" in name:
                    raise RuntimeError("A pasta do Drive contém um nome de arquivo inválido.")
                remote_files.append(item)
            page_token = response.get("nextPageToken")
            if not page_token:
                break

        if not remote_files:
            raise RuntimeError("A pasta do Google Drive não contém matrizes .xlsm acessíveis.")
        folded_names = [str(f["name"]).casefold() for f in remote_files]
        if len(folded_names) != len(set(folded_names)):
            raise RuntimeError("Há nomes de matrizes repetidos na pasta do Drive; renomeie-os antes da sincronização.")

        root = resolve(load_config()["matrices_dir"])
        root.mkdir(parents=True, exist_ok=True)
        state_path = APPDATA / "drive_sync_state.json"
        try:
            old_state = json.loads(state_path.read_text(encoding="utf-8"))
            old_files = old_state.get("files", {}) if isinstance(old_state, dict) else {}
        except (OSError, ValueError, TypeError):
            old_files = {}
        new_files = {}
        current_ids = {str(item["id"]) for item in remote_files}
        downloaded = 0

        for item in remote_files:
            file_id = str(item["id"])
            name = str(item["name"])
            destination = root / name
            metadata = {
                "name": name,
                "modifiedTime": item.get("modifiedTime", ""),
                "size": str(item.get("size", "")),
                "md5Checksum": item.get("md5Checksum", ""),
            }
            previous = old_files.get(file_id, {})
            if previous == metadata and destination.is_file() and destination.stat().st_size > 0:
                new_files[file_id] = metadata
                continue

            request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
            buffer = io.BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            content = buffer.getvalue()
            if not content:
                raise RuntimeError(f"A matriz {name} foi baixada vazia; os dados atuais foram preservados.")
            try:
                with zipfile.ZipFile(io.BytesIO(content)) as workbook:
                    if workbook.testzip() is not None:
                        raise ValueError("arquivo XLSM danificado")
            except Exception as exc:
                raise RuntimeError(f"A matriz {name} não passou na validação do arquivo; os dados atuais foram preservados.") from exc

            temporary = destination.with_name(destination.name + ".drive-part")
            temporary.write_bytes(content)
            os.replace(temporary, destination)
            new_files[file_id] = metadata
            downloaded += 1

        # Remove only files previously managed by this integration and removed from Drive.
        for removed_id, old_metadata in old_files.items():
            if removed_id in current_ids or not isinstance(old_metadata, dict):
                continue
            old_name = str(old_metadata.get("name", ""))
            old_path = root / old_name
            if old_path.parent.resolve() == root.resolve() and old_path.suffix.casefold() == ".xlsm":
                try:
                    old_path.unlink(missing_ok=True)
                except OSError:
                    logger.warning("Não foi possível remover uma matriz removida do Drive.")

        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_tmp = state_path.with_name(state_path.name + ".tmp")
        state_tmp.write_text(json.dumps({"files": new_files}, ensure_ascii=False), encoding="utf-8")
        os.replace(state_tmp, state_path)
        _last_drive_sync = time.monotonic()
        logger.info("SINCRONIZAÇÃO GOOGLE DRIVE | matrizes=%s | baixadas=%s", len(remote_files), downloaded)
        return {"enabled": True, "synced": True, "count": len(remote_files), "downloaded": downloaded}

def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")))
    except Exception as exc:
        logger.warning("Configuração padrão utilizada: %s", exc)
    if os.environ.get("MATRICES_DIR"):
        cfg["matrices_dir"] = os.environ["MATRICES_DIR"]
    if os.environ.get("RESPONSIBLES_FILE"):
        cfg["responsibles_file"] = os.environ["RESPONSIBLES_FILE"]
    cfg["refresh_seconds"] = max(15, int(cfg.get("refresh_seconds", 60)))
    cfg["pending_days"] = max(0, int(cfg.get("pending_days", 7)))
    cfg["pending_red_days"] = max(cfg["pending_days"], int(cfg.get("pending_red_days", 10)))
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
        self.snapshot_cache = None

    def refresh(self, force=False):
        with self.lock:
            if self.refreshing:
                return
            self.refreshing = True
            self.snapshot_cache = None
        changed = False
        cfg = load_config()
        drive_sync = sync_drive_matrices(force=force)
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
                    changed = True
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
                    changed = True
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
                    changed = True
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
                if drive_sync and drive_sync.get("synced"):
                    self.last_message += f" Google Drive sincronizado: {drive_sync['count']} matrizes; {drive_sync['downloaded']} arquivo(s) novo(s) ou alterado(s)."
                self.last_error=""
            elif any_ok:
                self.last_message="Leitura parcial; os últimos dados válidos foram mantidos nos arquivos com falha."
            else:
                self.last_message="Nenhum arquivo pôde ser lido; os dados anteriores permanecem disponíveis."
            if self.last_error:
                self.last_message += " " + self.last_error
            if changed or self.last_date_rebuild != date.today():
                self.rebuild(cfg)
            else:
                with self.lock:
                    self.snapshot_cache = None
        except Exception as exc:
            self.last_error=f"{type(exc).__name__}: {exc}"
            self.last_message="Atualização indisponível; os últimos dados carregados foram preservados."
            self.file_status=[{"name":str(root),"state":"Falha de leitura","rows":0,"error":self.last_error,"warnings":[]}]
            logger.exception("ATUALIZAÇÃO INDISPONÍVEL")
        finally:
            with self.lock:
                self.refreshing=False
                self.snapshot_cache = None

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
            records["pending_class"]=[calculate_pending(m,d,now,cfg["pending_days"],cfg.get("pending_red_days",10)) for m,d in zip(records.marker_status,records.last_contact)]
        with self.lock:
            self.records=records
            self.last_date_rebuild=now
            self.snapshot_cache = None

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
        cacheable = not any((values or [""])[0] for values in query.values())
        with self.lock:
            if cacheable and not self.refreshing and self.snapshot_cache is not None:
                return self.snapshot_cache
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
        pending=pending_records(records,cfg["pending_days"],now,cfg.get("pending_red_days",10)) if not records.empty else records
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
            inactive=g[g.status.map(is_inactive)] if not g.empty else g
            services.append({"key":service,"name":service,"active":len(a),"inactive":len(inactive),"completed":len(c),"inProgress":len(a)-len(c),"pending":len(p),"achieved":len(c),"target":target,"targetPeriod":str(goal.period) if goal is not None else None})
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
        response={"mode":"live","statusLabel":"ATUALIZAÇÃO EM ANDAMENTO" if refreshing else ("FONTE LOCAL · COM AVISOS" if last_error or any(f.get("error") for f in file_status) else "FONTE LOCAL · MATRIZES"),"statusMessage":message,"updatedAt":updated,"referenceDate":now.isoformat(),"kpis":{"active":len(active),"inactive":int(records.status.map(is_inactive).sum()) if not records.empty else 0,"completed":len(completed),"inProgress":len(active)-len(completed),"pending":len(pending),"owners":len({x["name"] for x in owner_groups if x["name"]!="Responsável não identificado"}),"unassigned":sum(x["name"]=="Responsável não identificado" for x in owner_groups),"completionRate":round(len(completed)/len(active)*100) if len(active) else 0,"markerReview":sum(x["marker"]=="Revisar" for x in mapped)},"services":services,"processes":mapped,"pendingRecords":pending_mapped,"owners":owner_groups,"priorities":priorities,"indicators":indicator_counts,"quality":quality_rows,"qualitySummary":qrecords,"fileStatus":file_status,"filterOptions":all_opts,"config":{"pendingDays":cfg["pending_days"],"pendingRedDays":cfg.get("pending_red_days",10),"refreshSeconds":cfg["refresh_seconds"],"staleContactDays":cfg["stale_contact_days"]},"duplicates":duplicate_count}
        if cacheable:
            with self.lock:
                if not self.refreshing:
                    self.snapshot_cache=response
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
        compressed="gzip" in self.headers.get("Accept-Encoding","").lower()
        if compressed:data=gzip.compress(data,compresslevel=5)
        self.send_response(status);self.send_header("Content-Type","application/json; charset=utf-8");self.send_header("Content-Length",str(len(data)));self.send_header("Cache-Control","no-store")
        if compressed:self.send_header("Content-Encoding","gzip");self.send_header("Vary","Accept-Encoding")
        self.end_headers();self.wfile.write(data)
    def authorize(self):
        if not os.environ.get("PORT") and os.environ.get("REQUIRE_AUTH","").casefold() not in {"1","true","yes"}:
            return True
        user=os.environ.get("DASHBOARD_USER","")
        password=os.environ.get("DASHBOARD_PASSWORD","")
        if not user or not password:
            self.send_json({"error":"Configure DASHBOARD_USER e DASHBOARD_PASSWORD como variáveis privadas do serviço."},503)
            return False
        expected="Basic "+base64.b64encode(f"{user}:{password}".encode("utf-8")).decode("ascii")
        if hmac.compare_digest(self.headers.get("Authorization", ""),expected):
            return True
        self.send_response(401);self.send_header("WWW-Authenticate",'Basic realm="Painel Duarte", charset="UTF-8"');self.send_header("Content-Length","0");self.end_headers()
        return False
    def do_GET(self):
        parsed=urlparse(self.path)
        if parsed.path=="/api/health":return self.send_json({"ok":True,"refreshing":STORE.refreshing,"updatedAt":STORE.last_success.isoformat(timespec="seconds") if STORE.last_success else None})
        if not self.authorize():return
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
        if not self.authorize():return
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
    cfg=load_config()
    port=int(os.environ.get("PORT",cfg["port"]))
    host="0.0.0.0" if os.environ.get("PORT") or cfg.get("listen_on_lan",False) else "127.0.0.1"
    server=Server((host,port),Handler)
    url=f"http://127.0.0.1:{port}"
    threading.Thread(target=STORE.refresh,daemon=True).start()
    threading.Thread(target=refresh_loop,daemon=True).start()
    print(f"Painel Duarte iniciado em {host}:{port}")
    print("A leitura inicial das matrizes está em andamento. Deixe esta janela aberta.")
    print("Para encerrar, pressione Ctrl+C.")
    if os.environ.get("DUARTE_NO_BROWSER")!="1":threading.Timer(1.2,lambda:webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:print("\nEncerrando o painel.")
    finally:server.server_close()

if __name__=="__main__":main()
