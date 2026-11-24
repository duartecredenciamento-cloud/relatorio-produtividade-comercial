from __future__ import annotations

from datetime import date, datetime, timedelta
import re
import unicodedata

import pandas as pd

SERVICES = {
    "CREDENCIAMENTO": ("cred_status", "cred_operator", "cred_date", "cred_detail", "cred_notes"),
    "REAJUSTE": ("reaj_status", "reaj_operator", "reaj_date", "reaj_detail", "reaj_notes"),
    "EXTENSÃO": ("ext_status", "ext_operator", "ext_date", "ext_detail", "ext_notes"),
    "APROXIMAÇÃO": ("apx_status", "apx_operator", "apx_date", "apx_detail", "apx_notes"),
}
INACTIVE = {"NAO SE APLICA", "CLIENTE SEM INTERESSE", "AGUARDANDO EDITAIS"}
COMPLETED = {"CONCLUIDO", "JA CREDENCIADO", "REALIZADO", "FINALIZADO", "IMPLANTADO"}
CARD_TERMS = ("CARTAO DE DESCONTO", "PARCERIA")


def norm_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()


def norm_key(value: object) -> str:
    return norm_text(value).replace(" ", "")


def is_inactive(status: object) -> bool:
    s = norm_text(status)
    return bool(s) and any(s == x or s.startswith(x + " ") for x in INACTIVE)


def is_active(status: object) -> bool:
    return bool(norm_text(status)) and not is_inactive(status)


def is_completed(status: object) -> bool:
    s = norm_text(status)
    return s in COMPLETED or any(s.startswith(x + " ") for x in COMPLETED)


def is_card_or_partnership(category: object) -> bool:
    c = norm_text(category)
    return any(term in c for term in CARD_TERMS)


def calculate_pending(marked: object, last_contact: object, today: date, days: int = 7, red_days: int = 10) -> str:
    """Classifica a data pelo critério da planilha: verde até 7 dias, laranja de 8 a 10, vermelho acima de 10.

    O parâmetro `marked` é mantido por compatibilidade, mas a regra da planilha usa a data diretamente.
    """
    if last_contact is None or pd.isna(last_contact):
        return "META PENDENTE — DATA NÃO INFORMADA"
    try:
        d = pd.Timestamp(last_contact).date()
    except Exception:
        return "META PENDENTE — DATA NÃO INFORMADA"
    age = (today - d).days
    if age > max(0, int(red_days)):
        return "META PENDENTE — VERMELHA"
    if age > max(0, int(days)):
        return "META PENDENTE — LARANJA"
    return "Dentro do prazo"


def classify_indicator_records(records: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Critérios explícitos dos seis indicadores de credenciamento do prompt.

    A referência histórica usa recortes por status/categoria; como seus filtros
    pivot estão parcialmente salvos como 'Vários itens', os indicadores abaixo
    aplicam os rótulos inequívocos do pedido e exibem o detalhamento filtrável.
    """
    if records.empty:
        return {k: records.copy() for k in (
            "FINALIZANDO AGUARDANDO CONTRATO", "FINALIZANDO CREDENCIAMENTO (EM NEGOCIAÇÃO)",
            "PROCESSOS FINALIZANDO CARTÃO (AG. CONTRATO/IMPLANTAÇÃO)",
            "PROCESSOS FINALIZANDO OPERADORA (AG. CONTRATO/IMPLANTAÇÃO)",
            "EM NEG. CARTÃO E PARCERIA (EM NEGOCIAÇÃO)", "EM NEG. OPERADORAS (EM NEGOCIAÇÃO)")}
    cred = records[(records.service == "CREDENCIAMENTO") & records.status.map(is_active)].copy()
    st = cred.status.map(norm_text)
    cat = cred.category.map(is_card_or_partnership)
    waiting = st.str.contains(r"AGUARDANDO CONTRATO", regex=True, na=False)
    implementation = st.str.contains(r"IMPLANTACAO", regex=True, na=False)
    negotiation = st.str.contains(r"EM NEGOCIACAO", regex=True, na=False)
    return {
        "FINALIZANDO AGUARDANDO CONTRATO": cred[waiting],
        "FINALIZANDO CREDENCIAMENTO (EM NEGOCIAÇÃO)": cred[negotiation],
        "PROCESSOS FINALIZANDO CARTÃO (AG. CONTRATO/IMPLANTAÇÃO)": cred[cat & (waiting | implementation)],
        "PROCESSOS FINALIZANDO OPERADORA (AG. CONTRATO/IMPLANTAÇÃO)": cred[(~cat) & (waiting | implementation)],
        "EM NEG. CARTÃO E PARCERIA (EM NEGOCIAÇÃO)": cred[cat & negotiation],
        "EM NEG. OPERADORAS (EM NEGOCIAÇÃO)": cred[(~cat) & negotiation],
    }


def service_summary(records: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for service in SERVICES:
        group = records[records.service == service] if not records.empty else records
        active = group[group.status.map(is_active)] if not group.empty else group
        done = active[active.status.map(is_completed)] if not active.empty else active
        rows.append({"Serviço": service, "Processos ativos": len(active), "Concluídos": len(done),
                     "Em andamento": len(active) - len(done), "Percentual concluído": len(done) / len(active) if len(active) else None})
    return pd.DataFrame(rows)


def add_responsible(records: pd.DataFrame, roster: pd.DataFrame, aliases: pd.DataFrame) -> pd.DataFrame:
    """Associa responsáveis sem fuzzy matching: ambiguidades ficam para revisão."""
    if records.empty:
        out = records.copy()
        out["responsible"] = []
        out["client_match"] = []
        return out
    roster = roster.copy() if roster is not None else pd.DataFrame()
    aliases = aliases.copy() if aliases is not None else pd.DataFrame()
    if roster.empty:
        out = records.copy()
        out["responsible"] = "Responsável não identificado"
        out["client_match"] = "Cadastro não carregado"
        return out

    # Resolver apelidos explicitamente cadastrados para um nome oficial.
    alias_lookup = {}
    if {"apelido_cliente", "cliente_cadastro"}.issubset(aliases.columns):
        for _, r in aliases.dropna(subset=["apelido_cliente", "cliente_cadastro"]).iterrows():
            alias_lookup[norm_key(r.apelido_cliente)] = str(r.cliente_cadastro).strip()

    roster["_client_key"] = roster["client"].map(norm_key)
    key_to_official = {}
    for key, part in roster.groupby("_client_key"):
        names = sorted(set(part.client.astype(str)))
        key_to_official[key] = names[0] if len(names) == 1 else None
    all_rows = roster[roster.client.map(lambda x: norm_key(x) == norm_key("TODOS OS CLIENTES"))]
    client_rows = roster[roster.client.map(lambda x: norm_key(x) != norm_key("TODOS OS CLIENTES"))]
    mapping = {}
    for _, r in client_rows.iterrows():
        mapping.setdefault(norm_key(r.client), []).append(r)

    owners, match_states = [], []
    for _, rec in records.iterrows():
        client = str(rec.get("client", ""))
        alias_name = alias_lookup.get(norm_key(client))
        key = norm_key(alias_name or client)
        target_names = key_to_official.get(key)
        if not target_names:
            owners.append("Responsável não identificado")
            match_states.append("Nome sem correspondência única")
            continue
        if key == norm_key("TODOS OS CLIENTES"):
            owners.append("Responsável não identificado")
            match_states.append("Cliente não resolvido")
            continue
        rows = mapping.get(key, [])
        priority = norm_text(rec.get("priority", ""))
        is_cred = priority == "CREDENCIAMENTO"
        is_meta_cred = priority == "META CREDENCIAMENTO"
        if is_cred or is_meta_cred:
            scope = "CREDENCIAMENTO" if is_cred else "META CREDENCIAMENTO"
            special = all_rows[all_rows.priority.map(norm_text) == norm_text(scope)]
            names = sorted(set(special.responsible.astype(str)))
            owners.append(names[0] if len(names) == 1 else "Responsável não identificado")
            match_states.append("Prioridade especial do cadastro" if len(names) == 1 else "Prioridade especial sem regra única")
        else:
            names = sorted(set(str(r.responsible) for r in rows if "TODAS AS PRIORIDADES" in norm_text(r.priority)))
            owners.append(names[0] if len(names) == 1 else "Responsável não identificado")
            match_states.append("Cadastro oficial" if len(names) == 1 else "Cadastro sem responsável único")
    out = records.copy()
    out["responsible"] = owners
    out["client_match"] = match_states
    return out


def make_quality(records: pd.DataFrame, files: list[dict], today: date, days: int) -> pd.DataFrame:
    if records.empty:
        return pd.DataFrame(columns=["tipo", "cliente", "operadora", "serviço", "status", "último_contato", "fonte", "detalhe"])
    rows = []
    for _, r in records.iterrows():
        if r.get("client_match") not in ("Cadastro oficial", "Prioridade especial do cadastro"):
            rows.append({"tipo": "Cliente/responsável para revisão", "cliente": r.client, "operadora": r.operator,
                         "serviço": r.service, "status": r.status, "último_contato": r.last_contact,
                         "fonte": r.source, "detalhe": r.client_match})
        if pd.isna(r.last_contact):
            rows.append({"tipo": "Data de contato ausente ou inválida", "cliente": r.client, "operadora": r.operator,
                         "serviço": r.service, "status": r.status, "último_contato": None,
                         "fonte": r.source, "detalhe": "Verificar a coluna ÚLTIMO CONTATO correspondente ao serviço"})
        elif is_active(r.status) and (today - pd.Timestamp(r.last_contact).date()).days > days:
            age = (today - pd.Timestamp(r.last_contact).date()).days
            rows.append({"tipo": "Contato antigo (sem inferir falta de atividade externa)", "cliente": r.client,
                         "operadora": r.operator, "serviço": r.service, "status": r.status,
                         "último_contato": r.last_contact, "fonte": r.source,
                         "detalhe": f"Último contato há {age} dias"})
    for f in files:
        if f.get("error"):
            rows.append({"tipo": "Falha de leitura; dados anteriores preservados", "cliente": "", "operadora": "",
                         "serviço": "", "status": "", "último_contato": None, "fonte": f.get("name", ""),
                         "detalhe": f["error"]})
        for warning in f.get("warnings", []):
            rows.append({"tipo": "Estrutura incompleta", "cliente": "", "operadora": "", "serviço": "",
                         "status": "", "último_contato": None, "fonte": f.get("name", ""), "detalhe": warning})
    return pd.DataFrame(rows)


def pending_records(records: pd.DataFrame, days: int, today: date | None = None, red_days: int = 10) -> pd.DataFrame:
    today = today or date.today()
    if records.empty:
        return records.copy()
    out = records[records.status.map(is_active)].copy()
    out["pending_class"] = [calculate_pending(mark, dt, today, days, red_days)
                            for mark, dt in zip(out.marker_status, out.last_contact)]
    return out[out.pending_class.isin({"META PENDENTE — LARANJA", "META PENDENTE — VERMELHA"})]


def indicator_name_records(indicators: dict[str, pd.DataFrame]) -> pd.DataFrame:
    parts = []
    for name, part in indicators.items():
        p = part.copy()
        p["indicator"] = name
        parts.append(p)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def assert_date_boundary() -> None:
    ref = date(2026, 1, 8)
    assert calculate_pending("Sim", ref - timedelta(days=7), ref, 7, 10) == "Dentro do prazo"
    assert calculate_pending("Sim", ref - timedelta(days=8), ref, 7, 10) == "META PENDENTE — LARANJA"
    assert calculate_pending("Revisar", ref - timedelta(days=10), ref, 7, 10) == "META PENDENTE — LARANJA"
    assert calculate_pending("Não", ref - timedelta(days=11), ref, 7, 10) == "META PENDENTE — VERMELHA"
    assert calculate_pending("Sim", None, ref, 7, 10) == "META PENDENTE — DATA NÃO INFORMADA"


if __name__ == "__main__":
    assert_date_boundary()
    print("Regras básicas validadas.")
