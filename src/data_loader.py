from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import logging
import os
import re
import time
import unicodedata

import pandas as pd
from openpyxl import load_workbook

from .rules import SERVICES, norm_text, norm_key

ALIASES = {
    "client_code": ["cod", "codigo cliente", "código cliente"],
    "client": ["cliente", "nome cliente"],
    "insurance_id": ["id", "id operadora"],
    "operator": ["operadora", "nome operadora"],
    "category": ["categoria", "tipo categoria"],
    "priority": ["prioridade", "prioridade comercial"],
    "cred_status": ["status credenciamento", "status cred"],
    "cred_operator": ["operador 1", "responsavel credenciamento", "responsavel pelo credenciamento", "operador credenciamento"],
    "cred_date": ["ult contato 1", "ultimo contato 1", "ult contato credenciamento", "ultimo contato credenciamento", "data ultimo contato credenciamento"],
    "cred_detail": ["detalhe do status de credenciamento para o cliente 1", "detalhe status credenciamento", "detalhe do status credenciamento"],
    "cred_notes": ["observacao 1", "observacoes credenciamento"],
    "reaj_status": ["status reajuste"],
    "reaj_operator": ["operador 2", "responsavel reajuste", "operador reajuste"],
    "reaj_date": ["ult contato 2", "ultimo contato 2", "ult contato reajuste", "ultimo contato reajuste"],
    "contract_date": ["contrato", "data contrato", "data do contrato"],
    "last_adjustment": ["ult reajuste", "ultimo reajuste", "data ultimo reajuste"],
    "reaj_detail": ["detalhe do status de reajuste para o cliente 2", "detalhe status reajuste"],
    "reaj_notes": ["observacao 2", "observacoes reajuste"],
    "ext_status": ["status extensao"],
    "ext_operator": ["operador 3", "responsavel extensao", "operador extensao"],
    "ext_date": ["ult contato 3", "ultimo contato 3", "ult contato extensao", "ultimo contato extensao"],
    "ext_type": ["tipo de extensao", "tipo extensao"],
    "ext_type_detail": ["detalhe tipo de extensao"],
    "ext_detail": ["detalhe do status para o cliente 3", "detalhe status extensao"],
    "ext_notes": ["observacao 3", "observacoes extensao"],
    "apx_status": ["status aproximacao"],
    "apx_operator": ["operador 4", "responsavel aproximacao", "operador aproximacao"],
    "apx_date": ["ult contato 4", "ultimo contato 4", "ult contato aproximacao", "ultimo contato aproximacao"],
    "apx_detail": ["detalhe do status de aproximacao para o cliente 4", "detalhe status aproximacao"],
    "apx_notes": ["observacao 4", "observacoes aproximacao"],
}
SERVICE_FIELDS = {field for values in SERVICES.values() for field in values}
REQUIRED_FIELDS = {"client", "operator", "category", "priority"}


def _normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _excel_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            dt = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            return None if pd.isna(dt) else dt.date()
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    for dayfirst in (True, False):
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if not pd.isna(parsed):
            return parsed.date()
    return None


def _find_header(ws):
    wanted = {_normalize_header(x) for vals in ALIASES.values() for x in vals}
    best = (0, None, None)
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 20), values_only=True), start=1):
        vals = [_normalize_header(v) for v in row]
        score = sum(1 for v in vals if v in wanted)
        if score > best[0]:
            best = (score, row_idx, vals)
    return best[1], best[2], best[0]


def _map_headers(normalized_headers):
    mapped = {}
    for field, options in ALIASES.items():
        option_keys = {_normalize_header(x) for x in options}
        for idx, header in enumerate(normalized_headers):
            if header in option_keys:
                mapped[field] = idx
                break
    # Fallbacks for legacy columns with a service name in a header suffix.
    fallback_terms = {
        "cred_date": ("ult contato", "ultimo contato"),
        "reaj_date": ("ult contato", "ultimo contato"),
        "ext_date": ("ult contato", "ultimo contato"),
        "apx_date": ("ult contato", "ultimo contato"),
    }
    for field, terms in fallback_terms.items():
        if field in mapped:
            continue
        ordinal = {"cred_date": "1", "reaj_date": "2", "ext_date": "3", "apx_date": "4"}[field]
        for idx, h in enumerate(normalized_headers):
            if any(term in h for term in terms) and (ordinal in h or (field == "cred_date" and h in terms)):
                mapped[field] = idx
                break
    return mapped


def read_matrix(path: Path) -> dict:
    """Reads only the FLUXO OPERACIONAL sheet and never writes to a source file."""
    result = {"records": [], "warnings": [], "rows": 0, "marker_mode": "unconfirmed"}
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["FLUXO OPERACIONAL"] if "FLUXO OPERACIONAL" in wb.sheetnames else wb[wb.sheetnames[0]]
        header_row, headers, score = _find_header(ws)
        if header_row is None or score < 4:
            raise ValueError("Não foi possível localizar os cabeçalhos da matriz (linha 1–20).")
        columns = _map_headers(headers)
        absent = sorted(REQUIRED_FIELDS - columns.keys())
        service_statuses = [f for f in SERVICES.values() if f[0] in columns]
        if absent:
            result["warnings"].append("Campos básicos ausentes: " + ", ".join(absent))
        if not service_statuses:
            raise ValueError("Nenhuma coluna de STATUS dos serviços foi localizada.")
        for service, fields in SERVICES.items():
            missing = [field for field in fields[:3] if field not in columns]
            if missing:
                result["warnings"].append(f"{service}: colunas ausentes: " + ", ".join(missing))
        marker_columns = [i for i, h in enumerate(headers) if any(t in h for t in ("bolinha verde", "marcacao verde", "marcador verde", "indicador verde"))]
        if not marker_columns:
            result["warnings"].append("A planilha não apresenta uma coluna dedicada de marcação verde; a formatação condicional de status não confirma a bolinha de meta.")
        identity_counts = {}
        for row_number, cells in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=False), start=header_row + 1):
            values = [c.value for c in cells]
            if not any(v not in (None, "") for v in values):
                continue
            result["rows"] += 1
            def get(field, default=None):
                idx = columns.get(field)
                return values[idx] if idx is not None and idx < len(values) else default
            client = str(get("client", "") or "").strip()
            operator = str(get("operator", "") or "").strip()
            if not client and not operator:
                continue
            marker = "Revisar"
            if marker_columns:
                cell = cells[marker_columns[0]]
                value = str(cell.value or "").strip().casefold()
                color = getattr(cell.fill.fgColor, "rgb", None)
                if value in {"sim", "yes", "true", "1", "x"}:
                    marker = "Sim"
                elif value in {"nao", "não", "no", "false", "0"}:
                    marker = "Não"
                elif isinstance(color, str) and color[-6:].upper() in {"00B050", "008000", "00FF00", "92D050", "70AD47"}:
                    marker = "Sim"
            base = {
                "client_code": str(get("client_code", "") or "").strip(),
                "client": client,
                "insurance_id": str(get("insurance_id", "") or "").strip(),
                "operator": operator,
                "category": str(get("category", "") or "").strip(),
                "priority": str(get("priority", "") or "").strip(),
                "source": path.name,
                "source_row": row_number,
                "contract_date": _excel_date(get("contract_date")),
                "last_adjustment": _excel_date(get("last_adjustment")),
            }
            for service, (status_field, operator_field, date_field, detail_field, notes_field) in SERVICES.items():
                extra = {"EXTENSÃO": ["ext_type", "ext_type_detail"], "REAJUSTE": [], "CREDENCIAMENTO": [], "APROXIMAÇÃO": []}.get(service, [])
                status = get(status_field)
                serv_operator = get(operator_field)
                contact = _excel_date(get(date_field))
                detail = get(detail_field)
                notes = get(notes_field)
                extra_vals = {k: get(k) for k in extra}
                if all(v in (None, "") for v in [status, serv_operator, contact, detail, notes, *extra_vals.values()]):
                    continue
                record = dict(base)
                record.update({"service": service, "status": str(status).strip() if status not in (None, "") else "",
                               "service_operator_source": str(serv_operator).strip() if serv_operator not in (None, "") else "",
                               "last_contact": contact, "detail": str(detail).strip() if detail not in (None, "") else "",
                               "notes": str(notes).strip() if notes not in (None, "") else "",
                               "marker_status": marker, "marker_evidence": "Coluna dedicada" if marker_columns else "Não confirmada"})
                for k, v in extra_vals.items():
                    record[k] = "" if v is None else str(v).strip()
                identity_parts = [base["client_code"], base["insurance_id"], base["operator"], base["category"], base["priority"], service]
                business_key = "|".join(norm_key(x) for x in identity_parts)
                if not business_key.replace("|", ""):
                    business_key = f"ROW{row_number}|{service}"
                occurrence = identity_counts.get(business_key, 0) + 1
                identity_counts[business_key] = occurrence
                record["record_id"] = f"{path.name}::{business_key}::{occurrence}"
                result["records"].append(record)
    finally:
        wb.close()
    return result


def read_responsibles(path: Path) -> tuple[pd.DataFrame, list[str]]:
    warnings = []
    if not path.exists():
        raise FileNotFoundError(f"Cadastro oficial não encontrado: {path}")
    with pd.ExcelFile(path, engine="openpyxl") as book:
        sheet = next((s for s in book.sheet_names if "CLIENTES" in _normalize_header(s).upper() or "RESPONSAVEIS" in _normalize_header(s).upper()), book.sheet_names[0])
        df = book.parse(sheet_name=sheet, dtype=str).fillna("")
    df.columns = [_normalize_header(c) for c in df.columns]
    mapping = {}
    for target, possible in {
        "client": ["cliente", "nome cliente"],
        "responsible": ["responsavel pela gestao comercial", "responsavel", "gestor"],
        "priority": ["prioridade", "escopo prioridade"],
    }.items():
        for col in df.columns:
            if col in {_normalize_header(p) for p in possible}:
                mapping[target] = col
                break
    if not {"client", "responsible", "priority"}.issubset(mapping):
        raise ValueError("Não encontrei as colunas Cliente, Responsável e Prioridade no cadastro oficial.")
    out = df[[mapping["client"], mapping["responsible"], mapping["priority"]]].copy()
    out.columns = ["client", "responsible", "priority"]
    out = out[(out.client.str.strip() != "") & (out.responsible.str.strip() != "")]
    return out.reset_index(drop=True), warnings


_IGNORED_LOGGED = set()

def list_matrix_files(root: Path) -> list[Path]:
    if not root.exists() or not root.is_dir():
        return []
    found = []
    def walk_error(error):
        logging.getLogger("produtividade").warning("Falha ao percorrer uma pasta de matrizes (%s).", type(error).__name__)
    for current, dirs, files in os.walk(root, topdown=True, onerror=walk_error):
        current_path = Path(current)
        kept_dirs = []
        for directory in dirs:
            if norm_key(directory) == norm_key("PASTA SUSPENSA"):
                continue
            else:
                kept_dirs.append(directory)
        dirs[:] = kept_dirs
        for name in files:
            path = current_path / name
            suffix = path.suffix.lower()
            if suffix not in {".xlsm", ".xlsx", ".xls"}:
                continue
            if name.startswith("~$") or suffix != ".xlsm":
                continue
            found.append(path)
    return sorted(found, key=lambda p: str(p).casefold())


def stable_fingerprint(path: Path, settle_seconds: float = 0.35, retries: int = 3):
    for _ in range(retries):
        try:
            a = path.stat()
            time.sleep(settle_seconds)
            b = path.stat()
            if (a.st_size, a.st_mtime_ns) == (b.st_size, b.st_mtime_ns) and b.st_size > 0:
                return (b.st_size, b.st_mtime_ns)
        except OSError:
            time.sleep(settle_seconds)
    raise OSError("Arquivo ocupado, incompleto ou sendo sincronizado; nova tentativa no próximo ciclo.")


def de_duplicate(records: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if records.empty:
        return records, 0
    keys = ["client_code", "insurance_id", "operator", "service", "priority", "status", "last_contact"]
    present = [c for c in keys if c in records.columns]
    duplicated = records.duplicated(subset=present, keep="first")
    count = int(duplicated.sum())
    # Only remove exact matches; preserve conflicting records for human review.
    return records.loc[~duplicated].reset_index(drop=True), count


def read_aliases(path: Path) -> pd.DataFrame:
    columns = ["apelido_cliente", "cliente_cadastro"]
    try:
        return pd.read_csv(path, dtype=str).reindex(columns=columns).fillna("")
    except FileNotFoundError:
        return pd.DataFrame(columns=columns)


def read_manual_csv(path: Path, columns: list[str]) -> pd.DataFrame:
    try:
        return pd.read_csv(path, dtype=str).reindex(columns=columns).fillna("")
    except FileNotFoundError:
        return pd.DataFrame(columns=columns)
