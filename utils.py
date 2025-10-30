
# ────────────────────────────────────────────────────────────────────────────────
# utils.py — funções auxiliares (Google Sheets, PDF e UI Streamlit)
# ────────────────────────────────────────────────────────────────────────────────
from __future__ import annotations

import io
import os
import re
import json
from math import ceil
from datetime import datetime
from typing import Dict, List, Tuple, Any, Optional

import qrcode
from fpdf import FPDF
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from barcode import Code128
from barcode.writer import ImageWriter

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None  # permite importar utils.py sem Streamlit

# ░░░ Config Google Sheets ░░░
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = "1VLDQUCO3Aw4ClAvhjkUsnBxG44BTjz-MjHK04OqPxYM"
SHEET_NAME = "Geral"

# ░░░ Rótulo e coluna de destino do novo campo O.S. ░░░
OS_FORM_LABEL = "Ordem de Serviço (O.S.)"
OS_TARGET_COL = "AH"  # coluna onde gravaremos a O.S. após o append A..AG

# ░░░ Cabeçalhos EXATOS da planilha (A..AG), sem a coluna OS (AH) ░░░
SHEET_HEADERS_EXCL_OS: List[str] = [
    "Estado de Origem",
    "Cliente",
    "Data da coleta",
    "Local de operação",
    "UGD",
    "Responsável Pela Coleta",
    "n.º da Amostra",
    "n.º de série Equipamento",
    "Frota",
    "Horímetro do Óleo",
    "Houve troca de óleo após coleta?",
    "Troca de Filtro após coleta",
    "Houve mudança do local de operação?",
    "Fabricante",
    "Modelo",
    "Horímetro do Motor",
    "Houve complemento de óleo",
    "Se sim, quantos litros",
    "Amostra coletada",
    "Fabricante do Óleo",
    "Grau de viscosidade",
    "Nome",
    "Apresentou limalha no filtro ou na tela?",
    "Apresentou limalhas no bujão magnético?",
    "Equipamento apresentou ruído anormal?",
    "Existem vazamentos no sistema",
    "A temperatura de operação está normal?",
    "O desempenho do sistema está normal?",
    "Detalhes das anormalidades (caso Haja)",
    "Pessoa de contato",
    "Telefone",
    "Status",        # AF — vazio nesta etapa
    "Data Status"    # AG — vazio nesta etapa
]

# ░░░ Mapeia cada cabeçalho -> label do formulário (quando diferir) ░░░
SHEET_HEADER_TO_FORM: Dict[str, str] = {
    "Estado de Origem": "Estado de Origem",
    "Cliente": "Cliente",
    "Data da coleta": "Data da coleta",
    "Local de operação": "Local de operação:",
    "UGD": "UGD:",
    "Responsável Pela Coleta": "Responsável Pela Coleta:",
    "n.º da Amostra": "n.º da Amostra",
    "n.º de série Equipamento": "n.º de série:",
    "Frota": "Frota:",
    "Horímetro do Óleo": "Horímetro do Óleo:",
    "Houve troca de óleo após coleta?": "Houve troca de óleo após coleta?",
    "Troca de Filtro após coleta": "Trocado o filtro após coleta?",
    "Houve mudança do local de operação?": "Houve mudança do local de operação?",
    "Fabricante": "Fabricante do Equipamento:",
    "Modelo": "Modelo:",
    "Horímetro do Motor": "Horímetro do Motor",
    "Houve complemento de óleo": "Houve complemento de óleo?",
    "Se sim, quantos litros": "Se sim, quantos litros?",
    "Amostra coletada": "Amostra coletada:",
    "Fabricante do Óleo": "Fabricante:",  # seção Óleo
    "Grau de viscosidade": "Grau de viscosidade:",
    "Nome": "Nome:",
    "Apresentou limalha no filtro ou na tela?": "Apresentou limalha no filtro ou na tela?",
    "Apresentou limalhas no bujão magnético?": "Apresentou limalhas no bujão magnético?",
    "Equipamento apresentou ruído anormal?": "Equipamento apresentou ruído anormal?",
    "Existem vazamentos no sistema": "Existem vazamentos no sistema?",
    "A temperatura de operação está normal?": "A temperatura de operação está normal?",
    "O desempenho do sistema está normal?": "O desempenho do sistema está normal?",
    "Detalhes das anormalidades (caso Haja)": "Detalhes das anormalidades (caso Haja):",
    "Pessoa de contato": "Pessoa de contato:",
    "Telefone": "Telefone:",
    # "Status" e "Data Status" são vazios nesta etapa
}

# ░░░ Autenticação Google ░░░
def _authorize_google_sheets() -> Credentials:
    # Preferencialmente usa token.json quando executado localmente
    token_path = "token.json"
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Busca client secret do Streamlit Secrets ou variável de ambiente
            client_secret_json = None
            if st:
                client_secret_json = st.secrets.get("GOOGLE_CLIENT_SECRET")
            if not client_secret_json:
                client_secret_json = os.getenv("GOOGLE_CLIENT_SECRET")
            if not client_secret_json:
                raise RuntimeError("Credenciais Google ausentes. Defina GOOGLE_CLIENT_SECRET nos secrets/ambiente.")
            # Usa InstalledAppFlow apenas se você rodar localmente e quiser abrir consent
            from google_auth_oauthlib.flow import InstalledAppFlow
            try:
                client_config = json.loads(client_secret_json)
            except Exception as exc:
                raise RuntimeError("GOOGLE_CLIENT_SECRET inválido (JSON).") from exc
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            # Nota: em ambientes sem navegador, use run_console()
            if st:
                creds = flow.run_console()
            else:
                creds = flow.run_console()
        # Persiste token.json quando possível
        try:
            with open(token_path, "w", encoding="utf-8") as fp:
                fp.write(creds.to_json())
        except Exception:
            pass
    return creds

def _get_sheets_service():
    return build("sheets", "v4", credentials=_authorize_google_sheets(), cache_discovery=False)

# ░░░ Estrutura do formulário (labels do formulário) ░░░
FORM_SECTIONS: List[Tuple[str, List[Tuple[str, Any]]]] = [
    (
        "Geral",
        [
            ("Estado de Origem", "AM"),
            ("Cliente", "Pie - Oliveira Energia"),
            ("Data da coleta", datetime.today().strftime("%d/%m/%Y")),
            ("Local de operação:", ""),
            ("UGD:", ""),
            ("Responsável Pela Coleta:", ""),
            ("n.º da Amostra", ""),           # obrigatório
            (OS_FORM_LABEL, ""),              # NOVO campo — lado a lado no PDF
        ],
    ),
    (
        "Equipamento",
        [
            ("n.º de série:", ""),
            ("Frota:", ""),
            ("Horímetro do Óleo:", ""),
            ("Houve troca de óleo após coleta?", False),
            ("Trocado o filtro após coleta?", False),
            ("Houve mudança do local de operação?", False),
            ("Fabricante do Equipamento:", "Scania"),
            ("Modelo:", "DC13"),
            ("Horímetro do Motor", ""),
        ],
    ),
    (
        "Óleo",
        [
            ("Houve complemento de óleo?", False),
            ("Se sim, quantos litros?", ""),
            ("Amostra coletada:", "Motor"),
            ("Fabricante:", "Mobil"),                 # Fabricante do Óleo
            ("Grau de viscosidade:", "15W40"),
            ("Nome:", "Mobil Delvac"),
            ("Apresentou limalha no filtro ou na tela?", False),
            ("Apresentou limalhas no bujão magnético?", False),
            ("Equipamento apresentou ruído anormal?", False),
            ("Existem vazamentos no sistema?", False),
            ("A temperatura de operação está normal?", False),
            ("O desempenho do sistema está normal?", False),
            ("Detalhes das anormalidades (caso Haja):", ""),
        ],
    ),
    (
        "Contato",
        [
            ("Pessoa de contato:", "Francisco Sampaio"),
            ("Telefone:", "(92) 99437-6579"),
        ],
    ),
]

BOOL_LABELS = {
    label
    for _, questions in FORM_SECTIONS
    for label, default in questions
    if isinstance(default, bool)
}


def _build_base_defaults() -> Dict[str, Any]:
    defaults: Dict[str, Any] = {}
    for _, questions in FORM_SECTIONS:
        for label, default in questions:
            defaults[label] = default
    return defaults


BASE_FORM_DEFAULTS = _build_base_defaults()

# ░░░ Helpers de estado do formulário ░░░
def _apply_form_values(values: Dict[str, Any]) -> None:
    if st is None:
        return
    form_values = st.session_state.setdefault("form_values", {})
    for label, value in values.items():
        form_values[label] = value
        if isinstance(value, bool):
            key_yes = f"{label}_yes"
            key_no = f"{label}_no"
            st.session_state[key_yes] = bool(value)
            st.session_state[key_no] = not bool(value)
        else:
            st.session_state[label] = "" if value is None else str(value)


def _queue_form_updates(values: Dict[str, Any]) -> None:
    """Armazena atualizações para aplicação segura após o próximo rerun."""
    if st is None:
        return

    form_values = st.session_state.setdefault("form_values", BASE_FORM_DEFAULTS.copy())
    form_values.update(values)

    pending = st.session_state.setdefault("_pending_form_values", {})
    pending.update(values)


def sync_sample_number(sample_number: str) -> None:
    """Atualiza o campo da amostra no estado do formulário e widgets."""
    if st is None:
        return
    _queue_form_updates({"n.º da Amostra": sample_number})


def _reset_form_defaults(keep_sample: Optional[str] = None) -> None:
    defaults = BASE_FORM_DEFAULTS.copy()
    if keep_sample is not None:
        defaults["n.º da Amostra"] = keep_sample
    if st is None:
        return
    st.session_state["form_values"] = defaults
    _queue_form_updates(defaults)


def _ensure_form_state() -> None:
    if st is None:
        return
    if "form_values" not in st.session_state:
        st.session_state["form_values"] = BASE_FORM_DEFAULTS.copy()
        _apply_form_values(st.session_state["form_values"])
    pending_updates = st.session_state.pop("_pending_form_values", None)
    if pending_updates:
        st.session_state["form_values"].update(pending_updates)
        _apply_form_values(pending_updates)
    st.session_state.setdefault("sample_row_index", None)
    st.session_state.setdefault("sample_lookup_status", None)
    st.session_state.setdefault("sample_lookup_message", "")
    st.session_state.setdefault("sample_lookup_warning", None)
    st.session_state.setdefault("sample_existing_extras", {})
    st.session_state.setdefault("sample_last_loaded_number", "")


def _column_letter_to_index(col: str) -> int:
    col = col.strip().upper()
    if not col:
        raise ValueError("Coluna vazia")
    idx = 0
    for ch in col:
        if not ch.isalpha():
            raise ValueError(f"Coluna inválida: {col}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _coerce_sheet_value(label: str, value: Any) -> Any:
    if label in BOOL_LABELS:
        text = "" if value is None else str(value).strip().lower()
        if text in {"sim", "s", "true", "1", "yes"}:
            return True
        if text in {"não", "nao", "n", "false", "0", "no"}:
            return False
        base_default = BASE_FORM_DEFAULTS.get(label)
        return bool(base_default) if isinstance(base_default, bool) else False
    return "" if value is None else str(value)


def _fetch_sample_from_sheets(sample_number: str) -> Optional[Tuple[int, Dict[str, Any], int, Dict[str, str]]]:
    service = _get_sheets_service()
    try:
        result = (
            service.spreadsheets()
            .values()
            .get(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!A:AH",
            )
            .execute()
        )
    except HttpError as exc:
        raise RuntimeError(f"Erro ao consultar planilha: {exc}") from exc

    rows = result.get("values", [])
    if not rows:
        return None

    header = rows[0]
    header_map = {name: idx for idx, name in enumerate(header)}
    sample_col_idx = header_map.get("n.º da Amostra")
    if sample_col_idx is None:
        raise RuntimeError("Cabeçalho 'n.º da Amostra' não encontrado na planilha.")

    os_col_idx = header_map.get(OS_FORM_LABEL)
    if os_col_idx is None:
        os_col_idx = _column_letter_to_index(OS_TARGET_COL)

    matches: List[Tuple[int, List[str]]] = []
    for idx, row in enumerate(rows[1:], start=2):
        cell_value = row[sample_col_idx] if sample_col_idx < len(row) else ""
        if str(cell_value).strip() == sample_number:
            matches.append((idx, row))

    if not matches:
        return None

    row_idx, row_values = matches[-1]
    form_data: Dict[str, Any] = {}
    extras: Dict[str, str] = {}
    for sheet_header, form_label in SHEET_HEADER_TO_FORM.items():
        col_idx = header_map.get(sheet_header)
        if col_idx is None:
            continue
        cell_value = row_values[col_idx] if col_idx < len(row_values) else ""
        form_data[form_label] = _coerce_sheet_value(form_label, cell_value)

    if os_col_idx is not None:
        cell_value = row_values[os_col_idx] if os_col_idx < len(row_values) else ""
        form_data[OS_FORM_LABEL] = "" if cell_value is None else str(cell_value)

    for header_name in ("Status", "Data Status"):
        idx = header_map.get(header_name)
        if idx is not None and idx < len(row_values):
            extras[header_name] = "" if row_values[idx] is None else str(row_values[idx])

    return row_idx, form_data, len(matches), extras


def _handle_sample_change() -> None:
    if st is None:
        return
    raw_value = st.session_state.get("n.º da Amostra", "")
    sample_value = str(raw_value).strip()
    st.session_state["sample_lookup_warning"] = None

    if sample_value:
        try:
            fetched = _fetch_sample_from_sheets(sample_value)
        except Exception as exc:  # noqa: BLE001
            st.session_state["sample_row_index"] = None
            st.session_state["sample_lookup_status"] = "error"
            st.session_state["sample_lookup_message"] = f"Erro ao buscar amostra: {exc}"
            _trigger_rerun()
            return

        if fetched is None:
            _reset_form_defaults(keep_sample=sample_value)
            st.session_state["sample_row_index"] = None
            st.session_state["sample_lookup_status"] = "new"
            st.session_state["sample_lookup_message"] = (
                f"Amostra {sample_value} não encontrada. Preencha os dados para criar um novo registro."
            )
            st.session_state["sample_existing_extras"] = {}
            st.session_state["sample_last_loaded_number"] = sample_value
        else:
            row_idx, form_data, count, extras = fetched
            form_data["n.º da Amostra"] = sample_value
            _queue_form_updates(form_data)
            st.session_state["sample_row_index"] = row_idx
            st.session_state["sample_lookup_status"] = "loaded"
            st.session_state["sample_lookup_message"] = (
                f"Amostra {sample_value} carregada a partir da linha {row_idx}."
            )
            st.session_state["sample_existing_extras"] = extras
            st.session_state["sample_last_loaded_number"] = sample_value
            if count > 1:
                st.session_state["sample_lookup_warning"] = (
                    f"Foram encontradas {count} linhas com este número. A mais recente (linha {row_idx}) foi carregada."
                )
    else:
        _reset_form_defaults(keep_sample="")
        st.session_state["sample_row_index"] = None
        st.session_state["sample_lookup_status"] = None
        st.session_state["sample_lookup_message"] = ""
        st.session_state["sample_existing_extras"] = {}
        st.session_state["sample_last_loaded_number"] = ""

    _trigger_rerun()


def _trigger_rerun() -> None:
    """Solicita um rerun compatível com versões antigas e novas do Streamlit."""
    if st is None:
        return

    for attr_name in ("experimental_rerun", "rerun"):
        rerun = getattr(st, attr_name, None)
        if callable(rerun):
            rerun()
            return


# ░░░ Helpers UI ░░░
def _render_sample_feedback() -> None:
    if st is None:
        return
    message = st.session_state.get("sample_lookup_message", "")
    status = st.session_state.get("sample_lookup_status")
    warning = st.session_state.get("sample_lookup_warning")

    if message:
        if status == "loaded":
            st.success(message)
        elif status == "new":
            st.info(message)
        elif status == "error":
            st.error(message)
        else:
            st.caption(message)

    if warning:
        st.warning(warning)


def _two_checkboxes(label: str, default: bool | None = None) -> bool:
    if st is None:
        raise RuntimeError("Streamlit não instalado – UI indisponível.")
    st.markdown(f"**{label}**")
    key_yes = f"{label}_yes"
    key_no  = f"{label}_no"
    if key_yes not in st.session_state and key_no not in st.session_state:
        if default is True:
            st.session_state[key_yes] = True
            st.session_state[key_no] = False
        elif default is False:
            st.session_state[key_yes] = False
            st.session_state[key_no] = True
        else:
            st.session_state[key_yes] = False
            st.session_state[key_no] = False

    col_yes, col_no = st.columns(2)

    def _sync_yes() -> None:
        if st.session_state[key_yes]:
            st.session_state[key_no] = False

    def _sync_no() -> None:
        if st.session_state[key_no]:
            st.session_state[key_yes] = False

    with col_yes:
        st.checkbox("Sim", key=key_yes, on_change=_sync_yes)
    with col_no:
        st.checkbox("Não", key=key_no, on_change=_sync_no)
    return bool(st.session_state[key_yes])


def build_form_and_get_responses() -> Dict[str, Any]:
    """Desenha o formulário completo e retorna um dicionário label->valor."""
    if st is None:
        raise RuntimeError("Streamlit não instalado – UI indisponível.")

    _ensure_form_state()
    form_values = st.session_state["form_values"]

    st.header("Formulário de Coleta de Amostras de Óleo 🛢️")
    responses: Dict[str, Any] = {}

    for section, questions in FORM_SECTIONS:
        st.subheader(section)

        if section == "Geral":
            sample_label = "n.º da Amostra"
            col_sample, col_os = st.columns(2)

            with col_sample:
                sample_default = form_values.get(sample_label, "")
                sample_value = st.text_input(
                    sample_label,
                    value="" if sample_default is None else str(sample_default),
                    key=sample_label,
                    on_change=_handle_sample_change,
                )
                sample_value = st.session_state.get(sample_label, sample_value)
            responses[sample_label] = sample_value
            form_values[sample_label] = sample_value

            with col_os:
                os_default = form_values.get(OS_FORM_LABEL, "")
                os_value = st.text_input(
                    OS_FORM_LABEL,
                    value="" if os_default is None else str(os_default),
                )
            responses[OS_FORM_LABEL] = os_value
            form_values[OS_FORM_LABEL] = os_value

            _render_sample_feedback()

            for label, default in questions:
                if label in {sample_label, OS_FORM_LABEL}:
                    continue
                effective_default = form_values.get(label, default)
                if isinstance(default, bool):
                    if isinstance(effective_default, bool):
                        default_bool = effective_default
                    else:
                        default_bool = default
                    value = _two_checkboxes(label, default=default_bool)
                else:
                    value = st.text_input(
                        label,
                        value="" if effective_default is None else str(effective_default),
                    )
                responses[label] = value
                form_values[label] = value
            continue

        for label, default in questions:
            effective_default = form_values.get(label, default)
            if isinstance(default, bool):
                if isinstance(effective_default, bool):
                    default_bool = effective_default
                else:
                    default_bool = default
                value = _two_checkboxes(label, default=default_bool)
            else:
                value = st.text_input(
                    label,
                    value="" if effective_default is None else str(effective_default),
                )
            responses[label] = value
            form_values[label] = value

    return responses

# ░░░ Persistência no Google Sheets ░░░
def _fmt(v: Any) -> str:
    if v is True:
        return "Sim"
    if v is False:
        return "Não"
    return "" if v is None else str(v)

def save_to_sheets(
    responses: Dict[str, Any],
    existing_row: Optional[int] = None,
    existing_extras: Optional[Dict[str, str]] = None,
) -> int:
    """
    Persiste os dados no Google Sheets.

    * Quando ``existing_row`` é ``None``: faz APPEND de A..AG e atualiza AH com a O.S.
    * Quando ``existing_row`` é informado: atualiza A..AH na linha indicada, preservando
      colunas não presentes no formulário (Status/Data Status) através de ``existing_extras``.
    Retorna o índice (1-based) da linha gravada/atualizada.
    """

    extras = existing_extras or {}

    row_out: List[str] = []
    for hdr in SHEET_HEADERS_EXCL_OS:
        if hdr in ("Status", "Data Status"):
            row_out.append(extras.get(hdr, ""))
            continue
        form_label = SHEET_HEADER_TO_FORM.get(hdr)
        val = responses.get(form_label, "") if form_label else ""
        row_out.append(_fmt(val))

    os_value = _fmt(responses.get(OS_FORM_LABEL, ""))

    try:
        service = _get_sheets_service()

        if existing_row is not None:
            row_idx_int = int(existing_row)
            row_full = list(row_out)
            row_full.append(os_value)
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!A{row_idx_int}:AH{row_idx_int}",
                valueInputOption="RAW",
                body={"values": [row_full]},
            ).execute()
            return row_idx_int

        body = {"values": [row_out]}
        append_result = service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body,
        ).execute()

        updated_range = (append_result or {}).get("updates", {}).get("updatedRange", "")
        m = re.search(r"!.*?(\d+):", updated_range) or re.search(r"!.*?(\d+)$", updated_range)
        if not m:
            raise RuntimeError(f"Não foi possível detectar a linha inserida: {updated_range}")
        row_idx_int = int(m.group(1))

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!{OS_TARGET_COL}{row_idx_int}",
            valueInputOption="RAW",
            body={"values": [[os_value]]},
        ).execute()

        return row_idx_int

    except HttpError as exc:
        if st:
            st.error("❌ Erro ao gravar no Google Sheets.")
        raise RuntimeError(f"Erro ao gravar → {exc}") from exc

# ░░░ PDF ░░░
_REPL = {
    "\u2013": "-",
    "\u2014": "-",
    "\u2011": "-",
    "\u00A0": " ",
    "\n": " ",
    "\r": " ",
}
def _safe(txt: object) -> str:
    if txt is None:
        return ""
    if not isinstance(txt, str):
        txt = str(txt)
    for bad, good in _REPL.items():
        txt = txt.replace(bad, good)
    return txt.encode("latin-1", "replace").decode("latin-1")

def generate_pdf(responses: Dict[str, Any]) -> bytes:
    """
    Gera um PDF A4 (retrato). O campo O.S. está logo após 'n.º da Amostra' no bloco 'Geral',
    então os dois caem lado a lado (duas colunas) sem criar linha extra.
    """
    sample_no = str(responses.get("n.º da Amostra", "SEM_NUMERO")).strip() or "SEM_NUMERO"

    # QR em memória
    qr_img = qrcode.make(sample_no)
    buf_qr = io.BytesIO()
    qr_img.save(buf_qr, format="PNG")
    buf_qr.seek(0)

    # Código de barras em memória
    buf_bar = io.BytesIO()
    barcode = Code128(sample_no, writer=ImageWriter())
    barcode.write(buf_bar, options={
        "module_width": 0.3,
        "module_height": 15,
        "font_size": 8,
    })
    buf_bar.seek(0)

    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.set_left_margin(10)
    pdf.set_top_margin(10)
    pdf.set_right_margin(10)
    pdf.add_page()

    # Cabeçalho
    qr_w = 25
    bar_w = 30
    y_start = pdf.get_y()
    x_qr  = pdf.l_margin
    x_bar = pdf.w - pdf.r_margin - bar_w
    pdf.image(buf_qr, x=x_qr, y=y_start, w=qr_w)
    pdf.image(buf_bar, x=x_bar, y=y_start + 5, w=bar_w)
    pdf.set_font("Helvetica", size=16)
    pdf.set_y(y_start + 8)
    pdf.set_x(0)
    pdf.cell(w=0, h=10, txt="Oliveira Energia - Amostra de óleo", align="C", ln=True)
    pdf.ln(8)

    # Corpo (duas perguntas por linha)
    inner_width = pdf.w - pdf.l_margin - pdf.r_margin
    LABEL_RATIO = 0.655
    group_width = inner_width / 2
    label_w  = group_width * LABEL_RATIO
    value_w  = group_width - label_w
    row_h    = 4.5

    pdf.set_font("Helvetica", size=7)

    for section, qs in FORM_SECTIONS:
        pdf.set_font("Helvetica", style="B", size=11)
        pdf.set_fill_color(240)
        pdf.cell(0, 7, _safe(section), ln=True, border=1, fill=True)
        pdf.set_font("Helvetica", size=9)

        pairs = []
        for label, _ in qs:
            val = responses.get(label, "")
            if val is True:
                val = "Sim"
            elif val is False:
                val = "Não"
            pairs.append((_safe(label), _safe(str(val))))

        n_rows = ceil(len(pairs) / 2)
        idx = 0
        for _ in range(n_rows):
            # esquerda
            if idx < len(pairs):
                lab, val = pairs[idx]
                pdf.cell(label_w, row_h, lab, border=1)
                pdf.cell(value_w, row_h, val, border=1)
                idx += 1
            else:
                pdf.cell(label_w, row_h, "", border=1)
                pdf.cell(value_w, row_h, "", border=1)
            # direita
            if idx < len(pairs):
                lab, val = pairs[idx]
                pdf.cell(label_w, row_h, lab, border=1)
                pdf.cell(value_w, row_h, val, border=1)
                idx += 1
            else:
                pdf.cell(label_w, row_h, "", border=1)
                pdf.cell(value_w, row_h, "", border=1)
            pdf.ln(row_h)
        pdf.ln(1)

    raw = pdf.output(dest="S")
    return bytes(raw) if isinstance(raw, (bytes, bytearray)) else str(raw).encode("latin-1")

if __name__ == "__main__":
    if st is None:
        raise SystemExit("Execute via `streamlit run streamlit_app.py`.")

