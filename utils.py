
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
from typing import Dict, List, Tuple, Any

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

# ░░░ Helpers UI ░░░
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
    def _sync_yes():
        if st.session_state[key_yes]:
            st.session_state[key_no] = False
    def _sync_no():
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
    st.header("Formulário de Coleta de Amostras de Óleo 🛢️")
    responses: Dict[str, Any] = {}

    for section, questions in FORM_SECTIONS:
        st.subheader(section)

        # Layout especial: em "Geral", alinhar 'n.º da Amostra' e 'O.S.' lado a lado
        if section == "Geral":
            # desenha os 6 primeiros normalmente
            for label, default in questions[:6]:
                if isinstance(default, bool):
                    responses[label] = _two_checkboxes(label, default=default)
                else:
                    responses[label] = st.text_input(label, value=str(default))

            # agora 'n.º da Amostra' e 'O.S.' em duas colunas
            col1, col2 = st.columns(2)
            with col1:
                label, default = questions[6]
                responses[label] = st.text_input(label, value=str(default))
            with col2:
                label, default = questions[7]
                responses[label] = st.text_input(label, value=str(default))
            continue

        # demais seções em lista simples
        for label, default in questions:
            if isinstance(default, bool):
                responses[label] = _two_checkboxes(label, default=default)
            else:
                responses[label] = st.text_input(label, value=str(default))

    return responses

# ░░░ Persistência no Google Sheets ░░░
def _fmt(v: Any) -> str:
    if v is True:
        return "Sim"
    if v is False:
        return "Não"
    return "" if v is None else str(v)

def save_to_sheets(responses: Dict[str, Any]) -> None:
    """
    1) Faz APPEND somente A..AG (sem OS), alinhado ao cabeçalho exato da planilha.
    2) Em seguida faz UPDATE apenas de AH{linha} com o valor de OS do formulário.
    """
    # Monta a linha exatamente na ordem A..AG
    row_out: List[str] = []
    for hdr in SHEET_HEADERS_EXCL_OS:
        if hdr in ("Status", "Data Status"):
            row_out.append("")  # AF e AG ficam vazios nesta etapa
            continue
        form_label = SHEET_HEADER_TO_FORM.get(hdr)
        val = responses.get(form_label, "") if form_label else ""
        row_out.append(_fmt(val))

    body = {"values": [row_out]}

    try:
        service = _get_sheets_service()
        append_result = service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",           # importante: NÃO inclui AH aqui
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body,
        ).execute()

        # Ex.: "Geral!A123:AG123" → extrai o 123
        updated_range = (append_result or {}).get("updates", {}).get("updatedRange", "")
        m = re.search(r"!.*?(\d+):", updated_range) or re.search(r"!.*?(\d+)$", updated_range)
        if not m:
            raise RuntimeError(f"Não foi possível detectar a linha inserida: {updated_range}")
        row_idx = m.group(1)

        # Atualiza AH (OS)
        os_value = _fmt(responses.get(OS_FORM_LABEL, ""))
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!{OS_TARGET_COL}{row_idx}",
            valueInputOption="RAW",
            body={"values": [[os_value]]},
        ).execute()

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

