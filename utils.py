# ────────────────────────────────────────────────────────────────────────────────
# utils.py – funções auxiliares (Google Sheets, PDF e UI Streamlit)
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
from fpdf.errors import FPDFException
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

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
OS_LABEL = "Ordem de Serviço (O.S.):"  # <- rótulo padrão do novo campo
OS_TARGET_COL = "AH"                   # <- coluna de destino no Google Sheets

# ░░░ Autenticação ░░░
@st.cache_resource
def _authorize_google_sheets() -> Credentials:
    from google_auth_oauthlib.flow import InstalledAppFlow
    token_path = "token.json"
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            except Exception:
                st.error("❌ Não foi possível carregar as credenciais do Google.")
                st.stop()
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()
        try:
            with open(token_path, "w", encoding="utf-8") as fp:
                fp.write(creds.to_json())
        except Exception:
            pass
    return creds

@st.cache_resource
def _get_sheets_service():
    return build("sheets", "v4", credentials=_authorize_google_sheets(), cache_discovery=False)

# ░░░ Estrutura do formulário ░░░
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
            ("n.º da Amostra", ""),        # obrigatório
            (OS_LABEL, ""),                # <- NOVO: vem logo depois, para ficar na mesma linha no PDF
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
            ("Fabricante:", "Mobil"),
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

SHEET_COLUMNS: List[str] = [label for _, block in FORM_SECTIONS for (label, _) in block]

# ░░░ UI: duas caixas (Sim/Não) para booleanos ░░░
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
    if st is None:
        raise RuntimeError("Streamlit não instalado – UI indisponível.")
    st.header("Formulário de Coleta de Amostras de Óleo 🛢️")
    responses: Dict[str, Any] = {}
    for section, questions in FORM_SECTIONS:
        st.subheader(section)
        for label, default in questions:
            if isinstance(default, bool):
                responses[label] = _two_checkboxes(label, default=default)
            else:
                responses[label] = st.text_input(label, value=str(default))
    return responses

# ░░░ Persistência no Google Sheets ░░░
def save_to_sheets(responses: Dict[str, Any]) -> None:
    """Append da linha principal + update da coluna AH com o O.S."""
    # prepara a linha “normal”, mas deixa o campo OS vazio nesta linha
    row = []
    os_value_raw = responses.get(OS_LABEL, "")
    for col in SHEET_COLUMNS:
        if col == OS_LABEL:
            row.append("")  # ← deixamos vazio no append para não bagunçar o layout
            continue
        val = responses.get(col, "")
        row.append("Sim" if val is True else "Não" if val is False else str(val))

    body = {"values": [row]}
    try:
        service = _get_sheets_service()
        append_result = service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body=body,
        ).execute()

        # pega a linha recém-escrita (ex.: "Geral!A123:AF123")
        updated_range = (append_result or {}).get("updates", {}).get("updatedRange", "")
        m = re.search(r"!.*?(\d+):", updated_range)
        if not m:
            # fallback: tenta o final do range
            m = re.search(r"!.*?(\d+)$", updated_range)
        if not m:
            raise RuntimeError(f"Não foi possível detectar a linha inserida: {updated_range}")
        row_idx = m.group(1)

        # grava APENAS o O.S. na coluna AH da linha correspondente
        os_value = "Sim" if os_value_raw is True else "Não" if os_value_raw is False else str(os_value_raw)
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

# ░░░ Sanitização de texto ░░░
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

# ░░░ PDF (1 página, 2 perguntas por linha) ░░░
def generate_pdf(responses: Dict[str, Any]) -> bytes:
    """
    Gera um PDF A4 (retrato). Como o campo O.S. está logo depois de 'n.º da Amostra'
    no bloco 'Geral', os dois caem lado a lado na MESMA LINHA (duas colunas).
    """
    sample_no = str(responses.get("n.º da Amostra", "SEM_NUMERO")).strip() or "SEM_NUMERO"

    # QR em memória
    qr_img = qrcode.make(sample_no)
    buf_qr = io.BytesIO()
    qr_img.save(buf_qr, format="PNG")
    buf_qr.seek(0)

    # Código de barras
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
    header_h = 20
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
            for _in_group in range(2):
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

# Execução direta (demo)
if __name__ == "__main__":
    if st is None:
        raise SystemExit("Execute via `streamlit run utils.py`.")
    st.title("Coleta de Óleo – Demo Utilitário")
    resps = build_form_and_get_responses()
    if st.button("Salvar e Gerar PDF"):
        if not resps.get("n.º da Amostra"):
            st.error("Por favor, preencha o número da amostra!")
        else:
            save_to_sheets(resps)
            pdf_bytes = generate_pdf(resps)
            st.success("Dados salvos! Faça o download do PDF abaixo.")
            st.download_button(
                "Baixar PDF",
                data=pdf_bytes,
                file_name=f"amostra_{resps.get('n.º da Amostra')}.pdf",
                mime="application/pdf",
            )
