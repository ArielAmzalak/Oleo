
# ────────────────────────────────────────────────────────────────────────────────
# streamlit_app.py — aplicativo Streamlit
# ────────────────────────────────────────────────────────────────────────────────
from typing import Dict

import streamlit as st
from utils import build_form_and_get_responses, save_to_sheets, generate_pdf

st.set_page_config(
    page_title="Registro de Coleta de Óleo",
    page_icon="🧾",
    layout="centered",
)
st.title("Oliveira Energia — Registro de coleta de amostra de óleo")

if "pdf_bytes" not in st.session_state:
    st.session_state["pdf_bytes"] = None

responses: Dict[str, object] = build_form_and_get_responses()

if st.button("✅ Enviar & Gerar PDF"):
    if not responses.get("n.º da Amostra"):
        st.error("⚠️ Preencha o campo *n.º da Amostra* (obrigatório).")
    else:
        with st.spinner("Salvando no Google Sheets..."):
            try:
                save_to_sheets(responses)
                st.success("📊 Dados gravados (A..AG) e O.S. atualizado em AH.")
            except Exception as exc:
                st.error(str(exc))
                st.stop()
        with st.spinner("Gerando PDF..."):
            st.session_state["pdf_bytes"] = generate_pdf(responses)
        st.info("✅ PDF gerado — utilize o botão abaixo para baixar.")

if st.session_state["pdf_bytes"]:
    st.download_button(
        label="⬇️ Baixar PDF",
        data=st.session_state["pdf_bytes"],
        file_name=f"amostra_{responses.get('n.º da Amostra', 'sem_numero')}.pdf",
        mime="application/pdf",
    )
