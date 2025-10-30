
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
    sample_no = str(responses.get("n.º da Amostra", "") or "").strip()
    if not sample_no:
        st.error("⚠️ Preencha o campo *n.º da Amostra* (obrigatório).")
    else:
        responses["n.º da Amostra"] = sample_no
        st.session_state["n.º da Amostra"] = sample_no
        if "form_values" in st.session_state:
            st.session_state["form_values"]["n.º da Amostra"] = sample_no

        last_loaded = st.session_state.get("sample_last_loaded_number", "") or ""
        existing_row = st.session_state.get("sample_row_index")
        existing_extras = dict(st.session_state.get("sample_existing_extras", {}))
        if sample_no != last_loaded:
            existing_row = None
            existing_extras = {}
            st.session_state["sample_row_index"] = None
            st.session_state["sample_existing_extras"] = {}

        with st.spinner("Salvando no Google Sheets..."):
            try:
                row_idx = save_to_sheets(
                    responses,
                    existing_row=existing_row,
                    existing_extras=existing_extras,
                )
                st.session_state["sample_row_index"] = row_idx
                st.session_state["sample_last_loaded_number"] = sample_no
                st.session_state["sample_existing_extras"] = existing_extras
                st.session_state["sample_lookup_status"] = "loaded"
                st.session_state["sample_lookup_message"] = (
                    f"Amostra {sample_no} sincronizada na linha {row_idx}."
                )
                if existing_row is not None:
                    st.success(f"♻️ Registro atualizado na linha {row_idx} (A..AH).")
                else:
                    st.success(f"📊 Dados gravados na linha {row_idx} (A..AH).")
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
