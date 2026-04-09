import streamlit as st

st.set_page_config(
    page_title="Central de Apps Streamlit",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Central de Aplicações")
st.caption("Selecione um aplicativo para começar")

st.divider()

st.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento")
st.page_link("pages/2_Ajuste_Filas.py", label="🛠 Ajuste de Filas")
st.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparação PR vs Plan")
st.page_link("pages/4_Nivelamento_Diario.py", label="📅 Nivelamento Diário")
