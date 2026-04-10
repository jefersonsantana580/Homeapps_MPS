import streamlit as st

st.set_page_config(
    page_title="Central de Apps Streamlit",
    page_icon="📊",
    layout="wide"
)


# IMAGEM NA SIDEBAR
st.sidebar.image(
    "images/logo.png",
    use_column_width=True
)

st.title("📊 Ferramentas para Otimizaçãoo de trabalho - MPS")
st.caption("Selecione um aplicativo para começar")

st.divider()

st.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento sem filas/ Período de forecast")
st.page_link("pages/2_NIvelar_com_Filas.py", label="🛠 Nivelar utliziando Filas")
st.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparação PR vs Plan")

