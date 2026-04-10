import streamlit as st

st.set_page_config(
    page_title="Central de Apps Streamlit",
    page_icon="📊",
    layout="wide"
)


# IMAGEM NA SIDEBAR
st.logo("images/agco.jpg", size="large")


st.sidebar.divider()


st.title("📊 Ferramentas para Otimização de trabalho - MPS")

st.markdown(
    "<p style='font-size:20px; color:#9ca3af;'>Selecione um aplicativo para começar</p>",
    unsafe_allow_html=True
)

st.divider()


st.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento sem filas/ Período de forecast")
st.page_link("pages/2_NIvelar_com_Filas.py", label="🛠 Nivelamento utilizando Filas")
st.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparativo Product Request VS Operational Plan")
