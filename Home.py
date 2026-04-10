import streamlit as st

st.set_page_config(
    page_title="Central de Apps Streamlit",
    page_icon="📊",
    layout="WIDE"
)

st.title("📊 Ferramentas para Otimizaçãoo de trabalho - MPS")
st.caption("Selecione um aplicativo para começar")

st.divider()

st.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento forecast w/o")
st.page_link("pages/2_NIvelar_com_Filas.py", label="🛠 Nivelar com Filas")
st.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparação PR vs Plan")

