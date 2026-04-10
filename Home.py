
import streamlit as st

st.set_page_config(
    page_title="Ferramentas para Otimização de trabalho - MPS",
    page_icon="📊",
    layout="wide"
)

# ESCONDE MENU PADRÃO
st.markdown(
    """
    <style>
        [data-testid="stSidebarNav"] {
            display: none;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# SIDEBAR PERSONALIZADA
st.sidebar.image("images/agco.jpg", width=240)
st.sidebar.divider()

st.sidebar.markdown("### 📊 Aplicações")

st.sidebar.page_link("Home.py", label="🏠 Home")
st.sidebar.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento sem filas")
st.sidebar.page_link("pages/2_Ajuste_Filas.py", label="🛠 Nivelamento com Filas")
st.sidebar.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparativo PR vs Plan")

# CONTEÚDO PRINCIPAL
st.title("📊 Ferramentas para Otimização de trabalho - MPS")
st.markdown(
    "<p style='font-size:18px; color:#9ca3af;'>Selecione um aplicativo para começar</p>",
    unsafe_allow_html=True
)
st.divider()
