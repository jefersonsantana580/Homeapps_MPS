
import streamlit as st


st.set_page_config(
    page_title="Ferramentas para Otimização de trabalho - MPS",
    page_icon="📊",
    layout="wide"
)


# ESCONDE MENU PADRÃO
st.markdown(    """
    <style>
        [data-testid="stSidebarNav"] {
            display: none;        }
    </style>    """,
    unsafe_allow_html=True
)

# SIDEBAR PERSONALIZADA
st.sidebar.image("images/agco.jpg")
st.sidebar.divider()

st.sidebar.markdown("### 📊 Aplicações")

st.sidebar.page_link("Home.py", label="🏠 Home")
st.sidebar.page_link("pages/1_Nivelamento.py", label="📈 Nivelamento sem filas")
st.sidebar.page_link("pages/2_NIvelar_com_Filas.py", label="🛠 Nivelamento com Filas")
st.sidebar.page_link("pages/3_Comparacao_Ciclo.py", label="🔄 Comparativo PR vs Plan")

st.sidebar.divider()  # 👈 SEPARAÇÃO CLARA


st.title("Ferramentas para Otimização de Trabalho – MPS")

st.markdown(
    """
    Conjunto de aplicações para suporte ao planejamento, nivelamento de carga
    e análise de aderência entre planejamento e execução.

    Selecione um aplicativo no menu à esquerda para começar.
    """
)

st.divider()

# ===== CARDS DOS APPS =====
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📊 Comparativo PV x Plan")
    st.write(
        """
        Compara o planejamento de produção com a execução real,
        identificando desvios, tendências e oportunidades de melhoria.
        """
    )

with col2:
    st.subheader("📈 Nivelamento sem filas")
    st.write(
        """
        Realiza o nivelamento da carga produtiva considerando apenas
        a capacidade disponível, sem aplicação de restrições.
        """
    )

with col3:
    st.subheader("📉 Nivelamento com filas")
    st.write(
        """
        Realiza o nivelamento considerando restrições produtivas,
        formação de filas e gargalos do processo.
        """
    )

st.divider()

# ===== COMO USAR =====
st.subheader("💡 Como usar")

st.markdown(
    """
    1. Selecione o aplicativo desejado no menu à esquerda  
    2. Baixe o **arquivo padrão** disponível no topo do aplicativo  
    3. Preencha o arquivo com seus dados  
    4. Faça o upload e analise os resultados gerados
    """
)

st.divider()

st.caption("Aplicação desenvolvida para suporte às análises do time MPS • Versão 1.0")

