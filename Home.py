
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
        Selecione um aplicativo no menu à esquerda para começar.
    """
)

st.divider()

# ===== COMO USAR =====
st.subheader("💡 Como usar")

st.markdown(
    """
    1. Selecione o aplicativo desejado no menu à esquerda  
    2. Baixe o **arquivo padrão** disponível no topo do aplicativo.
    3. Preencha o arquivo com seus dados  
    4. Faça o upload e analise os resultados gerados
    """
)

st.divider()

 
st.divider()
st.subheader("📌 Entenda os aplicativos")
  
st.markdown("### 📈 Nivelamento sem filas")
st.markdown(
    """
     Realiza o nivelamento diário do volume e faz a criação de filas considerando os parâmetros inseridos no menu do app. 
        Este é ideal para usar quando não tivermos filas criadas no JDE e nem as filas fictícias.
    """
)


st.markdown("### 📉 Nivelamento com filas")
st.markdown(
    """
     Realiza o nivelamento diário do volume considerando as filas existentes e sugere duas datas possíveis para ajuste:
        Cenário 1: Faz o nivelamento considerando como principal objetivo a antecipação mínima de datas.
        Cenário 2: Faz o nivelamento considerando como principal objetivo o nivelamento por modelos.
        Este App é ideal para um cenário onde já temos filas criadas e temos alguns dias com slots vazios.
    """
)

st.markdown("---")


st.markdown("📊 Comparativo P.Request x Op.Plan")
st.markdown(
    """
    Este app tem como objetivo comparar o P. Request com o Op. Plan identificando e mostrando diferenças por filial,
        produto, mercado etc de forma rápida e com um visual claro.
    """
)



st.markdown("---")


st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: transparent;
        color: #888;
        text-align: center;
        padding: 10px;
        font-size: 0.75rem;
        z-index: 100;
    }
    </style>

    <div class="footer">
        Aplicação desenvolvida para suporte às análises do time MPS • Versão 1.0
    </div>
    """,
    unsafe_allow_html=True
)




