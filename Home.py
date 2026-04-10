
import streamlit as st

st.set_page_config(
    page_title="Ferramentas para Otimização de trabalho - MPS",
    page_icon="📊",
    layout="wide"
)


st.markdown(
    """
    <style>
    /* Torna o título sticky no topo do conteúdo */
    div[data-testid="stMarkdownContainer"] > h1 {
        position: sticky;
        top: 0;
        background-color: #0e1117;
        padding: 10px 0;
        z-index: 50;
        border-bottom: 1px solid #333;
    }
    </style>
    """,
    unsafe_allow_html=True
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

 
st.subheader("📌 Entenda os aplicativos")  
st.markdown("### 📈 Nivelamento sem filas")
st.markdown(
    """
     Realiza o nivelamento diário do volume e faz a criação de filas considerando os parâmetros inseridos no menu do app. 
        Este é ideal para usar quando não tivermos filas criadas no JDE e nem as filas fictícias.
    """
)
st.divider()
st.markdown("### 📉 Nivelamento com filas")
st.markdown(
    """
     Realiza o nivelamento diário do volume considerando as filas existentes e sugere duas datas possíveis para ajuste:
     
     - Cenário 1: Faz o nivelamento considerando como principal objetivo a antecipação mínima de datas.
     - Cenário 2: Faz o nivelamento considerando como principal objetivo o nivelamento por modelos.
    
     Este App é ideal para um cenário onde já temos filas criadas e temos alguns dias com slots vazios.
    """
)

st.markdown("---")


st.markdown("### 📊 Comparativo P.Request x Op.Plan")
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
    .floating-footer {
        position: fixed;
        bottom: 20px;
        left: 50%;
        transform: translateX(-50%);
        background-color: rgba(14, 17, 23, 0.9);
        color: #ccc;
        padding: 8px 16px;
        border-radius: 8px;
        font-size: 0.75rem;
        z-index: 999;
        border: 1px solid #333;
        backdrop-filter: blur(4px);
    }
    </style>

    <div class="floating-footer">
        Aplicação desenvolvida para suporte às análises do time MPS • Versão 1.0 - Jeferson Santana/Copilot
    </div>
    """,
    unsafe_allow_html=True
)
