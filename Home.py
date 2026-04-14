
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
        Este é ideal para usar no período que não tiver filas criadas no JDE e nem as filas fictícias.
    """
)
st.divider()
st.markdown("### 📉 Nivelamento com filas")
st.markdown(
    """
    Realiza o nivelamento diário do volume considerando as filas existentes e sugere três possibilidades de ajuste, respeitando a capacidade diária informada.


    Cenário 1: prioriza a antecipação mínima das datas, buscando manter as filas o mais próximo possível da data planejada, com respeito ao limite diário.


    Cenário 2: prioriza o balanceamento diário por modelo, distribuindo o volume de forma mais uniforme entre os dias úteis do mês.


    Cenário 3: aplica um balanceamento híbrido: para modelos com menor volume no mês, tenta manter as filas próximas da data planejada, espalhando apenas quando necessário; 
                para os demais modelos, segue a lógica de nivelamento diário.

Este app é ideal para cenários em que já existem filas criadas e é necessário redistribuir o volume diário com melhor equilíbrio e menor impacto nas datas originais.
    
    """
)

st.markdown("---")


st.markdown("### 📊 Comparativo P.Request x Op.Plan")
st.markdown(
    """
    Este app tem como objetivo comparar o P. Request com o Op. Plan identificando e trazendo as diferenças por filial,
     produto, mercado etc de forma rápida com um visual claro e simples.
    """
)



st.markdown("---")




st.markdown(
    """
    <style>
    :root {
        --sidebar-width: 21rem;
    }

    .footer-bar {
        position: fixed;
        bottom: 0;
        left: var(--sidebar-width);
        width: calc(100% - var(--sidebar-width));
        background-color: rgba(14, 17, 23, 0.95);
        color: #ccc;
        font-size: 0.75rem;
        text-align: center;
        padding: 8px 0;
        z-index: 999;
        border-top: 1px solid #333;
        backdrop-filter: blur(4px);
    }
    </style>

    <div class="footer-bar">
        Aplicação desenvolvida para suporte às análises do time MPS • Versão 1.0 — Jeferson Santana / Copilot
    </div>
    """,
    unsafe_allow_html=True
)
