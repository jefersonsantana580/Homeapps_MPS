
import streamlit as st

def render_sidebar():
    # Esconde menu padrão do Streamlit
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
