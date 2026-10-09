import streamlit as st
import time

st.set_page_config(page_title="Disconnetti", page_icon="🚪", layout="wide")

# 🎨 STILE GRAFICO LEGGERO DI MANUELA
st.markdown("""
    <style>
    #MainMenu, footer, .stDecoration, [data-testid="stFooter"] { visibility: hidden !important; display: none !important; }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif !important; }
    h1 { color: #115e59 !important; font-size: 22px !important; text-align: center; font-weight: 800 !important; margin-bottom: 15px; }
    button, .stButton>button { 
        background-color: #f1f5f9 !important; color: #0f172a !important; border: 1px solid #cbd5e1 !important; 
        font-weight: 400 !important; font-size: 14px !important; width: 100%; border-radius: 8px !important; height: 38px !important; 
        transition: all 0.2s ease; 
    }
    button:hover, .stButton>button:hover { background-color: #e2e8f0 !important; border-color: #94a3b8 !important; color: #020617 !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🚪 Disconnessione dal Sistema WinGaming")
st.write("Clicca sul pulsante qui sotto per chiudere la sessione di lavoro in totale sicurezza e svuotare la memoria.")

# 🚀 IL MOTORE DI LOGOUT REALE INTERNO ISOLATO
if st.button("🚪 CONFERMA LOGOUT DEFINITIVO", key="btn_logout_pagina_autonoma"):
    st.session_state.clear()
    st.session_state.autenticato = False
    st.session_state.user_ruolo = "TECNICO"
    if "token_sessione" in st.query_params:
        del st.query_params["token_sessione"]
    
    st.success("🔒 Sessione chiusa. Ritorno al Login...")
    time.sleep(0.5)
    # Spinge duro l'utente fuori dalle sottopagine e lo rimanda alla Home iniziale di Login
    st.switch_page("1_🏠_Home.py")
