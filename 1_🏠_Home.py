import os
import time
import streamlit as st
import pandas as pd
from datetime import datetime

# 🎯 CONFIGURAZIONE HOME UFFICIALE WINGAMING
st.set_page_config(page_title="Home", page_icon="🏠", layout="wide", initial_sidebar_state="expanded")

FILE_TECNICI = "elenco_tecnici.xlsx"

def carica_database_tecnici():
    if os.path.exists(FILE_TECNICI): return pd.read_excel(FILE_TECNICI).fillna("")
    return pd.DataFrame(columns=["NOME", "EMAIL", "PASSWORD", "RUOLO"])

df_tecnici = carica_database_tecnici()

# Inizializzazione delle variabili in memoria RAM
if "autenticato" not in st.session_state: st.session_state.autenticato = False
if "user_ruolo" not in st.session_state: st.session_state.user_ruolo = "TECNICO"

# =====================================================================================
# ⏱️ CENTRALINA 2 ORE DI MANUELA: TIMEOUT CON RESET TOTALE ANTI-LOOP (INTEGRATO IN HOME)
# =====================================================================================
if "ora_creazione_sessione" not in st.session_state:
    st.session_state.ora_creazione_sessione = time.time()

# Calcola i secondi passati dal login iniziale
tempo_trascorso = time.time() - st.session_state.ora_creazione_sessione

if st.session_state.autenticato and "user_nome" in st.session_state:
    if tempo_trascorso > 7200: # 👈 7200 secondi corrispondono a 2 ore esatte di autonomia!
        st.session_state.clear()
        st.session_state.autenticato = False
        st.session_state.ora_creazione_sessione = time.time()
        if "st" in locals() and hasattr(st, "query_params"):
            st.query_params.clear()
        st.warning("🔒 Sessione scaduta per inattività (Limite 2 ore). Effettua nuovamente il login.")
        time.sleep(2.0)
        st.rerun()
# =====================================================================================

# -------------------------------------------------------------------------------------
# 🛡️ TELEPASS AUTOMATICO DI MANUELA: AGGANCIA IL TOKEN PER RESTERA LOGGATI 2 ORE (ANTI-F5)
# -------------------------------------------------------------------------------------
if "token_sessione" in st.query_params:
    token_salvato = str(st.query_params["token_sessione"]).strip()
    if "_" in token_salvato:
        try:
            email_t = token_salvato.split("_")[0].strip().lower()
            ut = df_tecnici[df_tecnici["EMAIL"].astype(str).str.lower().str.strip() == email_t]
            if not ut.empty:
                st.session_state.autenticato = True
                st.session_state.user_email = email_t
                st.session_state.user_nome = str(ut["NOME"].values[0]).replace("[","").replace("]","").replace("'","").strip()
                st.session_state.user_ruolo = str(ut["RUOLO"].values[0]).strip().upper()
        except Exception:
            pass

# 🎨 CENTRALINA GRAFICA TITOLO CENTRATO E FONT LEGGERO
st.markdown("""
    <style>
    #MainMenu, footer, .stDecoration, [data-testid="stFooter"] { visibility: hidden !important; display: none !important; }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif !important; }
    h1 { color: #115e59 !important; font-size: 22px !important; text-align: center !important; font-weight: 800 !important; margin-bottom: 25px; }
    div[data-testid="stForm"] { background-color: #ffffff !important; border: 2px solid #94a3b8 !important; border-radius: 14px !important; padding: 20px !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    .user-badge { background-color: #ffffff; padding: 10px; border-radius: 8px; border: 2px solid #115e59; margin-bottom: 20px; text-align: center; color: #115e59 !important; font-weight: 400; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# 🛡️ PRIVACY BARRA LATERALE
if not st.session_state.autenticato:
    st.markdown("""<style>[data-testid="stSidebar"] { display: none !important; }</style>""", unsafe_allow_html=True)
else:
    ruolo_attuale = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()
    if ruolo_attuale not in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
        st.markdown("""<style>[data-testid="stSidebar"] { display: none !important; }</style>""", unsafe_allow_html=True)

# 🛡️ SCHERMATA LOGIN CENTRATA COMPATTA
if not st.session_state.autenticato:
    st.markdown("<h1>🛡️ ACCESSO AREA UTENTI</h1>", unsafe_allow_html=True)
    col_sx, col_centro, col_dx = st.columns([1, 1.2, 1])
    with col_centro:
        with st.form(key="modulo_login_home_definitivo_f5_timer"):
            st.write("🔒 Autenticazione Richiesta")
            input_email = st.text_input("Nome Utente (E-mail):").strip().lower()
            input_password = st.text_input("Password di Sicurezza:", type="password").strip()
            if st.form_submit_button("🚀 ACCEDI AL PORTALE"):
                utente_trovato = df_tecnici[(df_tecnici["EMAIL"].astype(str).str.lower().str.strip() == input_email) & (df_tecnici["PASSWORD"].astype(str).str.strip() == input_password)]
                if not utente_trovato.empty:
                    st.session_state.user_nome = str(utente_trovato.iloc[0]["NOME"]).strip()
                    st.session_state.user_email = str(utente_trovato.iloc[0]["EMAIL"]).strip()
                    st.session_state.user_ruolo = str(utente_trovato.iloc[0]["RUOLO"]).strip().upper()
                    st.session_state.autenticato = True
                    st.session_state.ora_creazione_sessione = time.time() # Fissa l'ora di inizio login
                    st.query_params["token_sessione"] = f"{st.session_state.user_email}_attivo"
                    st.rerun()
                else:
                    st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.get("user_nome", "UFFICIO")
ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()

# 🛡️ REINDIRIZZAMENTO AUTOMATICO DEI TECNICI
if ruolo_utente not in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
    st.switch_page("pages/2_🧳_Ferie_Esercenti.py")

# 🤍 INTERFACCIA BIANCA DI BENVENUTO HOME ADMIN
st.markdown("<h1>🏠 WinGaming Home</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='user-badge'>👤 Pannello Direzione: {esecutore_nome}</div>", unsafe_allow_html=True)
st.info("💡 Benvenuto. Questa è la pagina Home riservata agli amministratori della WinGaming. Utilizza il menù laterale di sinistra per navigare liberamente tra i moduli aziendali (Ferie, Magazzino, Incassi).")
