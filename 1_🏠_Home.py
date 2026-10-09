import os
import time
import streamlit as st
import pandas as pd

# 🎯 CONFIGURAZIONE HOME UFFICIALE WINGAMING
st.set_page_config(page_title="Home - WinGaming", page_icon="🏠", layout="wide", initial_sidebar_state="expanded")

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

tempo_trascorso = time.time() - st.session_state.ora_creazione_sessione

if st.session_state.autenticato and "user_nome" in st.session_state:
    if tempo_trascorso > 7200:
        st.session_state.clear()
        st.session_state.autenticato = False
        st.session_state.ora_creazione_sessione = time.time()
        if "st" in locals() and hasattr(st, "query_params"):
            st.query_params.clear()
        st.warning("🔒 Sessione scaduta per inattività (Limite 2 ore). Effettua nuovamente il login.")
        time.sleep(2.0)
        st.rerun()

# -------------------------------------------------------------------------------------
# 🛡️ TELEPASS AUTOMATICO DI MANUELA: AGGANCIA IL TOKEN PER RESTARE LOGGATI 2 ORE
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

# 🎨 CENTRALINA GRAFICA TITOLO CENTRATO E COPERTURA DOPPIO MENU NATIVO
st.markdown("""
    <style>
    #MainMenu, footer, .stDecoration, [data-testid="stFooter"] { visibility: hidden !important; display: none !important; }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif !important; }
    h1 { color: #0f766e !important; font-size: 22px !important; text-align: center !important; font-weight: 800 !important; margin-bottom: 25px; }
    div[data-testid="stForm"] { background-color: #ffffff !important; border: 2px solid #94a3b8 !important; border-radius: 14px !important; padding: 20px !important; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
    .user-badge { background-color: #ffffff; padding: 10px; border-radius: 8px; border: 2px solid #0f766e; margin-bottom: 20px; text-align: center; color: #0f766e !important; font-weight: 400; font-size: 14px; }
    
    /* 🎯 PIALLATURA STRUTTURALE DEL MENU DI SERVIZIO AUTOMATICO DI STREAMLIT */
    [data-testid="stSidebarNav"] { display: none !important; }
    </style>
""", unsafe_allow_html=True)

# 🛡️ MASCHERA LOGIN CENTRATA COMPATTA
if not st.session_state.autenticato:
    st.markdown("""<style>[data-testid="stSidebar"] { display: none !important; }</style>""", unsafe_allow_html=True)
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
                    st.session_state.ora_creazione_sessione = time.time()
                    st.query_params["token_sessione"] = f"{st.session_state.user_email}_attivo"
                    st.rerun()
                else:
                    st.error("❌ Credenziali errate. Riprova.")
    st.stop()

esecutore_nome = st.session_state.get("user_nome", "UFFICIO")
esecutore_email = st.session_state.get("user_email", "")
ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()

# 🎯 CONFIGURAZIONE SIDEBAR DINAMICA INTEGRATA ED UNIFICATA
with st.sidebar:
  
    if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
        # Vista Amministratore Ufficio Completa
        st.page_link("1_🏠_Home.py", label="Home", icon="🏠")
        st.page_link("pages/2_🧳_Ferie_Esercenti.py", label="Ferie Esercenti", icon="🧳")
        st.page_link("pages/3_📦_Giacenza_Magazzino.py", label="Giacenza Magazzino", icon="📦")
        st.page_link("pages/4_📊_Report_Incassi.py", label="Report Incassi", icon="📊")
        st.page_link("pages/5_🎫_Gestione_Assegni.py", label="Gestione Assegni", icon="🎫")
        st.page_link("pages/9_🚪_Disconnetti.py", label="Logout", icon="🚪")
       
    else:
        # 📱 Vista Tecnici Territorio: Solo le 3 voci richieste sbloccate
        st.page_link("1_🏠_Home.py", label="Home", icon="🏠")
        st.page_link("pages/2_🧳_Ferie_Esercenti.py", label="Ferie Esercenti", icon="🧳")
        st.page_link("pages/5_🎫_Gestione_Assegni.py", label="Scansione Assegni", icon="🎫")
        st.page_link("pages/9_🚪_Disconnetti.py", label="Logout", icon="🚪")
       
# 🤍 INTERFACCIA DI BENVENUTO IN BASE AL PRIVILEGIO DI RUOLO
st.markdown("<h1>🏠 WinGaming Home</h1>", unsafe_allow_html=True)

if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
    st.markdown(f"<div class='user-badge'>👤 Pannello Direzione: {esecutore_nome}</div>", unsafe_allow_html=True)
    st.info("💡 Benvenuto nel Portale Amministrativo WinGaming. Utilizza la barra laterale sinistra per gestire e coordinare i registri aziendali del magazzino, degli incassi e dei titoli contabili in portafoglio.")
else:
    st.markdown(f"<div class='user-badge'>👤 Portale Operativo Tecnici: {esecutore_nome}</div>", unsafe_allow_html=True)
    st.success("🔓 Accesso eseguito correttamente. Questa è l'area di accoglienza per gli operatori sul territorio. Utilizza il menu di sinistra per registrare i periodi di ferie dei locali o per scansionare visivamente gli assegni riscossi.")

