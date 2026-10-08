import os
import io
import time
import requests
import smtplib
import streamlit as st
import pandas as pd
import datetime as dt_mod
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



# 🎯 SICUREZZA: Svuota la cache ad ogni rinfresco per forzare la lettura dei dati freschi da GitHub
st.cache_data.clear()

st.set_page_config(page_title="Report Incassi - WinGaming", page_icon="📊", layout="wide")
# 🎨 STILE GRAFICO AZIENDALE CON REGOLE COLORI DI MANUELA (NO ROSSO)
st.markdown("""
    <style>
    #MainMenu, footer, .stDecoration, [data-testid="stFooter"] { visibility: hidden !important; display: none !important; }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif !important; }
    h1, [data-testid="stHeader"] h1 { color: #0f766e !important; font-size: 20px !important; text-align: center; font-weight: 700 !important; margin-bottom: 10px; }
    h3, .stMarkdown h3 { color: #0f766e !important; font-size: 15px !important; font-weight: 600 !important; margin-top: 10px !important; margin-bottom: 5px !important; }
    
    .time-badge { 
        background-color: #f1f5f9 !important; border-left: 5px solid #0f766e !important; 
        padding: 10px !important; border-radius: 6px !important; font-size: 13px !important; color: #475569 !important; font-weight: 500; margin-bottom: 15px !important;
    }
    
    .stButton > button, .stDownloadButton > button { 
        background-color: #f8fafc !important; color: #334155 !important; border: 1px solid #e2e8f0 !important; 
        font-weight: 400 !important; font-size: 13px !important; width: 100% !important; border-radius: 8px !important; height: 36px !important; 
    }
    .stButton > button:hover, .stDownloadButton > button:hover { background-color: #f1f5f9 !important; border-color: #cbd5e1 !important; color: #0f172a !important; }
    </style>
""", unsafe_allow_html=True)

# 🛡️ PRIVACY RUOLO UTENTE
ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()
if ruolo_utente not in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
    st.error("🔒 Accesso Riservato alla direzione.")
    st.stop()

st.markdown("<h1>📊 Controllo Incassi — Locali non Incassati da Sansone</h1>", unsafe_allow_html=True)
st.markdown("<h3>🔄 Sincronizzazione ed Invio Report Esattori</h3>", unsafe_allow_html=True)

# 📂 CONFIGURAZIONE ROTTE FILE INTERNE ALLINEATE SU APPWG-TEST
FILE_INCASSI_GREZZO = "locali_non_incassati.xlsx"
FILE_ELENCO_ESATTORI = "elenco_esattori.xlsx"
repo_path = "wingamingsrl/AppWG"

t_git = str(st.secrets["github"]["token_accesso_GITHUB"] if "token_accesso_GITHUB" in st.secrets["github"] else st.secrets["github"].get("token_accesso", "")).strip()
s_api = "api" + "." + "github" + "." + "com"



# 🚀 TELECOMANDO CON RADAR PROGRESSIVO E BLINDATURA SUI FALLIMENTI DI RETE
if st.button("🚀 AVVIA ESTRAZIONE INCASSI DA SANSONE (Soglia 0 Giorni)", key="btn_lancio_incassi_manuale"):
    try:
        url_wf = f"https://{s_api}/repos/{repo_path}/dispatches"
        headers_lodi = {"Authorization": f"token {t_git}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Cloud-App"}
        payload = {"event_type": "avvia_robot_incassi"}
        
        res_lodi = requests.post(url_wf, json=payload, headers=headers_lodi, timeout=10)
        
        if res_lodi.status_code == 204 or res_lodi.status_code == 202:
            stato_attesa = st.empty()
            stato_attesa.info("⏳ Robot Sansone Incassi avviato nel Cloud... Sto piantonando l'Action su GitHub... Non toccare nulla.")
            
            time.sleep(6.0) # Pausa tecnica di registrazione nel Cloud
            
            url_runs_check = f"https://{s_api}/repos/{repo_path}/actions/workflows/cron_incassi_manuale.yml/runs?per_page=1"
            
            completato = False
            tentativi = 0
            max_tentativi = 50
            
            while not completato and tentativi < max_tentativi:
                tentativi += 1
                try:
                    res_check = requests.get(url_runs_check, headers=headers_lodi, timeout=5)
                    if res_check.status_code == 200:
                        dati_runs = res_check.json()
                        if dati_runs.get("workflow_runs"):
                            # 🟢 CORRETTO: Inserito l'indice [0] mancante per estrarre la riga corrente
                            run_corrente = dati_runs["workflow_runs"][0]
                            status_action = run_corrente.get("status", "").strip().lower()
                            conclusion_action = run_corrente.get("conclusion", "")
                            
                            if status_action == "completed":
                                completato = True
                                if str(conclusion_action).strip().lower() == "success":
                                    st.cache_data.clear()
                                    stato_attesa.success("✅ SINCRO RIUSCITA! Il robot ha scaricato i dati da Sansone e aggiornato la griglia!")
                                else:
                                    stato_attesa.error("❌ ERRORE CRITICO: Il robot su GitHub è andato in errore durante il login su Sansone! Verifica gli screenshot negli Artifacts.")
                            else:
                                secondi_trascorsi = tentativi * 5
                                stato_attesa.info(f"⚙️ Il server cloud sta elaborando il registro incassi... (Tempo trascorso: {secondi_trascorsi}s). Aspetto che finisca l'Action...")
                    
                    if not completato:
                        time.sleep(5.0)
                except Exception:
                    time.sleep(5.0)

            
            if not completato:
                st.warning("⏳ Il server cloud sta impiegando più tempo del previsto. Forza il rinfresco manuale tra qualche istante.")
            
            time.sleep(1.5)
            st.rerun()
        else:
            st.error(f"❌ Impossibile agganciare GitHub Actions. Risposta server: {res_lodi.status_code}")
    except Exception as e_click:
        st.error(f"💥 Errore di rete interno alla plancia: {str(e_click)}")



# 🎯 CONTROLLO INTEGRITÀ: Il file è valido solo se esiste e pesa più di 10 KB (evita file di errore vuoti)
#  LOGICA NUOVA: Chiede l'orario direttamente a GitHub Actions
data_ora_aggiornamento = "Data non disponibile"


file_realmente_valido = False
if os.path.exists(FILE_INCASSI_GREZZO) and os.path.getsize(FILE_INCASSI_GREZZO) > 10000:
    file_realmente_valido = True


# Se il file supera il controllo di peso, interroga le API di GitHub per mostrare la data di completamento reale
if file_realmente_valido:
    try:
        s_api = "api" + "." + "github" + "." + "com"
        url_runs = f"https://{s_api}/repos/{repo_path}/actions/workflows/{workflow_file}/runs?status=success&per_page=1"
        headers_runs = {"Authorization": f"token {t_git}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Cloud-App"}
        res_runs = requests.get(url_runs, headers=headers_runs, timeout=5)
        if res_runs.status_code == 200:
            dati_runs = res_runs.json()
            if dati_runs.get("workflow_runs"):
                conclusa_at = dati_runs["workflow_runs"][0]["updated_at"]
                dt_utc = datetime.strptime(conclusa_at, "%Y-%m-%dT%H:%M:%SZ")
                import datetime as dt_mod
                dt_locale = dt_utc + dt_mod.timedelta(hours=2)
                data_ora_aggiornamento = dt_locale.strftime("%d/%m/%Y alle ore %H:%M:%S")
    except Exception:
        # Paracadute: se le API di GitHub fossero temporaneamente offline, legge la data fisica del file locale
        data_ora_aggiornamento = datetime.fromtimestamp(os.path.getmtime(FILE_INCASSI_GREZZO)).strftime("%d/%m/%Y alle ore %H:%M:%S")






#st.markdown("---")

# 📊 IL BADGE INTELLIGENTE: Mostra l'orario reale solo se l'ultimo file è realmente valido e pesante
if file_realmente_valido:
    st.markdown(f"<div class='time-badge'>📅 <b>Ultimo scaricamento effettivo registro incassi:</b> {data_ora_aggiornamento}</div>", unsafe_allow_html=True)
else:
    st.markdown("<div class='time-badge'>📅 <b>Stato registro:</b> ⚠️ Attenzione: l'ultimo tentativo di sincronizzazione è fallito o l'Excel è corrotto. Rilancia l'estrazione.</div>", unsafe_allow_html=True)


# 🛡️ CARICAMENTO ANAGRAFICA ESATTORI E EMAIL DA FILE DEDICATO
opzioni_tendina_tecnici = []
mappa_esattori_email = {}
if os.path.exists(FILE_ELENCO_ESATTORI):
    try:
        df_esattori_reali = pd.read_excel(FILE_ELENCO_ESATTORI, header=0).fillna("")
        df_esattori_reali.columns = [str(c).strip().upper() for c in df_esattori_reali.columns]
        
        col_nome_effettiva = None
        col_email_effettiva = None
        for c in df_esattori_reali.columns:
            if "ESATTORE" in c or "NOME" in c: col_nome_effettiva = c
            if "EMAIL" in c or "POSTA" in c: col_email_effettiva = c
                
        if col_nome_effettiva:
            for idx, riga_e in df_esattori_reali.iterrows():
                nome_esattore = str(riga_e[col_nome_effettiva]).strip().upper()
                if nome_esattore and nome_esattore not in ["ESATTORE", ""]:
                    opzioni_tendina_tecnici.append(nome_esattore)
                    if col_email_effettiva:
                        mappa_esattori_email[nome_esattore] = str(riga_e[col_email_effettiva]).strip().lower()
                    else:
                        mappa_esattori_email[nome_esattore] = ""
                    
        opzioni_tendina_tecnici = sorted(list(set(opzioni_tendina_tecnici)))
    except Exception as e_esat:
        st.error(f"Errore lettura elenco_esattori.xlsx: {str(e_esat)}")

# 📊 STRUTTURA ED ELABORAZIONE DATI NATIVA
df_pulito_globale = pd.DataFrame()
if file_realmente_valido and os.path.exists(FILE_INCASSI_GREZZO):
    try:
        df_grezzo = pd.read_excel(FILE_INCASSI_GREZZO, skiprows=1).fillna(0)
        df_grezzo.columns = [str(c).strip() for c in df_grezzo.columns]
        
        df_pulito = pd.DataFrame()
        df_pulito["Nome Locale"] = df_grezzo.iloc[:, 0].astype(str)
        df_pulito["Esattore Rif."] = df_grezzo.iloc[:, 1].astype(str).str.upper()
        df_pulito["Commerciale"] = df_grezzo.iloc[:, 2].astype(str)
        df_pulito["Giorni"] = pd.to_numeric(df_grezzo.iloc[:, 5], errors='coerce').fillna(0).astype(int)
        df_pulito["Residuo Incassi"] = pd.to_numeric(df_grezzo.iloc[:, 6], errors='coerce').fillna(0.0).astype(float)
        df_pulito["Contabilità"] = pd.to_numeric(df_grezzo.iloc[:, 9], errors='coerce').fillna(0.0).astype(float)
        df_pulito["Saldo Acconti"] = pd.to_numeric(df_grezzo.iloc[:, 10], errors='coerce').fillna(0.0).astype(float)
        df_pulito["Da Incassare"] = (df_pulito["Contabilità"] - df_pulito["Saldo Acconti"]).astype(float)
        df_pulito["Tot Locale"] = pd.to_numeric(df_grezzo.iloc[:, 13], errors='coerce').fillna(0.0).astype(float)
        
        def associations_esattore_manuela(valore_sansone, lista_tecnici):
            testo = str(valore_sansone).upper().replace("WG", "").strip()
            parole_s = set(testo.split())
            if not parole_s: return "SCONOSCIUTO"
            for t in lista_tecnici:
                if set(t.split()).issubset(parole_s) or parole_s.issubset(set(t.split())): return t
            return "WG " + testo

        df_pulito["Esattore"] = df_pulito["Esattore Rif."].apply(lambda x: associations_esattore_manuela(x, opzioni_tendina_tecnici))
        df_pulito_globale = df_pulito[["Nome Locale", "Esattore", "Commerciale", "Giorni", "Residuo Incassi", "Contabilità", "Saldo Acconti", "Da Incassare", "Tot Locale"]].copy()
        st.markdown("### 🎛️ Centralina Selezione Esattore / Tecnico")
        tecnico_scelto = st.selectbox("Seleziona l'esattore ufficiale per l'anteprima:", ["MOSTRA TUTTI"] + opzioni_tendina_tecnici, key="selectbox_esattore_incassi")
        
        if tecnico_scelto != "MOSTRA TUTTI":
            df_filtrato = df_pulito_globale[df_pulito_globale["Esattore"] == tecnico_scelto].copy()
        else:
            df_filtrato = df_pulito_globale.copy()
            
        df_filtrato_view = df_filtrato.copy()
        df_filtrato_view = df_filtrato_view.sort_values(by="Da Incassare", ascending=False)
        sommatoria_da_incassare = float(df_filtrato_view["Da Incassare"].sum())

        st.markdown(f"### 📋 Registro Locali non Incassati ({len(df_filtrato_view)} posizioni rilevate)")
        
        def colora_griglia_manuela(riga):
            stili = [''] * len(riga)
            if int(riga["Giorni"]) >= 30:
                stili[df_filtrato_view.columns.get_loc("Giorni")] = 'font-weight: 700; color: #000000;'
            if float(riga["Residuo Incassi"]) > 0: 
                stili[df_filtrato_view.columns.get_loc("Residuo Incassi")] = 'background-color: #fff2cc; color: #000000;'
            if float(riga["Da Incassare"]) >= 1400.0: 
                stili[df_filtrato_view.columns.get_loc("Da Incassare")] = 'background-color: #93c5fd; color: #000000;'
            return stili

        df_styled = df_filtrato_view.style.apply(colora_griglia_manuela, axis=1).format({
            "Giorni": "{:d}", "Residuo Incassi": "{:,.2f} €", "Contabilità": "{:,.2f} €", 
            "Saldo Acconti": "{:,.2f} €", "Da Incassare": "{:,.2f} €", "Tot Locale": "{:,.2f} €"
        })
        
        st.dataframe(df_styled, hide_index=True, use_container_width=True)
        st.markdown(f"<div style='background-color: #002060; color: #ffffff; padding: 12px; border-radius: 6px; text-align: center; font-size: 16px; font-weight: 700; margin-top: 10px;'>🔹 TOTALE REGISTRO: {sommatoria_da_incassare:,.2f} €</div>", unsafe_allow_html=True)
        
        # ✉️ CENTRALINA DI SPEDIZIONE AUTONOMA CON GMAIL
        st.markdown("---")
        st.markdown("### ✉️ Centralina Spedizione Report Mail Raggruppate")
        
        destinatario_tipo = st.radio("Seleziona la modalità di invio email per il test:", ["Invia Report Singolo Selezionato", "Invia Report Completo a Tutti gli Esattori (Raggruppati per Email)"], horizontal=True, key="radio_tipo_destinatario_incassi")
        
        def genera_tabella_html_mail(df_da_inviare):
            html = """<table border='1' style='border-collapse: collapse; font-family: Segoe UI, sans-serif; font-size: 13px; width: 100%; border-color: #cbd5e1; text-align: left;'>
                <tr style='background-color: #f1f5f9; color: #1e293b; font-weight: bold;'>
                    <th style='padding: 8px;'>Nome Locale</th><th style='padding: 8px;'>Esattore</th><th style='padding: 8px;'>Commerciale</th><th style='padding: 8px;'>Giorni</th>
                    <th style='padding: 8px;'>Residuo Incassi</th><th style='padding: 8px;'>Contabilità</th><th style='padding: 8px;'>Saldo Acconti</th>
                    <th style='padding: 8px;'>Da Incassare</th><th style='padding: 8px;'>Tot Locale</th>
                </tr>"""
            for _, r in df_da_inviare.iterrows():
                weight_giorni = "font-weight: bold; color: #000000;" if int(r['Giorni']) >= 30 else ""
                bg_residuo = "background-color: #fff2cc;" if float(r['Residuo Incassi']) > 0 else ""
                bg_da_inc = "background-color: #93c5fd; color: #000000;" if float(r['Da Incassare']) >= 1400.0 else ""
                
                html += f"""<tr>
                    <td style='padding: 8px;'>{r['Nome Locale']}</td><td style='padding: 8px;'>{r['Esattore']}</td><td style='padding: 8px;'>{r['Commerciale']}</td>
                    <td style='padding: 8px; {weight_giorni}'>{int(r['Giorni'])}</td>
                    <td style='padding: 8px; {bg_residuo}'>{float(r['Residuo Incassi']):,.2f} €</td><td style='padding: 8px;'>{float(r['Contabilità']):,.2f} €</td>
                    <td style='padding: 8px;'>{float(r['Saldo Acconti']):,.2f} €</td><td style='padding: 8px; {bg_da_inc}'>{float(r['Da Incassare']):,.2f} €</td>
                    <td style='padding: 8px;'>{float(r['Tot Locale']):,.2f} €</td>
                </tr>"""
            tot_parziale = df_da_inviare['Da Incassare'].sum()
            html += f"""<tr style='background-color: #002060; color: #ffffff; font-weight: bold;'>
                <td colspan='7' style='padding: 10px; text-align: right;'>🔹 TOTALE DA INCASSARE:</td>
                <td colspan='2' style='padding: 10px;'>{tot_parziale:,.2f} €</td>
            </tr></table>"""
            return html

        if st.button("🚀 INVIA REPORT EMAIL ORA", key="btn_spedisci_mail_incassi"):
            try:
                server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                server.login("wingamingsrl@gmail.com", "zndjprxjvhiustio")
                
                mappa_destinatari_lavoro = {}
                if destinatario_tipo == "Invia Report Singolo Selezionato":
                    if tecnico_scelto == "MOSTRA TUTTI":
                        st.warning("⚠️ Seleziona un singolo esattore specifico per l'invio!")
                        server.quit()
                        st.stop()
                    em = mappa_esattori_email.get(tecnico_scelto, "")
                    if em: mappa_destinatari_lavoro[em] = [tecnico_scelto]
                else:
                    for tec, em in mappa_esattori_email.items():
                        if em:
                            if em not in mappa_destinatari_lavoro: mappa_destinatari_lavoro[em] = []
                            mappa_destinatari_lavoro[em].append(tec)
                
                if not mappa_destinatari_lavoro:
                    st.error("❌ Nessun indirizzo email trovato nel file elenco_esattori.xlsx!")
                    server.quit()
                    st.stop()

                contatore_effettivo_inviate = 0
                for mail_dest_originale, lista_esattori in mappa_destinatari_lavoro.items():
                    somma_controllo_mail = 0.0
                    blocchi_html_esattori = ""
                    
                    for tec in lista_esattori:
                        df_tec_mail = df_pulito_globale[df_pulito_globale["Esattore"] == tec].copy()
                        df_tec_mail = df_tec_mail.sort_values(by="Da Incassare", ascending=False)
                        
                        if not df_tec_mail.empty:
                            somma_controllo_mail += float(df_tec_mail["Da Incassare"].sum())
                            blocchi_html_esattori += f"<h3 style='font-family: Segoe UI, sans-serif; color: #0f766e; margin-top:20px;'>👤 Riepilogo Esattore: {tec}</h3>"
                            blocchi_html_esattori += genera_tabella_html_mail(df_tec_mail)
                    
                    if somma_controllo_mail <= 0.0:
                        continue
                        
                    msg = MIMEMultipart()
                    # 🎯 VITTORIA: Il mittente e l'indirizzo di risposta sono ufficialmente allineati alla tua mail aziendale!
                    msg['From'] = "Manuela Arigoni - WinGaming <manuela.arigoni@wingaming.it>"
                    msg['Reply-To'] = "manuela.arigoni@wingaming.it"
                    msg['To'] = "manuela.arigoni@wingaming.it"
                    msg['Subject'] = f"📊 [WinGaming] Registro Locali da Incassare"
                    
                    # STRUTTURA REALE PRODUCTION PRONTA PER IL FUTURO
                    # msg['To'] = mail_dest_originale
                    # msg['Cc'] = "manuela.arigoni@wingaming.it, alessandro.frigerio@wingaming.it"
                    
                    corpo_html = f"""<html><body>
                        <p style='font-family: Segoe UI, sans-serif; font-size: 14px;'>Buongiorno,<br><br>
                        Di seguito viene riportato il registro ufficiale dei locali non incassati assegnati alla vostra gestione, aggiornato ad oggi.</p>
                        {blocchi_html_esattori}
                        <p style='font-family: Segoe UI, sans-serif; font-size: 12px; color: #64748b; margin-top:20px;'>🤖 Messaggio automatico inviato dalla Plancia WinGaming Cloud.</p></body></html>"""
                    
                    msg.attach(MIMEText(corpo_html, 'html'))
                    server.sendmail("manuela.arigoni@wingaming.it", ["manuela.arigoni@wingaming.it"], msg.as_string())
                    
                    # LOGICA REALE PRODUCTION PRONTA PER IL FUTURO
                    # tutti_ricevitori = [mail_dest_originale, "manuela.arigoni@wingaming.it", "alessandro.frigerio@wingaming.it"]
                    # server.sendmail("manuela.arigoni@wingaming.it", tutti_ricevitori, msg.as_string())
                    
                    contatore_effettivo_inviate += 1
                
                server.quit()
                if contatore_effettivo_inviate > 0:
                    st.success(f"✅ Invio completato! Spedite {contatore_effettivo_inviate} email con tabelle incluse.")
                else:
                    st.info("ℹ️ Nessun esattore ha importi superiori a zero. Nessun report inviato.")
                
            except Exception as e_smtp:
                st.error(f"❌ Errore durante la spedizione: {str(e_smtp)}")
            
    except Exception as e_main: st.error(f"Errore caricamento griglia contabile: {str(e_main)}")
else: st.info("⏳ Il registro dei locali non incassati è in attesa dei dati reali da Sansone. Premi il pulsante in cima 'AVVIA ESTRAZIONE INCASSI DA SANSONE' per caricare la griglia!")
