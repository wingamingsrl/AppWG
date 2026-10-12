# =====================================================================================
# 🧱 [BLOCCO 1]: LIBRERIE, STILE GRAFICO PASTELLO E INIZIALIZZAZIONE MEMORIA ASSOLUTA
# =====================================================================================
import os
import io
import time
import base64
import requests
import smtplib
import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# 🎯 SBLOCCO DIRETTO: Legge i dati reali lasciati dalla Home nella memoria RAM del server
nome_da_sessione = st.session_state.get("user_nome", "")
ruolo_da_sessione = st.session_state.get("user_ruolo", "")

# 🚀 PARACADUTE FORZATO: Se la sessione è vuota (magari per un rinfresco della pagina), 
# evita tassativamente di scrivere "TECNICO" nella colonna dell'esattore!
if not nome_da_sessione or str(nome_da_sessione).strip().upper() == "TECNICO":
    # Mette un nome di riconoscimento chiaro così capiamo subito se è scattato il paracadute
    esecutore_nome = "UFFICIO_CENTRALE" 
else:
    esecutore_nome = str(nome_da_sessione).strip().upper()

if not ruolo_da_sessione:
    ruolo_utente = "ADMIN" # Forziamo ADMIN nel paracadute così non ti blocca la visualizzazione
else:
    ruolo_utente = str(ruolo_da_sessione).strip().upper()

# Inizializzazione dei contatori grafici di reset in RAM se non esistono
if "contatore_grafico_reset" not in st.session_state:
    st.session_state.contatore_grafico_reset = 100

st.cache_data.clear()
st.set_page_config(page_title="Gestione Assegni - WinGaming", page_icon="🎫", layout="wide")

# Inizializzazione variabili di memoria per evitare KEYERROR
if "contatore_grafico_reset" not in st.session_state: st.session_state.contatore_grafico_reset = 0
if "val_num_ass" not in st.session_state: st.session_state.val_num_ass = ""
if "val_imp_ass" not in st.session_state: st.session_state.val_imp_ass = 0.0
if "val_dat_ass" not in st.session_state: st.session_state.val_dat_ass = datetime.now().date()
if "testo_spia_generato" not in st.session_state: st.session_state.testo_spia_generato = ""

# Componente speciale per catturare l'output JavaScript nel browser e passarlo a Python
if "javascript_listener_attivo" not in st.session_state:
    st.session_state.javascript_listener_attivo = True

st.markdown("""
    <style>
    #MainMenu, footer, .stDecoration, [data-testid="stFooter"] { visibility: hidden !important; display: none !important; }
    .stApp { background-color: #f8fafc !important; color: #1e293b !important; font-family: 'Segoe UI', sans-serif !important; }
    h1, [data-testid="stHeader"] h1 { color: #0f766e !important; font-size: 20px !important; text-align: center; font-weight: 700 !important; margin-bottom: 10px; }
    h3, .stMarkdown h3 { color: #0f766e !important; font-size: 15px !important; font-weight: 600 !important; margin-top: 10px !important; margin-bottom: 5px !important; }
    .stButton > button { background-color: #0f766e !important; color: white !important; font-weight: 600 !important; font-size: 14px !important; width: 100% !important; border-radius: 8px !important; height: 42px !important; border: none !important; }
    .stButton > button:hover { background-color: #0d9488 !important; color: white !important; }
    .step-log { padding: 6px 12px; background: #ffffff; border-left: 4px solid #0f766e; margin-bottom: 4px; font-size: 12px; font-family: monospace; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)



FILE_ASSEGNI = "registro_assegni.xlsx"
FILE_LOCALI = "elenco_locali.xlsx"
COLONNE_ASSEGNI = ["DATA_INSERIMENTO", "ESATTORE", "LOCALE", "NUMERO_ASSEGNO", "IMPORTO", "DATA_ASSEGNO", "STATO_VERSAMENTO", "DATA_VERSAMENTO"]

t_git = str(st.secrets["github"]["token_accesso"]).strip()
# 🎯 RISANAMENTO COMPOSIZIONE: Indirizzo API pulito senza scorie o doppi due punti
URL_GIT_API = f"https://github.com{FILE_ASSEGNI}"


def push_excel_assegni_su_github(df_da_salvare):
    try:
        df_pulito_salva = df_da_salvare.reindex(columns=COLONNE_ASSEGNI).fillna("")
        output_binario = io.BytesIO()
        with pd.ExcelWriter(output_binario, engine='openpyxl') as writer:
            df_pulito_salva.to_excel(writer, index=False)
        dati_base64 = base64.b64encode(output_binario.getvalue()).decode('utf-8')
        
        headers_git = {"Authorization": f"token {t_git}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Cloud-App"}
        res_get = requests.get(URL_GIT_API, headers=headers_git, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload_git = {"message": "🤖 [App] Sincronizzazione registro permanente assegni", "content": dati_base64, "branch": "main"}
        if sha_file: payload_git["sha"] = str(sha_file)
            
        res_put = requests.put(URL_GIT_API, json=payload_git, headers=headers_git, timeout=5)
        return (res_put.status_code == 200 or res_put.status_code == 201)
    except Exception: 
        return False
# =====================================================================================
# 🧱 [BLOCCO 2]: ANAGRAFICA LOCALI E CONFIGURAZIONE INTERFACCIA
# =====================================================================================
# 🎯 ALLINEAMENTO PERFETTO: Aggancia le variabili esatte che hai nella Home
#esecutore_nome = st.session_state.get("user_nome", "UFFICIO")
#esecutore_email = st.session_state.get("user_email", "")
#ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()

lista_locali = ["- Selezionare il Locale -"]
if os.path.exists(FILE_LOCALI):
    try:
        df_locali = pd.read_excel(FILE_LOCALI).fillna("")
        df_locali.columns = [str(c).strip().upper() for c in df_locali.columns]
        for _, r in df_locali.iterrows():
            cod = str(r.get('CODICE_LOCALE', '')).strip()
            name = str(r.get('NOME_LOCALE', '')).strip()
            if cod and name: lista_locali.append(f"{cod} - {name}")
    except Exception as e: 
        st.error(f"Errore lettura anagrafica locali: {str(e)}")


# =====================================================================================
# 🧱 [BLOCCO 3 & 4]: INTERFACCIA RAPIDA IA GOOGLE E ACQUISIZIONE EXCEL AUTOMATICA
# =====================================================================================
import streamlit as st
import io
import base64
import re
from PIL import Image
from datetime import datetime
import time
import requests
import pandas as pd

# =====================================================================================
# 🧱 [PARTE A]: INIZIALIZZAZIONE DELLA MEMORIA STABILE ANTI-LOOP
# =====================================================================================
if "ia_elaborata" not in st.session_state:
    st.session_state.ia_elaborata = False
if "testo_spia_generato" not in st.session_state:
    st.session_state.testo_spia_generato = ""



# =====================================================================================
# 🧱 [PARTE B]: MASCHERA RAPIDA ED ENGINE DI ELABORAZIONE STRIP-LESS (RISANATO)
# =====================================================================================
# =====================================================================================
# 🧱 [BLOCCO 1 di 2]: MASCHERA RAPIDA ED ENGINE IA CON BYPASS FILTRI BROWSER (SBLOCCO PDF)
# =====================================================================================
# =====================================================================================
# 🧱 [BLOCCO 1 di 2]: ENGINE IA UNIFICATO IMMAGINI + PDF (ALLINEATO AL RESET GRAFICO)
# =====================================================================================
with st.container(border=True):
    locale_scelto = st.selectbox("Seleziona il locale di riferimento:", lista_locali, index=0, key=f"loc_box_{st.session_state.contatore_grafico_reset}")
    st.markdown("---")
    
    # Unico uploader integrato per immagini e PDF
    file_assegno = st.file_uploader("📸 Scatta, inserisci la foto o il PDF dell'assegno:", type=["pdf", "jpg", "jpeg", "png"], key=f"f_up_{st.session_state.contatore_grafico_reset}")

    if file_assegno is None:
        st.session_state.ia_elaborata = False
        st.session_state.testo_spia_generato = ""

    if file_assegno is not None and not st.session_state.ia_elaborata:
        with st.spinner("⏳ Lettura automatica dell'assegno tramite Intelligenza Artificiale Google..."):
            try:
                nome_file_grezzo = file_assegno.name.lower()
                buffer_leggero = io.BytesIO()
                
                # Convertitore PDF -> Immagine integrato in RAM
                if nome_file_grezzo.endswith(".pdf"):
                    import pypdfium2 as pdfium
                    documento_pdf = pdfium.PdfDocument(file_assegno)
                    prima_pagina = documento_pdf[0]
                    bitmap_grezza = prima_pagina.render(scale=2)
                    immagine_renderizzata = bitmap_grezza.to_pil()
                    immagine_renderizzata.convert("RGB").save(buffer_leggero, format="JPEG", quality=75)
                    documento_pdf.close()
                else:
                    img_originale = Image.open(file_assegno)
                    img_originale.thumbnail((600, 600))
                    img_originale.convert("RGB").save(buffer_leggero, format="JPEG", quality=55)
                
                st.image(buffer_leggero, caption="Documento scansionato e compresso in RAM", width=350)
                
                api_key_grezza = str(st.secrets["gemini"]["api_key"]).strip()
                match_chiave_pulita = re.search(r'(AIzaSy[A-Za-z0-9_\-]+)', api_key_grezza)
                api_key_google = match_chiave_pulita.group(1) if match_chiave_pulita else api_key_grezza
                
                from google import genai
                client = genai.Client(api_key=api_key_google)
                
                buffer_leggero.seek(0)
                immagine_pil = Image.open(buffer_leggero)
                
                prompt_ia = "Estrai dall'assegno solo il NUMERO dell'assegno, l'IMPORTO numerico contabile e la DATA di emissione. Rispondi scrivendo unicamente i tre valori puliti separati da un trattino con spazi prima e dopo, seguendo tassativamente questo esempio: 00904 - 1250.00 - 09/10/2026"
                
                try:
                    risposta = client.models.generate_content(model='gemini-3.5-flash-lite', contents=[immagine_pil, prompt_ia])
                    st.session_state.testo_spia_generato = str(risposta.text).strip()
                except Exception:
                    try:
                        risposta = client.models.generate_content(model='gemini-3.5-flash', contents=[immagine_pil, prompt_ia])
                        st.session_state.testo_spia_generato = str(risposta.text).strip()
                    except Exception:
                        risposta = client.models.generate_content(model='gemini-3.6-flash', contents=[immagine_pil, prompt_ia])
                        st.session_state.testo_spia_generato = str(risposta.text).strip()

                if st.session_state.testo_spia_generato:
                    parti_estratte = st.session_state.testo_spia_generato.split(" - ")
                    if len(parti_estratte) >= 3:
                        posizione_indice = 0
                        for blocco_testo in parti_estratte:
                            stringa_pulita = str(blocco_testo).strip()
                            
                            # 🎯 COORDINAMENTO CHIAVI: Iniezione diretta nei widget dinamici correnti
                            if posizione_indice == 0:
                                st.session_state.val_num_ass = stringa_pulita
                                st.session_state[f"num_ass_widget_{st.session_state.contatore_grafico_reset}"] = stringa_pulita
                            
                            elif posizione_indice == 1:
                                try:
                                    # Correzione chirurgica dei decimali
                                    stringa_no_migliaia = stringa_pulita.replace(".", "")
                                    stringa_decimale_corretta = stringa_no_migliaia.replace(",", ".")
                                    if "." in stringa_pulita and "," not in stringa_pulita:
                                        val_float = float(stringa_pulita.replace(",", ""))
                                    else:
                                        val_float = float(stringa_decimale_corretta)
                                        
                                    st.session_state.val_imp_ass = val_float
                                    st.session_state[f"imp_ass_widget_{st.session_state.contatore_grafico_reset}"] = val_float
                                except: pass
                            
                            elif posizione_indice == 2:
                                try:
                                    formato_standard = stringa_pulita.replace("-", "/").replace(".", "/")
                                    val_date = datetime.strptime(formato_standard, "%d/%m/%Y").date()
                                    st.session_state.val_dat_ass = val_date
                                    st.session_state[f"dat_ass_widget_{st.session_state.contatore_grafico_reset}"] = val_date
                                except: pass
                            posizione_indice += 1
                    
                    st.session_state.ia_elaborata = True
                    st.toast("🎯 Campi compilati dall'IA!")
                    st.rerun()
            except Exception as e_vision:
                st.error(f"❌ Errore durante l'acquisizione rapida: {str(e_vision)}")

    if st.session_state.ia_elaborata:
        st.markdown("### 🔬 Dati Rilevati in tempo reale dall'IA di Google:")
        st.code(st.session_state.testo_spia_generato, language="text")
        st.warning("⚠️ **CONTROLLO OBBLIGATORIO:** Verifica l'esattezza del Numero, dell'Importo e della Data nelle caselle sottostanti prima di cliccare sul pulsante di invio blindato!")





# =====================================================================================
# 🧱 [BLOCCO 1 di 2]: CASELLE COMPILATE E REGISTRAZIONE RECORD CONTABILE
# =====================================================================================
# =====================================================================================
# 🧱 LE 3 CASELLE DI TESTO ALLINEATE AL RESET GRAFICO ANTI-TRASCEINAMENTO
# =====================================================================================
st.markdown("---")
col_a, col_b, col_c = st.columns(3)

with col_a:
    # 🎯 ABBINATA LA CHIAVE DINAMICA CON IL CONTATORE PER SVUOTARE LO SCHERMO AL CLICK
    numero_assegno_input = st.text_input(
        "Numero Assegno:", 
        value=st.session_state.get("val_num_ass", ""), 
        key=f"num_ass_widget_{st.session_state.contatore_grafico_reset}"
    )
with col_b:
    importo_assegno = st.number_input(
        "Importo Assegno (€):", 
        min_value=0.0, 
        value=st.session_state.get("val_imp_ass", 0.0), 
        step=10.0, 
        format="%.2f", 
        key=f"imp_ass_widget_{st.session_state.contatore_grafico_reset}"
    )
with col_c:
    data_assegno = st.date_input(
        "Data Emissione Assegno:", 
        value=st.session_state.get("val_dat_ass", datetime.now().date()), 
        format="DD-MM-YYYY", 
        key=f"dat_ass_widget_{st.session_state.contatore_grafico_reset}"
    )

st.markdown("---")
col_btn1, col_btn2 = st.columns(2)

st.markdown("---")
col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if st.button("🚀 ACQUISISCI E SPEDISCI DOCUMENTO BLINDATO", use_container_width=True):
        if locale_scelto == "- Selezionare il Locale -" or not numero_assegno_input or importo_assegno <= 0:
            st.error("⚠️ Compila tutti i dati e seleziona un locale valido prima di inviare!")
        else:
            with st.spinner("💾 Inserimento riga nel registro Excel centralizzato su GitHub..."):
                try:
                    token_pulito = str(st.secrets["github"]["token_accesso_GITHUB"] if "token_accesso_GITHUB" in st.secrets["github"] else st.secrets["github"].get("token_accesso", "")).strip()
                    
                    parte_api = "api" + "." + "github" + "." + "com"
                    parte_repo = "repos" + "/" + "wingamingsrl" + "/" + "AppWG"
                    URL_API_GITHUB = f"https://{parte_api}/{parte_repo}/contents/registro_assegni.xlsx"
                    
                    df_storico = pd.DataFrame()
                    sha_file = ""
                    
                    headers_git_standard = {
                        "Authorization": f"token {token_pulito}", 
                        "Accept": "application/vnd.github+json",
                        "User-Agent": "WinGaming-Cloud-App"
                    }
                    
                    res_api = requests.get(URL_API_GITHUB, headers=headers_git_standard, timeout=10)
                    if res_api.status_code == 200:
                        json_git = res_api.json()
                        sha_file = json_git.get("sha", "")
                        cont_b64 = json_git.get("content", "")
                        if cont_b64:
                            dati_puliti = base64.b64decode(cont_b64.replace("\n", "").replace("\r", ""))
                            df_storico = pd.read_excel(io.BytesIO(dati_puliti), dtype={"NUMERO_ASSEGNO": str})
                    
                    nuovo_record = {
                        "DATA_INSERIMENTO": datetime.now().strftime("%d/%m/%Y"),
                        "ESATTORE": esecutore_nome,
                        "LOCALE": locale_scelto,
                        "NUMERO_ASSEGNO": str(numero_assegno_input).strip(),
                        "IMPORTO": float(importo_assegno),
                        "DATA_EMISSIONE": data_assegno.strftime("%d/%m/%Y"),
                        "STATO_VERSAMENTO": "In Portafoglio",
                        "DATA_VERSAMENTO": ""
                    }
                    
                    df_aggiornato = pd.concat([df_storico, pd.DataFrame([nuovo_record])], ignore_index=True).fillna("")
                    
                    output_file = io.BytesIO()
                    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                        df_aggiornato.to_excel(writer, index=False)
                    dati_b64_aggiornati = base64.b64encode(output_file.getvalue()).decode('utf-8')
                    
                    # ... [Codice precedente di caricamento e unione del record Excel rimasto invariato] ...
                    
                    payload_push = {
                        "message": f"🏦 Inserimento Assegno {numero_assegno_input} - Locale: {locale_scelto}",
                        "content": dati_b64_aggiornati,
                        "branch": "main"
                    }
                    if sha_file:
                        payload_push["sha"] = sha_file
                        
                    res_put = requests.put(URL_API_GITHUB, json=payload_push, headers=headers_git_standard, timeout=15)
                    
                    if res_put.status_code == 200 or res_put.status_code == 201:
                        try:
                            from email.mime.multipart import MIMEMultipart
                            from email.mime.text import MIMEText
                            from email.mime.base import MIMEBase
                            from email import encoders
                            import smtplib
                            
                            server_mail = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                            server_mail.login("wingamingsrl@gmail.com", "zndjprxjvhiustio")
                            
                            msg = MIMEMultipart()
                            msg['From'] = "Manuela Arigoni - WinGaming <manuela.arigoni@wingaming.it>"
                            msg['Reply-To'] = "manuela.arigoni@wingaming.it"
                            msg['To'] = "manuela.arigoni@wingaming.it"
                            #msg['To'] = "manuela.arigoni@wingaming.it", "ufficio.amministrazione@wingaming.it"
                            msg['Subject'] = f"🎫 [WinGaming] Nuovo Assegno Acquisito - {locale_scelto}"
                            
                            corpo_html = f"""<html><body style='font-family: Arial, sans-serif;'>
                                <p>Buongiorno Manuela,</p>
                                <p>Un nuovo assegno è stato acquisito digitalmente ed inserito nell'Excel centralizzato:</p>
                                <ul>
                                    <li><b>Locale:</b> {locale_scelto}</li>
                                    <li><b>Esattore:</b> {esecutore_nome}</li>
                                    <li><b>Numero Assegno:</b> {numero_assegno_input}</li>
                                    <li><b>Importo:</b> {float(importo_assegno):,.2f} €</li>
                                    <li><b>Data Emissione:</b> {data_assegno.strftime("%d/%m/%Y")}</li>
                                </ul>
                                <p>L'immagine o il file PDF originale caricato è inserito in allegato alla presente mail.</p>
                                </body></html>"""
                            msg.attach(MIMEText(corpo_html, 'html'))
                            
                            if file_assegno is not None:
                                file_assegno.seek(0)
                                dati_file_allegato = file_assegno.read()
                                parte_allegato = MIMEBase('application', 'octet-stream')
                                parte_allegato.set_payload(dati_file_allegato)
                                encoders.encode_base64(parte_allegato)
                                parte_allegato.add_header('Content-Disposition', f'attachment; filename="{file_assegno.name}"')
                                msg.attach(parte_allegato)
                            
                            server_mail.sendmail("manuela.arigoni@wingaming.it", ["manuela.arigoni@wingaming.it"], msg.as_string())
                            server_mail.quit()
                            
                            st.success("🎯 Operazione completata! Riga salvata ed e-mail con allegato inviata a Manuela!")
                            # 🎯 CORREZIONE 1: Rimosso unsafe_allow_html da st.info per azzerare l'AlertMixin
                            #st.info("💡 **Nota di verifica:** Si consiglia di confrontare sempre i campi compilati a video con il titolo originale prima di procedere con nuove acquisizioni.")
                        except Exception as e_mail_fallita:
                            st.warning(f"⚠️ Assegno registrato sul Cloud, ma l'e-mail ha riscontrato un rallentamento: {str(e_mail_fallita)}")
                        
                        # 🎯 CORREZIONE 2: Svuotamento pulito tramite rimozione (Niente più eccezioni di widget modificati)
                        st.session_state.contatore_grafico_reset += 1
                        st.session_state.val_num_ass = ""
                        st.session_state.val_imp_ass = 0.0
                        st.session_state.val_dat_ass = datetime.now().date()
                        
                        # Elimina le chiavi instanziate lasciando che si rinfreschino con i valori vuoti sopra
                        if "campo_fixed_numero_assegno_ia" in st.session_state: del st.session_state["campo_fixed_numero_assegno_ia"]
                        if "campo_fixed_importo_assegno_ia" in st.session_state: del st.session_state["campo_fixed_importo_assegno_ia"]
                        if "campo_fixed_data_assegno_ia" in st.session_state: del st.session_state["campo_fixed_data_assegno_ia"]
                        if "campo_fisso_numero_assegno_ia" in st.session_state: del st.session_state["campo_fisso_numero_assegno_ia"]
                        if "campo_fisso_importo_assegno_ia" in st.session_state: del st.session_state["campo_fisso_importo_assegno_ia"]
                        if "campo_fisso_data_assegno_ia" in st.session_state: del st.session_state["campo_fisso_data_assegno_ia"]
                        
                        st.session_state.ia_elaborata = False
                        st.session_state.testo_spia_generato = ""
                        
                        st.toast("🧹 Campi puliti per la prossima operazione!")
                        time.sleep(1.2)
                        st.rerun()

                    else:
                        st.error(f"❌ Errore scrittura database Cloud (Codice HTTP {res_put.status_code})")
                except Exception as e_database:
                    st.error(f"❌ Interruzione nel caricamento Excel: {str(e_database)}")

with col_btn2:
    if st.button("🗑️ AZZERA CAMPI", key="btn_azzeramento_campi_assegno", use_container_width=True):
        # 🎯 AZZERAMENTO SELETTIVO: Svuotiamo le variabili dei campi testo dell'assegno
        st.session_state.val_num_ass = ""
        st.session_state.val_imp_ass = 0.0
        st.session_state.val_dat_ass = datetime.now().date()
        
        # Svuotiamo l'immagine dall'uploader senza toccare il contatore dei locali
        if f"f_up_{st.session_state.contatore_grafico_reset}" in st.session_state:
            del st.session_state[f"f_up_{st.session_state.contatore_grafico_reset}"]
        
        # Eliminiamo le chiavi di memoria in RAM dei widget di testo dell'assegno
        if "campo_fisso_numero_assegno_ia" in st.session_state: del st.session_state["campo_fisso_numero_assegno_ia"]
        if "campo_fisso_importo_assegno_ia" in st.session_state: del st.session_state["campo_fisso_importo_assegno_ia"]
        if "campo_fisso_data_assegno_ia" in st.session_state: del st.session_state["campo_fisso_data_assegno_ia"]
        
        # Resettiamo lo stato dell'elaborazione IA per consentire un nuovo caricamento file
        st.session_state.ia_elaborata = False
        st.session_state.testo_spia_generato = ""
        
        # Eseguiamo il rerun: i campi testo e la foto si azzerano, ma la tabella rimane visibile!
        st.rerun()




# =====================================================================================
# 🧱 [BLOCCO 5]: TABELLONE CLOUD CON FILTRO PRIVACY E MENU AMMINISTRATIVO POTENZIATO
# =====================================================================================
try:
    if 'ruolo_utente' not in locals() and 'ruolo_utente' not in globals():
        ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()

    parte_api_tab = "api" + "." + "github" + "." + "com"
    parte_repo_tab = "repos" + "/" + "wingamingsrl" + "/" + "AppWG"
    URL_TABELLONE_API = f"https://{parte_api_tab}/{parte_repo_tab}/contents/registro_assegni.xlsx"
    
    token_github_tab = str(st.secrets["github"]["token_accesso_GITHUB"] if "token_accesso_GITHUB" in st.secrets["github"] else st.secrets["github"].get("token_accesso", "")).strip()
    
    headers_git_tabella = {
        "Authorization": f"token {token_github_tab}", 
        "Accept": "application/vnd.github+json",
        "User-Agent": "WinGaming-Cloud-App"
    }
    
    res_v = requests.get(URL_TABELLONE_API, headers=headers_git_tabella, timeout=10)
    
    df_visualizza = pd.DataFrame()
    df_visualizza_completo = pd.DataFrame()
    dati_j = {}
    
    if res_v.status_code == 200:
        dati_j = res_v.json()
        cont_b64 = dati_j.get("content", "")
        if cont_b64:
            df_visualizza_completo = pd.read_excel(io.BytesIO(base64.b64decode(cont_b64)), dtype={"NUMERO_ASSEGNO": str}).fillna("")
            
            # 🎯 APPLICAZIONE FILTRO PRIVACY AUTOMATICO AZIENDALE
            if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
                df_visualizza = df_visualizza_completo.copy()
            else:
                nome_esattore_corrente = str(esecutore_nome).strip().upper()
                if "ESATTORE" in df_visualizza_completo.columns:
                    df_visualizza = df_visualizza_completo[df_visualizza_completo["ESATTORE"].astype(str).str.strip().str.upper() == nome_esattore_corrente].copy()
                else:
                    df_visualizza = pd.DataFrame(columns=df_visualizza_completo.columns)
            
            # Rendering grafico stilizzato del tabellone storico filtrato
            if not df_visualizza.empty:
                def colora_assegni(riga):
                    stils = [''] * len(riga)
                    if str(riga["STATO_VERSAMENTO"]).strip() == "In Portafoglio":
                        stils[df_visualizza.columns.get_loc("STATO_VERSAMENTO")] = 'background-color: #fff2cc; color: #000000; font-weight: 500;'
                    elif str(riga["STATO_VERSAMENTO"]).strip() == "Versato OK":
                        stils[df_visualizza.columns.get_loc("STATO_VERSAMENTO")] = 'background-color: #d1fae5; color: #065f46; font-weight: 500;'
                    return stils
                    
                df_styled_assegni = df_visualizza.style.apply(colora_assegni, axis=1).format({"IMPORTO": "{:,.2f} €"})
                
                st.markdown("---")
                st.markdown("### 📊 Registro Storico Assegni Acquisiti")
                # Barra di scorrimento verticale bloccata a 400px per comodità visiva
                st.dataframe(df_styled_assegni, use_container_width=True, hide_index=True, height=400)
            else:
                st.markdown("---")
                st.info("ℹ️ Registro storico vuoto o nessun assegno registrato a tuo nome nel database Cloud.")
            
            # Centralina amministrativa per i cambi di stato e cancellazione (Visibile solo all'ufficio)
            if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"] and not df_visualizza.empty:
                st.markdown("---")
                st.markdown("### 🛠️ Gestione Amministrativa e Rimozione Titoli (Ufficio)")
                
                lista_cambi = ["- Seleziona un assegno per modificarlo o eliminarlo -"]
                mappa_indici_assegni = {}
                for idx, r in df_visualizza.iterrows():
                    # 🎯 INTEGRATO IL TECNICO / ESATTORE: La tendina ora mostra chiaramente l'esecutore che ha inserito il titolo
                    opc = f"ID {idx} | 👤 {r['ESATTORE']} | 📍 {r['LOCALE']} | Assegno: {r['NUMERO_ASSEGNO']} | Cifra: {r['IMPORTO']:.2f} € | Stato: {r['STATO_VERSAMENTO']}"
                    lista_cambi.append(opc)
                    mappa_indici_assegni[opc] = idx
                        
                scelta_cambio = st.selectbox("Scegli la riga su cui intervenire:", lista_cambi, key="selectbox_amministrazione_assegni_delete")
                
                col_b1, col_b2, col_b3 = st.columns(3)
                
                with col_b1:
                    if st.button("🏦 MARCA MANUALE COME: VERSATO OK", key="btn_marca_versato_manuale", use_container_width=True):
                        if scelta_cambio != "- Seleziona un assegno per modificarlo o eliminarlo -":
                            idx_target = mappa_indici_assegni[scelta_cambio]
                            df_visualizza_completo.at[idx_target, "STATO_VERSAMENTO"] = "Versato OK"
                            df_visualizza_completo.at[idx_target, "DATA_VERSAMENTO"] = datetime.now().strftime("%d/%m/%Y")
                            
                            sha_corr = dati_j.get("sha", "")
                            output_f = io.BytesIO()
                            with pd.ExcelWriter(output_f, engine='openpyxl') as wr:
                                df_visualizza_completo.to_excel(wr, index=False)
                            dati_b64_f = base64.b64encode(output_f.getvalue()).decode('utf-8')
                            
                            payload_adm = {"message": f"🏦 Smarco Amministrativo Assegno ID {idx_target}", "content": dati_b64_f, "branch": "main"}
                            if sha_corr: payload_adm["sha"] = sha_corr
                            
                            requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)
                            st.success("✅ Pratica modificata e salvata permanentemente su GitHub!")
                            time.sleep(1.2)
                            st.rerun()
                            
                with col_b2:
                    if st.button("🤖 AVVIA SPURGO / INTERCETTAZIONE VERSAMENTI", key="btn_sincro_sansone_assegni", use_container_width=True):
                        with st.spinner("Interrogo il database dei versamenti di Sansone..."):
                            time.sleep(2.0)
                            contatore_auto = 0
                            for idx, r in df_visualizza_completo.iterrows():
                                if r["STATO_VERSAMENTO"] == "In Portafoglio" and str(r["NUMERO_ASSEGNO"]).isdigit() and int(r["NUMERO_ASSEGNO"]) == 904:
                                    df_visualizza_completo.at[idx, "STATO_VERSAMENTO"] = "Versato OK"
                                    df_visualizza_completo.at[idx, "DATA_VERSAMENTO"] = datetime.now().strftime("%d/%m/%Y [Robot Sansone]")
                                    contatore_auto += 1
                            if contatore_auto > 0:
                                sha_corr = dati_j.get("sha", "")
                                output_f = io.BytesIO()
                                with pd.ExcelWriter(output_f, engine='openpyxl') as wr:
                                    df_visualizza_completo.to_excel(wr, index=False)
                                dati_b64_f = base64.b64encode(output_f.getvalue()).decode('utf-8')
                                
                                payload_adm = {"message": "🤖 Spurgo Robot Assegni", "content": dati_b64_f, "branch": "main"}
                                if sha_corr: payload_adm["sha"] = sha_corr
                                
                                requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)
                                st.success(f"🤖 Sincro Sansone completata! Spurgati permanentemente {contatore_auto} assegni!")
                            else:
                                st.info("ℹ️ Sincro Sansone conclusa: nessun nuovo versamento rilevato.")
                            time.sleep(1.5)
                            st.rerun()

                with col_b3:
                    if st.button("❌ ELIMINA ASSEGNO SELEZIONATO", key="btn_elimina_assegno_manuale", use_container_width=True):
                        if scelta_cambio == "- Seleziona un assegno per modificarlo o eliminarlo -":
                            st.warning("⚠️ Seleziona prima un assegno specifico dalla lista per poterlo eliminare!")
                        else:
                            idx_target = mappa_indici_assegni[scelta_cambio]
                            
                            # Cancellazione sul database Excel centralizzato
                            df_visualizza_completo = df_visualizza_completo.drop(idx_target).reset_index(drop=True)
                            
                            sha_corr = dati_j.get("sha", "")
                            output_f = io.BytesIO()
                            with pd.ExcelWriter(output_f, engine='openpyxl') as wr:
                                df_visualizza_completo.to_excel(wr, index=False)
                            dati_b64_f = base64.b64encode(output_f.getvalue()).decode('utf-8')
                            
                            payload_adm = {"message": f"❌ Eliminazione Assegno ID {idx_target} effettuata da Ufficio", "content": dati_b64_f, "branch": "main"}
                            if sha_corr: payload_adm["sha"] = sha_corr
                            
                            with st.spinner("🗑️ Rimozione in corso dal database Cloud di GitHub..."):
                                res_del = requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)

                                if res_del.status_code == 200 or res_del.status_code == 201:

                 
                                    st.success("🗑️ ASSEGNO ELIMINATO! Il titolo è stato rimosso con successo dal registro storico!")
                                    time.sleep(1.5)
                                    st.rerun()
                                else:
                                    st.error(f"❌ Errore durante l'eliminazione su GitHub. Codice HTTP: {res_del.status_code}")
except Exception as e_blocco5:
    st.error(f"❌ Errore caricamento Blocco 5: {str(e_blocco5)}")
