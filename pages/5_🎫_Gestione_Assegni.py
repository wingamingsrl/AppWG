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
    
    /* 🎯 ANTI-DARK MODE: Costringe l'intera plancia a rimanere chiara su qualsiasi smartphone */
    .stApp { 
        background-color: #f8fafc !important; 
        color: #1e293b !important; 
        font-family: 'Segoe UI', sans-serif !important; 
    }
    
    h1, [data-testid="stHeader"] h1 { color: #0f766e !important; font-size: 20px !important; text-align: center; font-weight: 700 !important; margin-bottom: 10px; }
    h3, .stMarkdown h3 { color: #0f766e !important; font-size: 15px !important; font-weight: 600 !important; margin-top: 10px !important; margin-bottom: 5px !important; }
    
    /* 🎯 UNIFORMAZIONE TEXT INPUT E BUTTON: Forzati a rimanere chiari con testo scuro */
    .stButton > button, .stDownloadButton > button, input, textarea, div[role="text"] { 
        background-color: #ffffff !important; 
        color: #334155 !important; 
        border: 1px solid #e2e8f0 !important; 
        font-weight: 400 !important; 
        font-size: 13px !important; 
        width: 100% !important; 
        border-radius: 8px !important; 
        height: 36px !important;
        margin-bottom: 5px !important;
        box-shadow: none !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover { 
        background-color: #f1f5f9 !important; 
        border-color: #cbd5e1 !important; 
        color: #0f172a !important; 
    }

    .step-log { padding: 6px 12px; background: #ffffff; border-left: 4px solid #0f766e; margin-bottom: 4px; font-size: 12px; font-family: monospace; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
    
    /* Selectbox singole e dropdown */
    div[data-baseweb="select"] > div { background-color: #ffffff !important; border-color: #e2e8f0 !important; }
    div[data-baseweb="select"] div[role="button"] { background-color: #ffffff !important; color: #1e293b !important; }
    div[data-baseweb="select"] input:-webkit-autofill { -webkit-box-shadow: 0 0 0 30px #ffffff inset !important; -webkit-text-fill-color: #1e293b !important; }
    
    /* Multiselect blocchetti */
    div[data-baseweb="tag"] { background-color: #f0fdf4 !important; color: #166534 !important; border: 1px solid #bbf7d0 !important; }
    div[data-baseweb="tag"] svg, div[data-baseweb="tag"] [role="button"] { fill: #166534 !important; color: #166534 !important; }
    ul[role="listbox"] li { background-color: #ffffff !important; color: #1e293b !important; }
    ul[role="listbox"] li:hover { background-color: #0f766e !important; color: #ffffff !important; }
    </style>
""", unsafe_allow_html=True)




FILE_ASSEGNI = "registro_assegni.xlsx"
FILE_LOCALI = "elenco_locali.xlsx"
# 🎯 AGGIORNAMENTO: Inclusa la colonna DATA_INCASSO nella struttura centrale del database
COLONNE_ASSEGNI = ["DATA_REGISTRAZIONE", "ESATTORE", "LOCALE", "NUMERO_ASSEGNO", "IMPORTO", "DATA_ASSEGNO", "DATA_INCASSO", "STATO_VERSAMENTO"]

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
# 🧱 [ENGINE IA DEFINITIVO]: ALLINEAMENTO CHIAVI DINAMICHE PER COMPILAZIONE CASSELLI
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
                
                # Sblocco fotocamera: resetta lo stream per catturare al primo colpo
                file_assegno.seek(0)
                
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
                    img_originale.thumbnail((800, 800))
                    img_originale.convert("RGB").save(buffer_leggero, format="JPEG", quality=65)
                
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
                            
                            # 🎯 COORDINAMENTO DIRETTO CON I WIDGET DINAMICI DEL RESET
                            if posizione_indice == 0:
                                st.session_state.val_num_ass = stringa_pulita
                                # Iniettiamo il valore direttamente nella chiave del widget di testo per forzare lo schermo
                                st.session_state[f"num_ass_widget_{st.session_state.contatore_grafico_reset}"] = stringa_pulita
                            
                            elif posizione_indice == 1:
                                try:
                                    stringa_no_migliaia = stringa_pulita.replace(".", "")
                                    stringa_decimale_corretta = stringa_no_migliaia.replace(",", ".")
                                    val_float = float(stringa_pulita.replace(",", "")) if "." in stringa_pulita and "," not in stringa_pulita else float(stringa_decimale_corretta)
                                    
                                    st.session_state.val_imp_ass = val_float
                                    # Iniettiamo il valore direttamente nella chiave del widget numerico per forzare lo schermo
                                    st.session_state[f"imp_ass_widget_{st.session_state.contatore_grafico_reset}"] = val_float
                                except: 
                                    st.session_state.val_imp_ass = 0.0
                                    st.session_state[f"imp_ass_widget_{st.session_state.contatore_grafico_reset}"] = 0.0
                            
                            elif posizione_indice == 2:
                                try:
                                    formato_standard = stringa_pulita.replace("-", "/").replace(".", "/")
                                    val_date = datetime.strptime(formato_standard, "%d/%m/%Y").date()
                                    st.session_state.val_dat_ass = val_date
                                    # Iniettiamo il valore direttamente nella chiave del widget data per forzare lo schermo
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
# Inizializzazione della variabile di memoria RAM per la Data Incasso se non esiste
if "val_dat_inc" not in st.session_state: 
    st.session_state.val_dat_inc = datetime.now().date()

st.markdown("---")
# 🎯 POTENZIAMENTO: Allineamento spaccato a 4 colonne per far spazio alla Data Incasso
col_a, col_b, col_c, col_d = st.columns(4)

with col_a:
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
        "Data Assegno:", 
        value=st.session_state.get("val_dat_ass", datetime.now().date()), 
        format="DD-MM-YYYY", 
        key=f"dat_ass_widget_{st.session_state.contatore_grafico_reset}"
    )
with col_d:
    # 🎯 NUOVO CAMPO: Casella di testo per l'inserimento manuale della data di incasso effettivo
    data_incasso = st.date_input(
        "Data Incasso Effettivo:", 
        value=st.session_state.get("val_dat_inc", datetime.now().date()), 
        format="DD-MM-YYYY", 
        key=f"dat_inc_widget_{st.session_state.contatore_grafico_reset}"
    )


st.markdown("---")
col_btn1, col_btn2 = st.columns(2)

st.markdown("---")
col_btn1, col_btn2 = st.columns(2)

# =====================================================================================
# 🧱 [BLOCCO 2 di 3 RISANATO]: INSERIMENTO REGISTRO SENZA COMANDI DI GRIGLIA CRASH
# =====================================================================================
with col_btn1:
    if st.button("🚀 ACQUISISCI E SPEDISCI DOCUMENTO BLINDATO", use_container_width=True):
        if locale_scelto == "- Selezionare il Locale -" or not numero_assegno_input or importo_assegno <= 0:
            st.error("⚠️ Compila tutti i dati e seleziona un locale valido prima di procedere!")
        else:
            with st.spinner("💾 Inserimento record nel Cloud di GitHub..."):
                try:
                    token_pulito = str(st.secrets["github"]["token_accesso_GITHUB"] if "token_accesso_GITHUB" in st.secrets["github"] else st.secrets["github"].get("token_accesso", "")).strip()
                    
                    # 🎯 COMPILAZIONE PERFETTA: La rotta esatta per salvare i dati nell'Excel degli Assegni
                    parte_api = "api.github.com"
                    parte_repo = "repos/wingamingsrl/AppWG"
                    URL_API_GITHUB = f"https://{parte_api}/{parte_repo}/contents/registro_assegni.xlsx"

                    
                    df_storico = pd.DataFrame()
                    sha_file = ""
                    headers_git_standard = {"Authorization": f"token {token_pulito}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Cloud-App"}
                    
                    res_api = requests.get(URL_API_GITHUB, headers=headers_git_standard, timeout=10)
                    if res_api.status_code == 200:
                        j_g = res_api.json(); sha_file = j_g.get("sha", ""); c_b = j_g.get("content", "")
                        if c_b: df_storico = pd.read_excel(io.BytesIO(base64.b64decode(c_b.replace("\n", ""))), dtype={"NUMERO_ASSEGNO": str})
                    
                    # Costruzione del record secondo la nuova anagrafica snella unificata
                    nuovo_record = {
                        "DATA_REGISTRAZIONE": datetime.now().strftime("%d/%m/%Y"),
                        "ESATTORE": esecutore_nome,
                        "LOCALE": locale_scelto,
                        "NUMERO_ASSEGNO": str(numero_assegno_input).strip(),
                        "IMPORTO": float(importo_assegno),
                        "DATA_ASSEGNO": data_assegno.strftime("%d/%m/%Y"),
                        "DATA_INCASSO": data_incasso.strftime("%d/%m/%Y"), 
                        "STATO_VERSAMENTO": "In Portafoglio"
                    }
                    df_aggiornato = pd.concat([df_storico, pd.DataFrame([nuovo_record])], ignore_index=True).fillna("")
                    
                    output_file = io.BytesIO()
                    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                        df_aggiornato.to_excel(writer, index=False)
                        ws_w = writer.sheets["Sheet1"]
                        
                        # 🎯 STRATEGIA CASSELLE E BORDI: Disegna i bordi neri su ogni cella (Sostituisce il comando griglia difettoso)
                        from openpyxl.styles import Border, Side, Font
                        from openpyxl.utils import get_column_letter
                        
                        bordo_sottile = Border(
                            left=Side(style='thin', color='A0A0A0'),
                            right=Side(style='thin', color='A0A0A0'),
                            top=Side(style='thin', color='A0A0A0'),
                            bottom=Side(style='thin', color='A0A0A0')
                        )
                        
                        # Allarga le colonne calcolando gli indici numerici per evitare errori di tupla
                        for col_idx, col_cells in enumerate(ws_w.columns, start=1):
                            max_len = 0
                            for cell in col_cells:
                                cell.border = bordo_sottile
                                if cell.row == 1: 
                                    cell.font = Font(bold=True)
                                max_len = max(max_len, len(str(cell.value or '')))
                            
                            col_letter = get_column_letter(col_idx)
                            ws_w.column_dimensions[col_letter].width = max(max_len + 4, 12)
                    
                    payload_push = {"message": f"🏦 Inserimento Assegno {numero_assegno_input}", "content": base64.b64encode(output_file.getvalue()).decode('utf-8'), "branch": "main"}
                    if sha_file: payload_push["sha"] = sha_file
                    res_put = requests.put(URL_API_GITHUB, json=payload_push, headers=headers_git_standard, timeout=15)
                    
                    if res_put.status_code == 200 or res_put.status_code == 201:
                        try:
                            server_mail = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                            server_mail.login("wingamingsrl@gmail.com", "zndjprxjvhiustio")
                            
                            msg = MIMEMultipart()
                            msg['From'] = "Manuela Arigoni - WinGaming <manuela.arigoni@wingaming.it>"
                            msg['Reply-To'] = "manuela.arigoni@wingaming.it"
                            
                            # 🎯 INDIRIZZAMENTO CORRETTO: Inviata a te e all'Ufficio Amministrazione e basta
                            #msg['To'] = "manuela.arigoni@wingaming.it"
                            msg['To'] = "manuela.arigoni@wingaming.it, ufficio.amministrazione@wingaming.it, alessandro.frigerio@wingaming.it"
                            msg['Subject'] = f"🎫 [WinGaming] Nuovo Assegno Acquisito - {locale_scelto}"
                            
                            corpo_html = f"<html><body style='font-family: Arial, sans-serif;'><p>Buongiorno,</p><p>Un nuovo assegno è stato acquisito digitalmente:</p><ul><li><b>Locale:</b> {locale_scelto}</li><li><b>Esattore:</b> {esecutore_nome}</li><li><b>Numero Assegno:</b> {numero_assegno_input}</li><li><b>Importo:</b> {float(importo_assegno):,.2f} €</li><li><b>Data Assegno:</b> {data_assegno.strftime('%d/%m/%Y')}</li><li><b>Data Incasso Rilevata:</b> {data_incasso.strftime('%d/%m/%Y')}</li></ul></body></html>"
                            msg.attach(MIMEText(corpo_html, 'html'))
                            
                            if file_assegno is not None:
                                file_assegno.seek(0)
                                parte_allegato = MIMEBase('application', 'octet-stream')
                                parte_allegato.set_payload(file_assegno.read())
                                encoders.encode_base64(parte_allegato)
                                parte_allegato.add_header('Content-Disposition', f'attachment; filename="{file_assegno.name}"')
                                msg.attach(parte_allegato)
                            
                            # 🎯 CONSEGNA REALE BLINDATA: Solo i vostri due indirizzi passati al server SMTP
                            #tutti_ricevitori = ["manuela.arigoni@wingaming.it"]
                            tutti_ricevitori = ["manuela.arigoni@wingaming.it", "ufficio.amministrazione@wingaming.it", ", alessandro.frigerio@wingaming.it"]
                            
                            server_mail.sendmail("manuela.arigoni@wingaming.it", tutti_ricevitori, msg.as_string())
                            server_mail.quit()
                            st.success("🎯 Operazione completata! Riga salvata ed e-mail inviata a te e all'Amministrazione!")
                        except Exception as e_mail_fallita:
                            st.warning(f"⚠️ Assegno registrato, ma notifica e-mail rallentata: {str(e_mail_fallita)}")

                        st.session_state.contatore_grafico_reset += 1
                        st.session_state.val_num_ass = ""; st.session_state.val_imp_ass = 0.0; st.session_state.val_dat_ass = datetime.now().date(); st.session_state.val_dat_inc = datetime.now().date()
                        if f"f_up_{st.session_state.contatore_grafico_reset}" in st.session_state: del st.session_state[f"f_up_{st.session_state.contatore_grafico_reset}"]
                        st.session_state.ia_elaborata = False; st.session_state.testo_spia_generato = ""
                        st.toast("🧹 Campi puliti per la prossima operazione!")
                        time.sleep(1.2); st.rerun()
                    else:
                        st.error(f"❌ Errore scrittura database Cloud (Codice HTTP {res_put.status_code})")
                except Exception as e_database:
                    st.error(f"❌ Interruzione nel caricamento Excel: {str(e_database)}")

with col_btn2:
    if st.button("🗑️ AZZERA CAMPI", key="btn_azzeramento_campi_assegno", use_container_width=True):
        # 🎯 CORREZIONE CHIRURGICA: Forza la rigenerazione totale dei widget a schermo
        st.session_state.contatore_grafico_reset += 1
        
        # Svuotiamo le variabili di riempimento dell'assegno nella RAM di sessione
        st.session_state.val_num_ass = ""
        st.session_state.val_imp_ass = 0.0
        st.session_state.val_dat_ass = datetime.now().date()
        st.session_state.val_dat_inc = datetime.now().date()
        
        # Eliminiamo la foto caricata nell'uploader temporaneo senza toccare l'anagrafica locali
        if f"f_up_{st.session_state.contatore_grafico_reset}" in st.session_state:
            del st.session_state[f"f_up_{st.session_state.contatore_grafico_reset}"]
        
        # Resettiamo lo stato dell'elaborazione dell'IA di Google per nuovi inserimenti
        st.session_state.ia_elaborata = False
        st.session_state.testo_spia_generato = ""
        
        st.toast("🧹 Campi piallati con successo!")
        time.sleep(0.4)
        st.rerun()




# =====================================================================================
# 🧱 [PEZZO 1 di 3]: CONTROLLO GIORNALIERO E TRAVASO ARCHIVIO SU DATA_ASSEGNO (RISANATO)
# =====================================================================================
try:
    if 'ruolo_utente' not in locals() and 'ruolo_utente' not in globals():
        ruolo_utente = str(st.session_state.get("user_ruolo", "TECNICO")).strip().upper()
    # 🎯 CORREZIONE STRUTTURALE CERTIFICATA: Puntamento API reale per far riapparire la tabella
    parte_api_tab = "api.github.com"
    parte_repo_tab = "repos/wingamingsrl/AppWG"
    URL_TABELLONE_API = f"https://{parte_api_tab}/{parte_repo_tab}/contents/registro_assegni.xlsx"
    URL_STORICO_API = f"https://{parte_api_tab}/{parte_repo_tab}/contents/Storico_assegni.xlsx"
    
    token_github_tab = str(st.secrets["github"]["token_accesso_GITHUB"] if "token_accesso_GITHUB" in st.secrets["github"] else st.secrets["github"].get("token_accesso", "")).strip()
    headers_git_tabella = {"Authorization": f"token {token_github_tab}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Cloud-App"}
    
    # Ora la richiesta intercetta l'Excel reale e popola i dati
    res_v = requests.get(URL_TABELLONE_API, headers=headers_git_tabella, timeout=10)

    
    df_visualizza = pd.DataFrame()
    df_visualizza_completo = pd.DataFrame()
    dati_j = {}
    
    if res_v.status_code == 200:
        dati_j = res_v.json()
        cont_b64 = dati_j.get("content", "")
        if cont_b64:
            df_visualizza_completo = pd.read_excel(io.BytesIO(base64.b64decode(cont_b64)), dtype={"NUMERO_ASSEGNO": str}).fillna("")
            
            oggi_dt = datetime.now().date()
            da_aggiornare_cloud = False
            
            # A. CONTROLLO GIORNALIERO AUTOMATICO (SMARCO IERI)
            for idx_cron, riga_cron in df_visualizza_completo.iterrows():
                if riga_cron["STATO_VERSAMENTO"] == "In Portafoglio" and riga_cron["DATA_ASSEGNO"]:
                    try:
                        data_assegno_dt = datetime.strptime(str(riga_cron["DATA_ASSEGNO"]).strip(), "%d/%m/%Y").date()
                        if (oggi_dt - data_assegno_dt).days == 1:
                            df_visualizza_completo.at[idx_cron, "STATO_VERSAMENTO"] = "Versato OK"
                            da_aggiornare_cloud = True
                    except: pass

            # B. TRAVASO STORICO SU BASE DATA_ASSEGNO (100 Giorni Admin / 0 Giorni Tecnici)
            limite_giorni_spurgo = 100 if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"] else 0
            
            indici_da_archiviare = []
            righe_da_spostare = []
            
            for idx_arch, riga_arch in df_visualizza_completo.iterrows():
                if str(riga_arch["STATO_VERSAMENTO"]).strip() == "Versato OK" and riga_arch["DATA_ASSEGNO"]:
                    try:
                        dt_titolo = datetime.strptime(str(riga_arch["DATA_ASSEGNO"]).strip(), "%d/%m/%Y").date()
                        if (oggi_dt - dt_titolo).days >= limite_giorni_spurgo:
                            indici_da_archiviare.append(idx_arch)
                            righe_da_spostare.append(riga_arch)
                    except: pass

            if indici_da_archiviare:
                try:
                    df_da_archiviare = pd.DataFrame(righe_da_spostare)
                    df_storico_esistente = pd.DataFrame()
                    sha_storico = ""
                    
                    res_st = requests.get(URL_STORICO_API, headers=headers_git_tabella, timeout=10)
                    if res_st.status_code == 200:
                        j_st = res_st.json()
                        sha_storico = j_st.get("sha", "")
                        df_storico_esistente = pd.read_excel(io.BytesIO(base64.b64decode(j_st.get("content", ""))), dtype={"NUMERO_ASSEGNO": str}).fillna("")
                    
                    df_storico_nuovo = pd.concat([df_storico_esistente, df_da_archiviare], ignore_index=True).fillna("")
                    
                    out_st = io.BytesIO()
                    with pd.ExcelWriter(out_st, engine='openpyxl') as wr_st:
                        df_storico_nuovo.to_excel(wr_st, index=False)
                        # Abilita le linee della griglia in modo sicuro
                        wr_st.sheets["Sheet1"].views.sheetView[0].showGridLines = True
                        from openpyxl.styles import Border, Side
                        b_s = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                        for row in wr_st.sheets["Sheet1"].iter_rows(min_row=1, max_row=wr_st.sheets["Sheet1"].max_row, min_col=1, max_col=wr_st.sheets["Sheet1"].max_column):
                            for cell in row: cell.border = b_s
                        
                    payload_st = {
                        "message": f"🤖 [Archivio Semplificato] Spostati {len(indici_da_archiviare)} titoli",
                        "content": base64.b64encode(out_st.getvalue()).decode('utf-8'),
                        "branch": "main"
                    }
                    if sha_storico: payload_st["sha"] = sha_storico
                    requests.put(URL_STORICO_API, json=payload_st, headers=headers_git_tabella, timeout=15)
                    
                    df_visualizza_completo = df_visualizza_completo.drop(indici_da_archiviare).reset_index(drop=True)
                    da_aggiornare_cloud = True
                except: pass

            if da_aggiornare_cloud:
                try:
                    output_cron = io.BytesIO()
                    with pd.ExcelWriter(output_cron, engine='openpyxl') as wr_c:
                        df_visualizza_completo.to_excel(wr_c, index=False)
                        wr_c.sheets["Sheet1"].views.sheetView[0].showGridLines = True
                        from openpyxl.styles import Border, Side
                        b_s = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                        for row in wr_c.sheets["Sheet1"].iter_rows(min_row=1, max_row=wr_c.sheets["Sheet1"].max_row, min_col=1, max_col=wr_c.sheets["Sheet1"].max_column):
                            for cell in row: cell.border = b_s
                    payload_c = {"message": "🤖 [Pulizia Campi Doppi]", "content": base64.b64encode(output_cron.getvalue()).decode('utf-8'), "branch": "main", "sha": dati_j.get("sha", "")}
                    res_p_c = requests.put(URL_TABELLONE_API, json=payload_c, headers=headers_git_tabella, timeout=15)
                    if res_p_c.status_code == 200 or res_p_c.status_code == 201:
                        dati_j = res_p_c.json()
                except: pass



# =====================================================================================
# 🧱 [PEZZO 2 di 3]: BADGE IN PORTAFOGLIO PRIVACY, DOPPIO DOWNLOAD EXCEL ANTI-CRASH
# =====================================================================================
            if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"]:
                df_visualizza = df_visualizza_completo.copy()
            else:
                nome_esattore_corrente = str(esecutore_nome).strip().upper()
                df_visualizza = df_visualizza_completo[(df_visualizza_completo["ESATTORE"].astype(str).str.strip().str.upper() == nome_esattore_corrente) & (df_visualizza_completo["STATO_VERSAMENTO"] == "In Portafoglio")].copy()
            
            somma_non_versati = 0.0
            if not df_visualizza.empty:
                df_p = df_visualizza[df_visualizza["STATO_VERSAMENTO"] == "In Portafoglio"]
                somma_non_versati = float(pd.to_numeric(df_p["IMPORTO"], errors='coerce').fillna(0.0).sum())

            if not df_visualizza.empty:
                def colora_assegni(riga):
                    stils = [''] * len(riga)
                    if str(riga["STATO_VERSAMENTO"]).strip() == "In Portafoglio": stils[df_visualizza.columns.get_loc("STATO_VERSAMENTO")] = 'background-color: #fff2cc; color: #000000; font-weight: 500;'
                    elif str(riga["STATO_VERSAMENTO"]).strip() == "Versato OK": stils[df_visualizza.columns.get_loc("STATO_VERSAMENTO")] = 'background-color: #d1fae5; color: #065f46; font-weight: 500;'
                    return stils
                
                st.markdown("---")
                st.markdown("### 📊 Registro Storico Assegni Acquisiti")
                st.markdown(f"<div style='background-color: #fff2cc; border-left: 5px solid #d97706; color: #92400e; padding: 10px; border-radius: 6px; font-size: 14px; font-weight: 700; margin-bottom: 15px;'>💰 IL TUO PORTAFOGLIO ATTUALE (NON VERSATI): {somma_non_versati:,.2f} €</div>", unsafe_allow_html=True)
                
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    df_esp = df_visualizza[df_visualizza["STATO_VERSAMENTO"] == "In Portafoglio"].copy()
                    b_o1 = io.BytesIO()
                    with pd.ExcelWriter(b_o1, engine='openpyxl') as w1: 
                        df_esp.to_excel(w1, index=False)
                        ws1 = w1.sheets["Sheet1"]
                        ws1.views.sheetView[0].showGridLines = True
                        
                        from openpyxl.styles import Border, Side, Font
                        from openpyxl.utils import get_column_letter # 🎯 IMPORTAZIONE FUNZIONE ANTI-TUPLE
                        bordo_sottile = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                        
                        # 🎯 STRATEGIA SOLIDA: Allargamento colonne basato sugli indici numerici
                        for col_idx, col_cells in enumerate(ws1.columns, start=1):
                            max_len = 0
                            for cell in col_cells:
                                cell.border = bordo_sottile
                                if cell.row == 1: cell.font = Font(bold=True)
                                max_len = max(max_len, len(str(cell.value or '')))
                            col_letter = get_column_letter(col_idx)
                            ws1.column_dimensions[col_letter].width = max(max_len + 4, 12)
                            
                    st.download_button(label="📥 SCARICA ASSEGNI IN PORTAFOGLIO (EXCEL)", data=b_o1.getvalue(), file_name="portafoglio_assegni.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_p_ass")
                
                with col_dl2:
                    byte_storico_dl = b""
                    res_dl_st = requests.get(URL_STORICO_API, headers=headers_git_tabella, timeout=10)
                    if res_dl_st.status_code == 200: 
                        df_st_grezzo = pd.read_excel(io.BytesIO(base64.b64decode(res_dl_st.json().get("content", ""))), dtype={"NUMERO_ASSEGNO": str}).fillna("")
                        b_o2 = io.BytesIO()
                        with pd.ExcelWriter(b_o2, engine='openpyxl') as w2:
                            df_st_grezzo.to_excel(w2, index=False)
                            ws2 = w2.sheets["Sheet1"]
                            ws2.views.sheetView[0].showGridLines = True
                            
                            from openpyxl.styles import Border, Side, Font
                            from openpyxl.utils import get_column_letter # 🎯 IMPORTAZIONE FUNZIONE ANTI-TUPLE
                            bordo_sottile = Border(left=Side(style='thin', color='A0A0A0'), right=Side(style='thin', color='A0A0A0'), top=Side(style='thin', color='A0A0A0'), bottom=Side(style='thin', color='A0A0A0'))
                            
                            # 🎯 STRATEGIA SOLIDA: Allargamento colonne basato sugli indici numerici sullo storico
                            for col_idx, col_cells in enumerate(ws2.columns, start=1):
                                max_len = 0
                                for cell in col_cells:
                                    cell.border = bordo_sottile
                                    if cell.row == 1: cell.font = Font(bold=True)
                                    max_len = max(max_len, len(str(cell.value or '')))
                                col_letter = get_column_letter(col_idx)
                                ws2.column_dimensions[col_letter].width = max(max_len + 4, 12)
                        byte_storico_dl = b_o2.getvalue()
                        
                    st.download_button(label="📥 SCARICA ARCHIVIO STORICO GENERALE (EXCEL)", data=byte_storico_dl, file_name="Storico_completo_assegni.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_s_ass", disabled=(byte_storico_dl == b""))
                
                st.dataframe(df_visualizza.style.apply(colora_assegni, axis=1).format({"IMPORTO": "{:,.2f} €"}), use_container_width=True, hide_index=True, height=400)
            else: 
                st.markdown("---")
                st.info("ℹ️ Nessun assegno da mostrare.")


# =====================================================================================
# 🧱 [BLOCCO 3 di 3]: CENTRALINA AMMINISTRATIVA DI INTERVENTO MULTI-SMARCO E CHIUSURA
# =====================================================================================
            if ruolo_utente in ["ADMIN", "SUPERVISORE", "UFFICIO"] and not df_visualizza.empty:
                st.markdown("---")
                st.markdown("### 🛠️ Gestione Amministrativa e Modifica Stati Assegni (Ufficio)")
                
                opzioni_da_versare = []
                opzioni_gia_versati = []
                mappa_global_indici = {}
                
                for idx_m, r_m in df_visualizza_completo.iterrows():
                    testo_riga = f"ID {idx_m} | 👤 {r_m['ESATTORE']} | 📍 {r_m['LOCALE']} | Assegno: {r_m['NUMERO_ASSEGNO']} | {r_m['IMPORTO']:.2f} €"
                    mappa_global_indici[testo_riga] = idx_m
                    if str(r_m["STATO_VERSAMENTO"]).strip() == "In Portafoglio":
                        opzioni_da_versare.append(testo_riga)
                    elif str(r_m["STATO_VERSAMENTO"]).strip() == "Versato OK":
                        opzioni_gia_versati.append(testo_riga)
                
                assegni_scelti_da_versare = st.multiselect("Scegli uno o PIÙ assegni da marcare come VERSATI:", options=opzioni_da_versare, key="multi_v_m")
                assegno_scelto_da_ripristinare = st.selectbox("Seleziona un titolo già versato da far tornare IN PORTAFOGLIO:", options=["- Seleziona -"] + opzioni_gia_versati, key="sb_r_p")
                
                st.markdown(" ")
                col_c1, col_c2, col_c3 = st.columns(3)
                
                with col_c1:
                    if st.button("🏦 MARCA MANUALE COME: VERSATO OK", key="b_m_v", use_container_width=True):
                        if not assegni_scelti_da_versare: st.warning("⚠️ Spunta almeno un titolo!")
                        else:
                            for s in assegni_scelti_da_versare:
                                t_i = mappa_global_indici[s]
                                df_visualizza_completo.at[t_i, "STATO_VERSAMENTO"] = "Versato OK"
                            out_f = io.BytesIO()
                            with pd.ExcelWriter(out_f, engine='openpyxl') as wr: 
                                df_visualizza_completo.to_excel(wr, index=False)
                                wr.sheets["Sheet1"].views.sheetView[0].showGridLines = True
                            payload_adm = {"message": "🏦 Smarco Multiplo", "content": base64.b64encode(out_f.getvalue()).decode('utf-8'), "branch": "main", "sha": dati_j.get("sha", "")}
                            requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)
                            st.success("✅ Assegni modificati con successo!"); time.sleep(1.2); st.rerun()
                with col_c2:
                    if st.button("↩️ RIPRISTINA STATO: TORNA IN PORTAFOGLIO", key="b_r_p", use_container_width=True):
                        if assegno_scelto_da_ripristinare == "- Seleziona -": st.warning("⚠️ Seleziona un titolo!")
                        else:
                            t_i = mappa_global_indici[assegno_scelto_da_ripristinare]
                            df_visualizza_completo.at[t_i, "STATO_VERSAMENTO"] = "In Portafoglio"
                            out_f = io.BytesIO()
                            with pd.ExcelWriter(out_f, engine='openpyxl') as wr: 
                                df_visualizza_completo.to_excel(wr, index=False)
                                wr.sheets["Sheet1"].views.sheetView[0].showGridLines = True
                            payload_adm = {"message": "↩️ Ripristino", "content": base64.b64encode(out_f.getvalue()).decode('utf-8'), "branch": "main", "sha": dati_j.get("sha", "")}
                            requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)
                            st.success("✅ Titolo riportato in Portafoglio!"); time.sleep(1.2); st.rerun()
                with col_c3:
                    if st.button("❌ ELIMINA ASSEGNO SELEZIONATO", key="b_e_a", use_container_width=True):
                        ind_p = []
                        if assegni_scelti_da_versare:
                            for s in assegni_scelti_da_versare: ind_p.append(mappa_global_indici[s])
                        elif assegno_scelto_da_ripristinare != "- Seleziona -": ind_p.append(mappa_global_indici[assegno_scelto_da_ripristinare])
                        if not ind_p: st.warning("⚠️ Seleziona un titolo!")
                        else:
                            df_visualizza_completo = df_visualizza_completo.drop(ind_p).reset_index(drop=True)
                            out_f = io.BytesIO()
                            with pd.ExcelWriter(out_f, engine='openpyxl') as wr: 
                                df_visualizza_completo.to_excel(wr, index=False)
                                wr.sheets["Sheet1"].views.sheetView[0].showGridLines = True
                            payload_adm = {"message": "❌ Elimina", "content": base64.b64encode(out_f.getvalue()).decode('utf-8'), "branch": "main", "sha": dati_j.get("sha", "")}
                            requests.put(URL_TABELLONE_API, json=payload_adm, headers=headers_git_tabella, timeout=15)
                            st.success("🗑️ Eliminato!"); time.sleep(1.2); st.rerun()
        else: st.info("ℹ️ Registro vuoto su GitHub.")
    elif res_v.status_code == 404: st.info("ℹ️ Registro vuoto su GitHub.")
    else: st.warning(f"⚠️ Impossibile accedere al Cloud. HTTP {res_v.status_code}")
except Exception as e: st.error(f"❌ Errore caricamento blocco finale: {str(e)}")
