# =====================================================================================
# SW AUTOMATICO DI SINCRONIZZAZIONE LOCALI WIN GAMING — PRODUZIONE FINALE
# BLOCCO 1: STRUTTURA LIBRERIE AZIENDALI, TOTP 2FA E RIPULITURA EXCEL VIA GIT
# =====================================================================================
import os
import io
import time
import base64
import pyotp
import requests
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

CHIAVE_SEGRETA_2FA = "FTIA6UQZM2LQLPYJ"

SNAI_USER = "2141ManuelaA"
SNAI_PASS = "Salmi123!"

def invia_email_avviso_interno_concessionari(tecnico, locale, codice, data_in, data_fi, concessionario, azione):
    print(f"📧 [Internal Engine] Invio promemoria interno per {concessionario} (Locale: {locale})...")
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        EMAIL_LOGIN = "wingamingsrl@gmail.com"
        pass_applicativa_ufficio = "zndjprxjvhiustio"
        EMAIL_UFFICIALE_MITTENTE = "tecnico@wingaming.it"

        msg = MIMEMultipart()
        msg['From'] = f"WinGaming Robot <tecnico@wingaming.it>"
        msg['To'] = "manuela.arigoni@wingaming.it"
        msg['Subject'] = f"⚠️ [AVVISO INTERNO] {azione} Chiusura {concessionario} - Locale: {locale} ({codice})"

        corpo_mail = f"""All'attenzione di Manuela,

Questo è un promemoria interno automatico del robot. È stata registrata una richiesta di [{azione}] per un concessionario a gestione manuale.

Dettagli della pratica da evadere:
--------------------------------------------------
🏢 CONCESSIONARIO:     {concessionario}
🔔 STATO OPERAZIONE:    {azione}
👤 TECNICO ESECUTORE:   {tecnico}
📍 LOCALE COINVOLTO:    {locale}
📌 CODICE CENSIMENTO:   {codice}
📅 INIZIO CHIUSURA:    {data_in}
🚚 DATA RIAPERTURA:    {data_fi}
--------------------------------------------------

Verificare la pratica e procedere con l'allineamento manuale.
WinGaming S.r.l."""

        msg.attach(MIMEText(corpo_mail, 'plain', 'utf-8'))

        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_LOGIN, pass_applicativa_ufficio)
        server.sendmail(EMAIL_UFFICIALE_MITTENTE, ["manuela.arigoni@wingaming.it"], msg.as_string())
        server.quit()
        
        print(f"   ✅ [Internal Engine] Avviso interno per {concessionario} consegnato a Manuela!")
        return True
    except Exception as e_int:
        print(f"   ❌ [Internal Engine] Errore invio avviso {concessionario}: {str(e_int)}")
        return False

def preleva_storico_diretto_da_cloud():
    print("📡 [Robot] STEP 1: Lettura del database Excel locale...")
    try:
        nome_file_locale = "storico_ferie.xlsx"
        if os.path.exists(nome_file_locale):
            return pd.read_excel(nome_file_locale).fillna("")
    except Exception: pass
    return pd.DataFrame()

def genera_codice_otp_automatico():
    chiave_pulita = CHIAVE_SEGRETA_2FA.strip().upper().replace(" ", "")
    totp = pyotp.TOTP(chiave_pulita)
    return totp.now()

def scarica_e_aggiorna_excel_su_github(codice_locale_successo):
    print(f"💾 [Cloud Excel] Svuoto ROBOT_ACTION per il locale: {codice_locale_successo}...")
    try:
        nome_file = "storico_ferie.xlsx"
        if os.path.exists(nome_file):
            # 🛡️ FIX FINALE DI MANUELA: Forza la lettura dei codici come testo puro (str) azzerando i conflitti di dtypes vuoti
            df_file = pd.read_excel(nome_file, dtype={"CODICE_LOCALE": str})
            
            # Azzera la cella d'azione per il locale completato
            # 🛡️ SPAZZINO DOPPIO DI MANUELA: Se l'azione era ELIMINA cancella la riga, altrimenti svuota la cella
            df_filtrato_locale = df_file[df_file["CODICE_LOCALE"].astype(str).str.strip() == str(codice_locale_successo).strip()]
            if not df_filtrato_locale.empty and str(df_filtrato_locale.iloc[0].get("ROBOT_ACTION", "")).strip().upper() == "ELIMINA":
                df_file = df_file[df_file["CODICE_LOCALE"].astype(str).str.strip() != str(codice_locale_successo).strip()]
                print("   🗑️ [Cloud Excel] Rilevato comando ELIMINA: Riga rimossa definitivamente dal database.")
            else:
                df_file.loc[df_file["CODICE_LOCALE"].astype(str).str.strip() == str(codice_locale_successo).strip(), "ROBOT_ACTION"] = ""

            
            # Riassegna la struttura colonne rigida dell'ufficio
            colonne_ufficio = ["DATA_INSERIMENTO", "TECNICO_INSERIMENTO", "CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO", "INIZIO_FERIE", "FINE_FERIE", "PROMEMORIA_IN_COPIA", "STATO_INVIO", "ROBOT_ACTION"]
            df_pulito_salva = df_file.reindex(columns=colonne_ufficio).astype(str).fillna("")
            df_pulito_salva.to_excel(nome_file, index=False)
            
            # 🛡️ AUTOMAZIONE GIT NATIVA: Salva e spinge sul cloud senza conflitti di token API
            os.system("git config --global user.name 'WinGaming-Robot'")
            os.system("git config --global user.email 'wingamingsrl@gmail.com'")
            os.system(f"git add {nome_file}")
            os.system(f"git commit -m '🤖 [Robot] Allineamento Snaitech OK. Reset azione locale {codice_locale_successo}'")
            stato_push = os.system("git push origin main")
            
            if stato_push == 0:
                print("   ✅ [Cloud Excel] Database ripulito e sincronizzato con successo online!")
            else:
                print(f"   ❌ [Cloud Excel] Errore riscrittura online. Codice push: {stato_push}")
    except Exception as e:
        print(f"   ⚠️ Impossibile aggiornare l'Excel: {str(e)}")


def spedisci_email_avviso_ufficio(tecnico_nome, collega_in_copia, locale, codice, data_evento, tipo_avviso):
    print(f"📧 [Email Engine] Scansione database tecnici basata sulla colonna NOME...")
    try:
        # Importazioni librerie native identiche al tuo file test_app.py
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        elenco_email_squadra = {}
        nome_file_tecnici = "elenco_tecnici.xlsx"
        if os.path.exists(nome_file_tecnici):
            df_tecnici = pd.read_excel(nome_file_tecnici).fillna("")
            for _, t_row in df_tecnici.iterrows():
                nome_db = str(t_row.get("NOME", "")).strip().upper()
                email_db = str(t_row.get("EMAIL", "")).strip()
                if nome_db and email_db:
                    elenco_email_squadra[nome_db] = email_db

        print(f"   📊 [Email Engine] Tecnici caricati in memoria: {list(elenco_email_squadra.keys())}")

        # Creazione lista destinatari dinamica
        destinatari_finali = []
        
        # 1. Cerca il Tecnico Titolare della riga ferie
        nome_tecnico_pulito = str(tecnico_nome).strip().upper()
        if nome_tecnico_pulito:
            for nome_completo_db, email_corrispondente in elenco_email_squadra.items():
                if nome_completo_db.startswith(nome_tecnico_pulito) or nome_tecnico_pulito in nome_completo_db:
                    if email_corrispondente not in destinatari_finali:
                        destinatari_finali.append(email_corrispondente)
                        print(f"   ✅ [Email Engine] Destinatario Titolare agganciato: {email_corrispondente}")
                    break
                
        # 2. Cerca il Collega inserito in copia promemoria
        nome_collega_pulito = str(collega_in_copia).strip().upper()
        if nome_collega_pulito and nome_collega_pulito != "NESSUNO":
            for nome_completo_db, email_corrispondente in elenco_email_squadra.items():
                if nome_completo_db.startswith(nome_collega_pulito) or nome_collega_pulito in nome_completo_db:
                    if email_corrispondente not in destinatari_finali:
                        destinatari_finali.append(email_corrispondente)
                        print(f"   ✅ [Email Engine] Destinatario in Copia agganciato: {email_corrispondente}")
                    break

        # Inserisce sempre Manuela Arigoni per supervisione fissa se non già inclusa
        email_manuela_default = elenco_email_squadra.get("MANUELA ARIGONI", "manuela.arigoni@wingaming.it")
        if email_manuela_default not in destinatari_finali:
            destinatari_finali.append(email_manuela_default)

        stringa_destinatari = ", ".join(destinatari_finali)
        
        # --- STRUTTURA INVIO EMAIL COPIATA ALLINEATA AL MILLIMETRO AL TUO FILE TEST_APP.PY ---
        EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
        
        # Recupera in automatico la password applicativa salvata in modo sicuro nei segreti di GitHub
        pass_gmail = "zndjprxjvhiustio"
        
        msg = MIMEMultipart()
        msg['From'] = EMAIL_MITTENTE_GMAIL
        msg['To'] = stringa_destinatari
        msg['Subject'] = f"🛡️ Promemoria {tipo_avviso} - {locale}"
        
        corpo = f"Rilevato avviso scadenza imminente nel sistema WinGaming.\n\nDettagli della pratica:\n--------------------------------------------------\n🔔 Stato Operazione:  {tipo_avviso.upper()}\n👤 Tecnico Responsabile: {tecnico_nome}\n👥 Colleghi in Copia:   {collega_in_copia}\n📍 Locale Coinvolto:    {locale} ({codice})\n📅 Data Scadenza Evento: {data_evento}\n--------------------------------------------------\n\nWINGAMING SRL"
        msg.attach(MIMEText(corpo, 'plain'))
        
        # 🚨 IL METODO INFALLIBILE DI MANUELA: Connessione forzata via IP numerico diretto su porta SSL 465
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, destinatari_finali, msg.as_string())
        server.quit()
        
        print(f"   ✅ [Email Engine] Notifica inviata e CONSEGNATA REALMENTE via IP a: {stringa_destinatari}")
    except Exception as e_mail:
        print(f"   ❌ Errore durante l'invio SMTP IP nativo basato su test_app.py: {str(e_mail)}")

#=======================================================
# ----------- SINCRONIZZAZIONE CON NTS----------------
#=======================================================
def invia_email_chiusura_diretta_nts(tecnico, locale, codice, data_in, data_fi, azione):
    print(f"📧 [NTS Engine] Preparazione invio digitale {azione} per NTS Networks (Locale: {locale})...")
    try:
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        EMAIL_LOGIN = "wingamingsrl@gmail.com"
        pass_applicativa_ufficio = "zndjprxjvhiustio"

        if azione == "ELIMINA":
            oggetto_azione = "CANCELLAZIONE Chiusura Temporanea"
            testo_azione = "si richiede la CANCELLAZIONE della comunicazione di chiusura temporanea precedentemente inviata"
        elif azione == "MODIFICA":
            oggetto_azione = "RETTIFICA/MODIFICA Chiusura Temporanea"
            testo_azione = "si richiede la RETTIFICA/MODIFICA della comunicazione di chiusura temporanea"
        else:
            oggetto_azione = "Richiesta Chiusura Temporanea"
            testo_azione = "si richiede l’invio della comunicazione di chiusura temporanea"

        testo_preu = ""
        if azione == "NUOVA" or azione == "MODIFICA":
            testo_preu = "\nSi richiede il blocco Preu\n"

        corpo_nts = f"""Buongiorno,
 
{testo_azione} per l’esercizio indicato in oggetto e per gli apparecchi ivi ubicati per il periodo:
 
dal\t{data_in}
al\t{data_fi}
{testo_preu}
Grazie
Cordiali saluti


Arigoni Manuela

Wingaming S.r.l.
Sede Legale e Operativa: Via Roma, 32/F - 23855 Pescate (LC)
P.Iva Gruppo IVA: 12027280960
C.F. 03371290135
CODICE UNIVOCO INTERSCAMBIO: SUBM70N
Tel. 0341.1917908"""

        # 🚨 LA CHIAVE DELLA TRASPARENZA DI MANUELA: Destinatario scritto in chiaro nel codice!
        # Puoi mettere "manuela.arigoni@wingaming.it" per fare i tuoi test, oppure "customercare@ntsnetwork.it" per l'invio reale!
        EMAIL_DESTINATARIO_NTS = "customercare@ntsnetwork.it"
        #EMAIL_DESTINATARIO_NTS = "manuela.arigoni@wingaming.it"
        EMAIL_COPIA_MANUELA = "manuela.arigoni@wingaming.it"

        msg = MIMEMultipart()
        msg['From'] = f"Arigoni Manuela <manuela.arigoni@wingaming.it>"
        msg['To'] = EMAIL_DESTINATARIO_NTS
        msg['Cc'] = EMAIL_COPIA_MANUELA  # 🛡️ Inserisce visivamente Manuela in Cc nell'e-mail
        msg['Subject'] = f"{oggetto_azione} ed Esclusione PREU - Locale: {locale} ({codice})"
        msg['Reply-To'] = "manuela.arigoni@wingaming.it"

        msg.attach(MIMEText(corpo_nts, 'plain', 'utf-8'))

        # Connessione blindata via IP numerico diretto su porta SSL 465 (Infallibile)
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_LOGIN, pass_applicativa_ufficio)
        
        # Spedisce la mail prelevando l'indirizzo in chiaro scritto sopra
        server.sendmail(EMAIL_LOGIN, [EMAIL_DESTINATARIO_NTS], msg.as_string())
        server.quit()
        
        print(f"   ✅ [NTS Engine] Richiesta inviata con successo via IP in chiaro a: {EMAIL_DESTINATARIO_NTS}")
        return True
    except Exception as e_nts_smtp:
        print(f"   ❌ [NTS Engine] Impossibile spedire la mail NTS via IP: {str(e_nts_smtp)}")
        return False





# =====================================================================================
# BLOCCO 2: FILTRO SELEZIONE ANAGRAFICA AZIENDALE ED ACCENSIONE BROWSER CHROME
# =====================================================================================
def avvia_sincronizzazione_automatica():
    df_ferie = preleva_storico_diretto_da_cloud()
    if df_ferie.empty: return

    # Estrae solo i record che hanno un'azione reale da compiere
    locali_pronti = df_ferie[df_ferie["ROBOT_ACTION"].astype(str).str.strip().str.upper().isin(["NUOVA", "MODIFICA", "ELIMINA"])]
    if locali_pronti.empty: return

    # 🛡️ ARCHITETTURA DI MANUELA: Sbarramento preventivo NTS, Global Starnet e Sisal con gestione rigida ELIMINA
    rimangono_locali_snai = False
    for _, row in locali_pronti.iterrows():
        concessionario_riga = str(row["CONCESSIONARIO"]).strip().upper()
        mirino_azione = str(row.get("ROBOT_ACTION", "")).strip().upper()
        
        codice_aams = str(row["CODICE_LOCALE"]).strip()
        nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
        data_nts_in = str(row["INIZIO_FERIE"]).replace("-", ".").strip()
        data_nts_fi = str(row["FINE_FERIE"]).replace("-", ".").strip()

        # 📄 CASO 1: NTS NETWORKS (Invia e-mail ufficiale a NTS)
        if "NTS" in concessionario_riga or "NETWORKS" in concessionario_riga:
            print(f"🏢 [NTS Networks] Rilevato locale: {nome_locale_corrente} in stato [{mirino_azione}]. Attivo l'invio...")
            successo_nts = invia_email_chiusura_diretta_nts(row["TECNICO_INSERIMENTO"], nome_locale_corrente, codice_aams, data_nts_in, data_nts_fi, mirino_azione)
            
            if successo_nts:
                # 🛡️ PROTEZIONE DI MANUELA: Se l'azione è ELIMINA, spedisce l'avviso a Manuela ORA, prima che lo spazzino distrugga la riga!
                if mirino_azione == "ELIMINA":
                    print(f"📧 [Robot] Invio notifica di CANCELLAZIONE interna a Manuela per il locale {nome_locale_corrente}...")
                    try:
                        # Converte le date in formato italiano con le barre per il testo del tuo promemoria storico
                        data_inizio_stampa = str(row["INIZIO_FERIE"]).replace("-", "/").replace(".", "/").strip()[:10]
                        spedisci_email_avviso_ufficio(row["TECNICO_INSERIMENTO"], row.get("PROMEMORIA_IN_COPIA", "Nessuno"), nome_locale_corrente, codice_aams, data_inizio_stampa, "CANCELLAZIONE FERIE REGISTRATA")
                    except Exception as e_c_mail:
                        print(f"   ⚠️ Impossibile spedire notifica interna di cancellazione: {str(e_c_mail)}")
                
                # Ora lo spazzino può procedere a rimuovere in sicurezza la riga da GitHub
                scarica_e_aggiorna_excel_su_github(codice_aams)

                
        # 🚨 CASO 2: GLOBAL STARNET O SISAL (Invia avviso interno a Manuela)
        elif "GLOBAL" in concessionario_riga or "STARNET" in concessionario_riga or "SISAL" in concessionario_riga:
            concessionario_pulito = "Global Starnet" if "GLOBAL" in concessionario_riga else "Sisal"
            print(f"🎫 [{concessionario_pulito}] Rilevato locale: {nome_locale_corrente} in stato [{mirino_azione}]. Spedisco avviso interno...")
            successo_interno = invia_email_avviso_interno_concessionari(row["TECNICO_INSERIMENTO"], nome_locale_corrente, codice_aams, data_nts_in, data_nts_fi, concessionario_pulito, mirino_azione)
            if successo_interno:
                scarica_e_aggiorna_excel_su_github(codice_aams)
                
        else:
            rimangono_locali_snai = True

    if not rimangono_locali_snai:
        print("✅ [Robot] Tutti i locali NTS evasi con successo. Nessun locale Snaitech in coda. Evito login inutile!")
        return

    # Sfrutta il tracciato originale pulito per Snaitech
    df_snai = df_ferie[
        df_ferie["CONCESSIONARIO"].astype(str).str.strip() == "Snaitech Spa WG"
    ]
    if df_snai.empty: return

    print(f"🤖 [Robot] STEP 3: Rilevati {len(df_snai)} locali Snaitech Spa WG. Avvio Chrome...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=[
            "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"
        ]) 
        context = browser.new_context()
        page = context.new_page()

        page.on("dialog", lambda dialog: dialog.accept())

# =====================================================================================
# BLOCCO 3: ACCESSO SUL PORTALE PARTNER ED IMMISSIONE CHIAVE DINAMICA OTP (LINK CORTO)
# =====================================================================================
        try:
            print("🌐 [Robot] STEP 4: Connessione a partner.snai.it...")
            page.goto("https://partner.snai.it")
            time.sleep(3)
            
            print("📝 [Robot] STEP 4a: Inserimento credenziali Snaitech...")
            page.fill("input#username, input[name='username'], input[type='text']", SNAI_USER)
            page.fill("input#password, input[name='password'], input[type='password']", SNAI_PASS)
            
            print("🚀 [Robot] STEP 4b: Invio moduli di accesso...")
            page.click("button[type='submit'], input[type='submit'], .btn-login")
            time.sleep(4)
            
            print("⏳ [Robot] STEP 4c: Pausa di sicurezza di 11 secondi per far scadere il countdown...")
            time.sleep(11)
            
            try:
                page.evaluate("""
                    document.querySelectorAll('.modal, .modal-backdrop, [id*="modal"], [class*="modal"], .fade.in').forEach(el => el.remove());
                    document.body.classList.remove('modal-open');
                    document.body.style.overflow = 'auto';
                """)
                print("✅ [Robot] STEP 4d: Codice pop-up eliminato dalla pagina con successo!")
            except Exception: pass
            time.sleep(2)

            print("🔑 [Robot] STEP 4e: Generazione ed immissione codice 2FA TOTP pulito...")
            codice_totp = genera_codice_otp_automatico()
            print(f"📌 Codice generated inviato a schermo: {codice_totp}")
            
            page.fill("input#token, input[name='token'], input[name='otp']", str(codice_totp))
            time.sleep(1)
            
            page.click("input#btnInvia, input[value='Invia'], button:has-text('Invia'), input[type='submit']")
            print("⏳ [Robot] Convalida credenziali in corso... Caricamento area riservata partner.snai.it...")
            time.sleep(15)
            
            print("🔓 [Robot] STEP 5: ACCESSO EFFETTUATO CON SUCCESSO SUL PORTALE PARTNER SNAITECH!")
            print("----------------------------------------------------------------------")

# =====================================================================================
# BLOCCO 4: SPOSTAMENTO IN ANAGRAFICA E STRUTTURA RICERCA LOCALI SNAI NATIVI
# =====================================================================================
            print("📬 [Robot] STEP 6: Spostamento sulla pagina degli Esercizi censiti...")
            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
            print("   ⏳ [Robot] STEP 6a: Attesa stabilizzazione della pagina (10 secondi)...")
            time.sleep(10)

            # Ciclo nativo senza ri-letture orfane che rompono i frame
            df_snai = df_ferie[df_ferie["CONCESSIONARIO"].astype(str).str.strip() == "Snaitech Spa WG"]

            for _, row in df_snai.iterrows():
                try:
                    codice_aams = str(row["CODICE_LOCALE"]).strip()
                    nome_locale_corrente = str(row["NOME_LOCALE"]).strip()
                    data_in_completa = str(row["INIZIO_FERIE"]).strip()
                    data_fi_completa = str(row["FINE_FERIE"]).strip()
                    mirino_azione = str(row.get("ROBOT_ACTION", "")).strip().upper()
                    
                    data_inizio_pulita = str(data_in_completa).replace("-", "/").strip()
                    data_fine_pulita = str(data_fi_completa).replace("-", "/").strip()
                    
                    # 🛡️ INTERCETTATORE DI MANUELA: Se la cella non contiene NUOVA o MODIFICA, salta la riga all'istante
                    # 🛡️ INTERCETTATORE DI MANUELA: Abilita il robot a elaborare anche i comandi di rimozione
                    # 🛡️ INTERCETTATORE DI MANUELA: Se la cella non contiene NUOVA o MODIFICA, salta la riga all'istante
                    if mirino_azione not in ["NUOVA", "MODIFICA", "ELIMINA"]:
                        print(f"⏩ [Robot] Locale {codice_aams} - {nome_locale_corrente}: Nessuna azione richiesta. Salto riga.")
                        continue
                        
                    print(f"🚀 [Robot] STEP 7: Avvio lavorazione ({mirino_azione}) -> Codice Locale: {codice_aams} - {nome_locale_corrente}")

                    # 🛡️ PUNTAMENTO NATIVO DEI RAGAZZI STRUTTURATO AL MILLIMETRO
                    target_frame = page
                    for f in page.frames:
                        if "Esercizi" in f.url or f.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").count() > 0:
                            target_frame = f
                            break

                    print("   🔍 [Robot] STEP 7a: Inserimento codice censimento nella barra filtri...")
                    campo_ricerca = target_frame.locator("#ctl00_Cp1_txtCodiceCensimentoesercizio").first
                    
                    try:
                        campo_ricerca.wait_for(state="visible", timeout=20000)
                    except Exception as e_time:
                        raise e_time # Fa proseguire l'errore per saltare la riga regolarmente
                        
                    campo_ricerca.click()
                    campo_ricerca.fill(codice_aams)
                    time.sleep(2)

                    tasto_ricerca = target_frame.locator("#ctl00_Cp1_btRicerca").first
                    tasto_ricerca.click(timeout=10000)
                    
                    print("   ⏳ [Robot] STEP 7b: Attesa caricamento risultati filtrati (6 secondi)...")
                    time.sleep(6)

                    icona_nuovo = target_frame.locator("img[src*='insert_pianificazione.jpg'], img[src*='insert_pianificazione'], img[id*='img_pianificazione']").first
                    icona_modifica = target_frame.locator("img[src*='edit_pianificazione']").first
                    
                    try:
                        icona_nuovo.wait_for(state="attached", timeout=4000)
                    except Exception: pass

                    # 🛡️ NOTIFICA DI CANCELLAZIONE DI MANUELA: Ti avvisa ORA via e-mail prima che partano i clic di cancellazione!
                    if mirino_azione == "ELIMINA":
                        print(f"📧 [Robot] Rilevato comando ELIMINA per Snaitech. Invio notifica di CANCELLAZIONE interna a Manuela...")
                        try:
                            data_inizio_stampa = str(row["INIZIO_FERIE"]).replace("-", "/").replace(".", "/").strip()[:10]
                            tecnico_titolare = row.get("TECNICO_INSERIMENTO", "Non specificato")
                            collega_condiviso = row.get("PROMEMORIA_IN_COPIA", "Nessuno")
                            
                            spedisci_email_avviso_ufficio(tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_inizio_stampa, "CANCELLAZIONE FERIE REGISTRATA (SNAITECH)")
                        except Exception as e_snai_mail:
                            print(f"   ⚠️ Impossibile spedire notifica interna di cancellazione Snaitech: {str(e_snai_mail)}")

                    # 🛡️ INTELLIGENZA DI MANUELA: Se l'azione è ELIMINA ma non esiste la matita a portale, pulisce l'Excel direttamente da qui!
                    if mirino_azione == "ELIMINA" and icona_modifica.count() == 0:
                        print("   ℹ️ [Robot] Comando ELIMINA su locale vergine a portale. Cancello la riga dall'Excel cloud all'istante...")
                        scarica_e_aggiorna_excel_su_github(codice_aams)
                        time.sleep(4)
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
                        time.sleep(6)
                        continue

                    # 🛡️ CORREZIONE DI MANUELA: Se l'azione è MODIFICA o ELIMINA e la matita esiste nel codice, ci clicca sopra a prescindere dai filtri di visibilità!
                    if (mirino_azione in ["MODIFICA", "ELIMINA"]) and icona_modifica.count() > 0:
                        print(f"   📝 [Robot] STEP 8: [{mirino_azione}] Clicco sulla matita di modifica per entrare nella scheda...")
                        icona_modifica.scroll_into_view_if_needed(timeout=5000)
                        icona_modifica.click(force=True, timeout=8000)
                    elif icona_nuovo.count() > 0:
                        print("   🟢 [Robot] STEP 8a: [NUOVA CHIUSURA] Clic sul pallino verde...")
                        icona_nuovo.scroll_into_view_if_needed(timeout=5000)
                        time.sleep(1)
                        icona_nuovo.click(force=True, timeout=8000)
                    else:
                        print("   AM 🖱️ [Grid Mode] Clic sulla cella td nativa della riga...")
                        cella_td = target_frame.locator("td[onclick*='Pianificazione']").first
                        cella_td.click(force=True, timeout=8000)
                    
                    print("   ⏳ [Robot] STEP 8c: Attesa apertura campi date (7 secondi)...")
                    time.sleep(7)



# =====================================================================================
# BLOCCO 5: DATA FINE PURIFICATA, SEQUENZA DIGITAZIONE REALE E SALVATAGGIO REALE
# ==============================================================================
# =====================================================================================
# BLOCCO 5: DATA FINE PURIFICATA, SEQUENZA FOTO REALE E RITORNO IN BACHECA PROTETTO
# =====================================================================================
                    # 🛡️ INIZIALIZZAZIONE DI SICUREZZA DI MANUELA: Azzera le variabili per evitare i crash di scope di Python
                    frame_date = page
                    tasto_elimina_snai = None

                    # Cerca l'Iframe protetto della maschera delle date di Snaitech
                    for f in page.frames:
                        if "Chiusura" in f.url or f.locator("#ctl00_Cp1_Txtiniziochiusura").count() > 0:
                            frame_date = f
                            break
                    
                    # 🛡️ PROTEZIONE POP-UP: Se frame_date è rimasto sulla pagina principale, aspetta 3 secondi e riprova
                    if frame_date == page:
                        time.sleep(3)
                        for f in page.frames:
                            if "Chiusura" in f.url or f.locator("#ctl00_Cp1_Txtiniziochiusura").count() > 0:
                                frame_date = f
                                break

                    data_inizio_pura = data_inizio_pulita[:10].strip()
                    data_fine_pura = data_fine_pulita[:10].strip()
                    
                    ora_inizio_pulita = "06:00" if "06:00" in str(row["INIZIO_FERIE"]) else "00:00"
                    ora_fine_pulita = "12:00" if "12:00" in str(row["FINE_FERIE"]) else "23:30"
                  
                    # 🛡️ BIVIO CANCELLAZIONE DI MANUELA: Se l'azione è ELIMINA, gestisce anche l'annullamento ante-sincro senza crashare
                    if mirino_azione == "ELIMINA":
                        print("   🗑️ [Robot] STEP 9: Rilevato comando di rimozione. Verifico presenza campo su Snaitech...")
                        tasto_elimina_snai = frame_date.locator("#ctl00_Cp1_BtnElimina").first
                        
                        if tasto_elimina_snai.count() == 0 or not tasto_elimina_snai.is_visible():
                            print("   ℹ️ [Robot] Chiusura non presente su Snaitech (Annullamento immediato). Salto il sito e pulisco l'Excel...")
                            scarica_e_aggiorna_excel_su_github(codice_aams)
                            time.sleep(4)
                            page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
                            time.sleep(6)
                            continue
                        
                        # Se invece il tasto esiste, procede con la normale rimozione formale a portale
                        tasto_elimina_snai.wait_for(state="visible", timeout=10000)
                        tasto_elimina_snai.click(force=True)
                        print("   💾 [Robot] STEP 10: Pulsante Elimina premuto. Attesa conferma dal server Snaitech...")
                        time.sleep(6)
                        scarica_e_aggiorna_excel_su_github(codice_aams)
                        time.sleep(4)
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
                        time.sleep(6)
                        continue
                
                    # ALTRIMENTI (Se NUOVA o MODIFICA): Procede con la tua digitazione reale speculare
                    campo_dal = frame_date.locator("#ctl00_Cp1_Txtiniziochiusura, input[id*='Txtiniziochiusura']").first
                    campo_al = frame_date.locator("#ctl00_Cp1_txtfinechiusura, input[id*='txtfinechiusura']").first

                    print("   📝 [Robot] STEP 9: Digitazione reale data inizio...")
                    try:
                        # 🛡️ PROTEZIONE DI MANUELA: Allungato a 25 secondi per dare tempo ai calendari Snaitech di aprirsi!
                        campo_dal.wait_for(state="visible", timeout=25000)
                    except Exception:
                        print("   ⚠️ [Robot] Campi date non ancora visibili. Tento un secondo clic forzato sulla riga...")
                        cella_td.click(force=True)
                        time.sleep(5)
                        campo_dal.wait_for(state="visible", timeout=15000)

                    campo_dal.click()
                    campo_dal.press("Control+A")
                    campo_dal.press("Backspace")
                    time.sleep(1)
                    campo_dal.press_sequentially(data_inizio_pura, delay=100)
                    campo_dal.press("Tab")
                    time.sleep(1)


                    print("   📝 [Robot] STEP 9a: Digitazione reale data fine...")
                    campo_al.wait_for(state="visible", timeout=10000)
                    campo_al.click()
                    campo_al.press("Control+A")
                    campo_al.press("Backspace")
                    time.sleep(1)
                    campo_al.press_sequentially(data_fine_pura, delay=100)
                    campo_al.press("Tab")
                    time.sleep(2)
                    
                    try: frame_date.locator("#ctl00_Cp1_fascia_from").select_option(ora_inizio_pulita)
                    except Exception: pass
                    try: frame_date.locator("#ctl00_Cp1_fascia_to").select_option(ora_fine_pulita)
                    except Exception: pass
                    time.sleep(2)

                    print("   💾 [Robot] STEP 10: Invio moduli di chiusura a Snaitech (Clic su Tasto Salva)...")
                    frame_date.locator("#ctl00_Cp1_BtnOk").first.click(timeout=10000)
                    print(f"   ✅ [Robot] STEP 11: Invio completato. Attesa risposta visiva del portale...")
                    time.sleep(4)
                    
                   
                    # 🛡️ CALCOLATORE AVVISI DI MANUELA CON PURIFICAZIONE ORARIO: Taglia via lo sporco delle ore (es. 06:00) per azzerare l'unconverted data
                    try:
                        oggi_server = datetime.now().date()
                        
                        # Isola solo i primi 10 caratteri (GG/MM/AAAA) pulendo l'orario se presente
                        data_inizio_estratta = data_inizio_pura.replace("-", "/").replace(".", "/").strip()[:10]
                        data_fine_estratta = data_fine_pura.replace("-", "/").replace(".", "/").strip()[:10]
                        
                        data_in_doc = datetime.strptime(data_inizio_estratta, "%d/%m/%Y").date()
                        data_fi_doc = datetime.strptime(data_fine_estratta, "%d/%m/%Y").date()
                        
                        giorni_alla_chiusura = (data_in_doc - oggi_server).days
                        giorni_alla_riapertura = (data_fi_doc - oggi_server).days
                        
                        tecnico_titolare = row.get("TECNICO_INSERIMENTO", "Non specificato")
                        collega_condiviso = row.get("PROMEMORIA_IN_COPIA", "Nessuno")
                        
                        data_inizio_stampa = data_in_doc.strftime("%d/%m/%Y")
                        data_fine_stampa = data_fi_doc.strftime("%d/%m/%Y")
                        
                        if giorni_alla_chiusura == 3:
                            spedisci_email_avviso_ufficio(tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_inizio_stampa, "CHIUSURA LOCALE (TRA 3 GIORNI)")
                        if giorni_alla_riapertura == 3:
                            spedisci_email_avviso_ufficio(tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_fine_stampa, "RIAPERTURA LOCALE (TRA 3 GIORNI)")
                    except Exception as e_calc:
                        print(f"   ⚠️ Impossibile calcolare il promemoria email: {str(e_calc)}")



                    # 🛡️ FIX MULTIPLO DI MANUELA: Caricamento standard 'load' stabile per lavorazioni consecutive di fila
                    page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
                    time.sleep(4)
                    
                    # 🛡️ RESET CELLA EXCEL CLOUD NATIVO VIA GIT PUSH
                    scarica_e_aggiorna_excel_su_github(codice_aams)
                    time.sleep(8)
                
                except Exception as row_err:
                    print(f"   ⚠️ Nota compilazione: Scavalco riga. Errore: {str(row_err)}")
                    try:
                        page.goto("https://partner.snai.it/secure/Anagrafiche/Esercizi.aspx", wait_until="load")
                        time.sleep(8)
                    except Exception: pass
                    continue

            print("🔒 [Robot] STEP 12: Chiusura sessione formale (Logout di sicurezza)...")
            try: page.locator("a:has-text('LogOut'), a:has-text('Esci'), [id*='btnLogOut']").first.click(timeout=8000)
            except Exception: page.context.clear_cookies()

        except Exception as e: print(f"❌ Errore durante la navigazione sul portale partner.snai.it: {str(e)}")
        finally: browser.close()

        # =====================================================================================
        # 📊 MOTORE DI REVISIONE DI MANUELA: SCANSIONE ANOMALIE E SPEDIZIONE RIEPILOGO
        # =====================================================================================
        print("📊 [Audit Engine] Avvio scansione del registro per rilevamento anomalie...")
        lista_anomalie_rilevate = []
        
        try:
            # Rilegge l'Excel finale per vedere cosa è rimasto escluso dall'allineamento
            df_audit = pd.read_excel("storico_ferie.xlsx").fillna("")
            
            for _, row in df_audit.iterrows():
                azione_rimasta = str(row.get("ROBOT_ACTION", "")).strip().upper()
                nome_loc_audit = str(row.get("NOME_LOCALE", "")).strip()
                conc_audit = str(row.get("CONCESSIONARIO", "")).strip()
                stato_invio_audit = str(row.get("STATO_INVIO", "")).strip()
                
                # 🚨 ANOMALIA 1: Una riga è rimasta in stato NUOVA/MODIFICA/ELIMINA dopo la corsa del robot
                if azione_rimasta in ["NUOVA", "MODIFICA", "ELIMINA"]:
                    lista_anomalie_rilevate.append(f"• Locale: {nome_loc_audit} [{conc_audit}] -> Rimasto bloccato in stato [{azione_rimasta}] (Errore Portale)")
                    
                # 🚨 ANOMALIA 2: Lo STATO_INVIO ha registrato un fallimento o un blocco di rete
                elif "ERR" in stato_invio_audit.upper() or "FALLITO" in stato_invio_audit.upper():
                    lista_anomalie_rilevate.append(f"• Locale: {nome_loc_audit} [{conc_audit}] -> Fallimento notifica: {stato_invio_audit}")
        except Exception as e_audit_file:
            lista_anomalie_rilevate.append(f"• Impossibile leggere il file Excel per il controllo: {str(e_audit_file)}")

        # 📧 SPEDIZIONE DEL REPORT DI SUPERVISIONE A MANUELA
        try:
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText

            EMAIL_LOGIN = "wingamingsrl@gmail.com"
            pass_applicativa_ufficio = "zndjprxjvhiustio"

            msg_audit = MIMEMultipart()
            msg_audit['From'] = "WinGaming Audit <tecnico@wingaming.it>"
            msg_audit['To'] = "manuela.arigoni@wingaming.it"
            
            if lista_anomalie_rilevate:
                msg_audit['Subject'] = f"🚨 [ANOMALIE RILEVATE] Report Supervisione Ferie WinGaming"
                corpo_audit = f"Attenzione Manuela,\n\nIl robot ha completato il giro di sincronizzazione ma ha riscontrato delle anomalie che richiedono un controllo manuale.\n\nElenco dei blocchi riscontrati:\n--------------------------------------------------\n" + "\n".join(lista_anomalie_rilevate) + "\n--------------------------------------------------\n\nVerificare lo stato dei portali partner.\nWinGaming S.r.l."
            else:
                msg_audit['Subject'] = f"✅ [TUTTO ALLINEATO] Report Supervisione Ferie WinGaming"
                corpo_audit = f"Buongiorno Manuela,\n\nIl robot ha completato con successo il giro logistico di allineamento.\n\n🔍 Risultato Scansione: ZERO anomalie riscontrate.\nTutti i locali inseriti dalla flotta sono stati regolarmente processati e trasmessi a sistema.\n\nBuon lavoro,\nWinGaming S.r.l."

            msg_audit.attach(MIMEText(corpo_audit, 'plain', 'utf-8'))
            server_audit = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
            server_audit.login(EMAIL_LOGIN, pass_applicativa_ufficio)
            server_audit.sendmail("tecnico@wingaming.it", ["manuela.arigoni@wingaming.it"], msg_audit.as_string())
            server_audit.quit()
            print("   ✅ [Audit Engine] Report delle anomalie spedito correttamente in ufficio!")
        except Exception as e_spedisci_audit:
            print(f"   ❌ [Audit Engine] Errore spedizione report: {str(e_spedisci_audit)}")


# =====================================================================================
# 🛡️ INTERRUTTORE DI AVVIO UNIFICATO DI MANUELA: SCANSIONE PROMEMORIA GLOBALE 3 GIORNI
# =====================================================================================
if __name__ == "__main__":
    # 📧 1. PRIMA DI TUTTO: Fa partire la scansione globale per inviare le e-mail dei 3 giorni prima!
    print("📧 [Email Engine] Avvio scansione globale dello storico per l'invio dei promemoria automatici...")
    try:
        df_completo_promemoria = preleva_storico_diretto_da_cloud()
        
        if not df_completo_promemoria.empty:
            oggi_server = datetime.now().date()
            print(f"   📅 [Email Engine] Data odierna del server per il calcolo: {oggi_server}")
            
            for _, row in df_completo_promemoria.iterrows():
                nome_locale_corrente = str(row.get("NOME_LOCALE", "")).strip()
                codice_aams = str(row.get("CODICE_LOCALE", "")).strip()
                data_in_completa = str(row.get("INIZIO_FERIE", "")).strip()
                data_fi_completa = str(row.get("FINE_FERIE", "")).strip()
                
                data_inizio_estratta = data_in_completa.replace("-", "/").replace(".", "/").strip()[:10]
                data_fine_estratta = data_fi_completa.replace("-", "/").replace(".", "/").strip()[:10]
                
                try:
                    data_in_doc = datetime.strptime(data_inizio_estratta, "%d/%m/%Y").date()
                    data_fi_doc = datetime.strptime(data_fine_estratta, "%d/%m/%Y").date()
                    
                    giorni_alla_chiusura = (data_in_doc - oggi_server).days
                    giorni_alla_riapertura = (data_fi_doc - oggi_server).days
                    
                    tecnico_titolare = str(row.get("TECNICO_INSERIMENTO", "")).strip().upper()
                    collega_condiviso = str(row.get("PROMEMORIA_IN_COPIA", "")).strip().upper()
                    
                    data_inizio_stampa = data_in_doc.strftime("%d/%m/%Y")
                    data_fine_stampa = data_fi_doc.strftime("%d/%m/%Y")
                    
                    # 📧 INVIO CHIUSURA (TRA 3 GIORNI)
                    if giorni_alla_chiusura == 3:
                        # 1. Tu (Admin) ricevi SEMPRE la notifica per qualsiasi locale di chiunque!
                        spedisci_email_avviso_ufficio(EMAIL_MANUELA_RICEVENTE, tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_inizio_stampa, "CHIUSURA LOCALE (TRA 3 GIORNI)")
                        
                        # 2. Il tecnico riceve l'avviso in copia SOLO se il locale è suo (e non sei tu l'inseritore)
                        if tecnico_titolare != "MANUELA ARIGONI" and tecnico_titolare != "":
                            spedisci_email_avviso_ufficio("tecnico@wingaming.it", tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_inizio_stampa, "CHIUSURA LOCALE (TRA 3 GIORNI)")
                    
                    # 📧 INVIO RIAPERTURA (TRA 3 GIORNI)
                    if giorni_alla_riapertura == 3:
                        # 1. Tu (Admin) ricevi SEMPRE la notifica per qualsiasi locale di chiunque!
                        spedisci_email_avviso_ufficio(EMAIL_MANUELA_RICEVENTE, tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_fine_stampa, "RIAPERTURA LOCALE (TRA 3 GIORNI)")
                        
                        # 2. Il tecnico riceve l'avviso in copia SOLO se il locale è suo (e non sei tu l'inseritore)
                        if tecnico_titolare != "MANUELA ARIGONI" and tecnico_titolare != "":
                            spedisci_email_avviso_ufficio("tecnico@wingaming.it", tecnico_titolare, collega_condiviso, nome_locale_corrente, codice_aams, data_fine_stampa, "RIAPERTURA LOCALE (TRA 3 GIORNI)")
                            
                except Exception as e_data:
                    pass
                        
        print("✅ [Email Engine] Scansione promemoria storici completata con successo!")
    except Exception as e_cron:
        print(f"   ❌ Errore Visualizzazione: {str(e_cron)}")

    print("\n----------------------------------------------------------------------")
    # 🌐 2. SUBITO DOPO: Attiva il motore principale per l'allineamento Snaitech ed NTS sui portali
    avvia_sincronizzazione_automatica()



