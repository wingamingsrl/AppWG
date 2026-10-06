import os
import time
import requests
import pandas as pd
from playwright.sync_api import sync_playwright

URL_GESTIONALE = "https://gestionale.gameslodi.it"
FILE_GREZZO_LOCALI = "locali_grezzi_gameslodi.xlsx"
FILE_GREZZO_SCHEDE = "schede_grezze_gameslodi.xlsx"
FILE_ANAGRAFICA_FINALE = "elenco_locali.xlsx"

G_EMAIL = os.environ.get("GAMESLODI_EMAIL", "manuela.arigoni@wingaming.it")
G_PASSWORD = os.environ.get("GAMESLODI_PASSWORD", "")
TOKEN_GITHUB = os.environ.get("TOKEN_ACCESSO_GITHUB", "")

def elabora_incrocio_e_purifica():
    print("📊 [Data Engine] Avvio incrocio database ed eliminazione duplicati...")
    try:
        # 1. Carica il file delle schede AWP scaricato (il più importante per i concessionari reali)
        if not os.path.exists(FILE_GREZZO_SCHEDE):
            print("❌ Errore: File schede AWP non trovato sul server.")
            return False
            
        df_schede = pd.read_excel(FILE_GREZZO_SCHEDE).fillna("")
        df_schede.columns = [str(col).strip() for col in df_schede.columns]
        
        # Isola ed estrae solo le colonne indicate da Manuela per l'incrocio delle combinazioni
        df_incrocio = pd.DataFrame()
        
        # Mappa il Codice AAMS (Codice Esercizio)
        if "Codice Esercizio" in df_schede.columns:
            df_incrocio["CODICE_LOCALE"] = df_schede["Codice Esercizio"].astype(str).str.strip()
        else:
            print("⚠️ Colonna 'Codice Esercizio' non trovata. Uso la prima colonna disponibile.")
            df_incrocio["CODICE_LOCALE"] = df_schede.iloc[:, 0].astype(str).str.strip()
            
        # Mappa il Nome Locale (Luogo)
        if "Luogo" in df_schede.columns:
            df_incrocio["NOME_LOCALE"] = df_schede["Luogo"].astype(str).str.strip().str.upper()
        else:
            df_incrocio["NOME_LOCALE"] = df_schede.iloc[:, 1].astype(str).str.strip().str.upper()
            
        # Mappa il Concessionario reale delle schede machine
        if "Concessionario" in df_schede.columns:
            df_incrocio["CONCESSIONARIO"] = df_schede["Concessionario"].astype(str).str.strip()
        else:
            df_incrocio["CONCESSIONARIO"] = "Snaitech Spa WG"

        # Pulisce i nomi dei concessionari per farli combaciare con i portali
        # Pulisce i nomi dei concessionari per farli combaciare con i portali
        def normalizza_concessionario(val):
            v = str(val).upper().strip()
            if "SNAITECH WNG" in v: return "Snaitech Spa WG"
            if "GLOBAL" in v: return "Global Starnet"
            if "NTS" in v: return "NTS Networks"
            if "EUROBET" in v: return "Eurobet"
            if "LOTTOMATICA" in v or "GBM" in v: return "Lottomatica"
            if "SISAL" in v: return "Sisal Entertainment S.p.a."  # 🛡️ INTEGRATO DA MANUELA
            return val

            
        df_incrocio["CONCESSIONARIO"] = df_incrocio["CONCESSIONARIO"].apply(normalizza_concessionario)

        # 🚨 TRATTAMENTO RIGHE DOPPIE DI MANUELA: Elimina i duplicati tenendo solo le combinazioni singole uniche!
        print(f"🔍 Righe grezze totali rilevate nelle schede: {len(df_incrocio)}")
        df_pulito = df_incrocio.drop_duplicates(subset=["CODICE_LOCALE", "CONCESSIONARIO"]).reset_index(drop=True)
        
        # Riorganizza la struttura finale rigida ed elimina eventuali codici vuoti o intestazioni spurie
        df_pulito = df_pulito[df_pulito["CODICE_LOCALE"] != ""].reset_index(drop=True)
        df_pulito = df_pulito.reindex(columns=["CODICE_LOCALE", "NOME_LOCALE", "CONCESSIONARIO"]).fillna("Snaitech Spa WG")
        
        # 🔤 ORDINAMENTO ALFABETICO DI MANUELA: Cataloga la lista dalla A alla Z per il Nome del Locale
        df_pulito = df_pulito.sort_values(by=["NOME_LOCALE"], ascending=True).reset_index(drop=True)
        
        # Salva sovrascrivendo l'anagrafica ufficiale
        df_pulito.to_excel(FILE_ANAGRAFICA_FINALE, index=False)
        print(f"✅ [Data Engine] Incrocio riuscito! Create {len(df_pulito)} combinazioni uniche pulite per la tendina.")
        return True
    except Exception as e_data:
        print(f"❌ [Data Engine] Fallimento elaborazione: {str(e_data)}")
        return False

def spingi_nuovo_excel_su_github():
    try:
        import base64
        s_api = "api" + "." + "github" + "." + "com"
        url_api = f"https://{s_api}/repos/wingamingsrl/AppWG/contents/{FILE_ANAGRAFICA_FINALE}"
        
        
        with open(FILE_ANAGRAFICA_FINALE, "rb") as f: contenuto_binario = f.read()
        dati_base64 = base64.b64encode(contenuto_binario).decode('utf-8')
        headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Automation-Engine"}
        res_get = requests.get(url_api, headers=headers, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        payload = {"message": "🔄 [Automazione] Aggiornamento anagrafica locali tramite incrocio AWP GamesLodi", "content": dati_base64, "branch": "main"}
        if sha_file: payload["sha"] = sha_file
        res_put = requests.put(url_api, json=payload, headers=headers, timeout=10)
        return res_put.status_code == 200 or res_put.status_code == 201
    except Exception: return False

def esegui_estrazione_gameslodi():
    print("🚀 [GamesLodi Robot] Avvio estrazione anagrafica combinata ed incrociata...")
    if not G_PASSWORD or not TOKEN_GITHUB:
        print("❌ ERRORE: Chiavi di sicurezza mancanti nei Secrets. Blocco.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        try:
            print(f"🌐 Collegamento a: {URL_GESTIONALE}...")
            page.goto(URL_GESTIONALE, wait_until="load", timeout=45000)
            page.wait_for_timeout(4000)
            
            print("🔑 Inserimento credenziali amministrative (Sansone)...")
            page.fill("input[name='username']", G_EMAIL)
            page.fill("input[name='userpassword']", G_PASSWORD)
            page.click("button[type='submit'], input[type='submit']")
            page.wait_for_timeout(5000)
            
            # ➡️ GIRO 1: SCARICAMENTO REGISTRO LOCALI
            print("📂 [GIRO 1] Navigazione visiva verso 'Elenco locali'...")
            page.click("text=Area Anagrafica")
            page.wait_for_timeout(1000)
            page.click("text=Locali")
            page.wait_for_timeout(1000)
            page.click("text=Elenco locali")
            page.wait_for_timeout(4000)
            
            print("🟢 Scarico il primo Excel dei Locali...")
            with page.expect_download(timeout=30000) as info_dl1:
                page.click("button:has-text('Genera excel'), button[name='excel']")
            download1 = info_dl1.value
            download1.save_as(FILE_GREZZO_LOCALI)
            print("   -> File Locali archiviato temporaneamente.")
            
            # ➡️ GIRO 2: SCARICAMENTO REGISTRO SCHEDE AWP CON FILTRO "TUTTI"
            print("📂 [GIRO 2] Navigazione visiva verso 'Elenco Schede AWP'...")
            page.click("text=Area Logistica")
            page.wait_for_timeout(1000)
            page.click("text=Awp - Comma 6")
            page.wait_for_timeout(1000)
            page.click("text=Elenco Schede")
            page.wait_for_timeout(4000)
            
            print("🎛️ Applico il filtro 'Tutti' sul campo Con Cabinet...")
            # 🛡️ MIRINO DI MANUELA: Aggancia specificamente la tendina cabinet impostando il valore vuoto di "Tutti"
            page.wait_for_selector("select[name='cabinet']", timeout=15000)
            page.select_option("select[name='cabinet']", value="")

            page.wait_for_timeout(1000)
            
            print("🔍 Clicco sul bottone Filtra...")
            page.click("#btn-cerca")
            page.wait_for_timeout(4000)
            
            print("🟢 Scarico il secondo Excel delle Schede AWP usando il mirino di Manuela...")
            with page.expect_download(timeout=45000) as info_dl2:
                # 🛡️ FIX FINALE DI MANUELA: Punta al tasto verde basandosi sul testo reale stampato a schermo
                page.click("button:has-text('Genera Excel'), button[name='excel']")
            download2 = info_dl2.value
            download2.save_as(FILE_GREZZO_SCHEDE)

            print("   -> File Schede AWP archiviato temporaneamente.")
            
            # ➡️ GIRO 3: INCROCIO PROGREDITO E SALVATAGGIO DEFINITIVO
            if elabora_incrocio_e_purifica():
                if spingi_nuovo_excel_su_github():
                    print("✅ [MIRACOLO] Allineamento anagrafico incrociato completato! File spinto online!")
                    
        except Exception as e:
            errore_rilevato = str(e)
            print(f"💥 ERRORE CRITICO DURANTE LA NAVIGAZIONE: {errore_rilevato}")
            
            # 📸 Scatta la foto visiva del blocco sul server
            try: page.screenshot(path="screenshot_errore_blocco.png")
            except Exception: pass
            
            # 📧 SPEDIZIONE ALLERTA EMAIL IMMEDIATA A MANUELA
            try:
                import smtplib
                from email.mime.multipart import MIMEMultipart
                from email.mime.text import MIMEText

                EMAIL_LOGIN = "wingamingsrl@gmail.com"
                # Recupera la password applicativa aziendale inserita nei Secrets di GitHub
                PASS_APP = os.environ.get("GAMESLODI_PASSWORD", "zndjprxjvhiustio") 

                msg_allerta = MIMEMultipart()
                msg_allerta['From'] = "WinGaming Backup System <tecnico@wingaming.it>"
                msg_allerta['To'] = "manuela.arigoni@wingaming.it"
                msg_allerta['Subject'] = f"🚨 [FALLIMENTO NOTTURNO] Errore Estrattore GamesLodi"
                
                corpo_allerta = (
                    f"Attenzione Manuela,\n\n"
                    f"Il robot notturno delle 05:45 ha riscontrato un fallimento bloccante durante l'estrazione dei locali da GamesLodi.\n\n"
                    f"Il database elenco_locali.xlsx NON è stato aggiornato per evitare di danneggiare la plancia.\n\n"
                    f"Dettaglio dell'errore tecnico riscontrato:\n"
                    f"--------------------------------------------------\n"
                    f"⚠️ {errore_rilevato}\n"
                    f"--------------------------------------------------\n\n"
                    f"Verificare se il gestionale Sansone è in manutenzione o se è cambiata la grafica dei pulsanti.\n"
                    f"Su GitHub Actions trovi lo screenshot salvato del blocco.\n\n"
                    f"WinGaming S.r.l."
                )

                msg_allerta.attach(MIMEText(corpo_allerta, 'plain', 'utf-8'))
                server_allerta = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
                server_allerta.login(EMAIL_LOGIN, PASS_APP)
                server_allerta.sendmail("tecnico@wingaming.it", ["manuela.arigoni@wingaming.it"], msg_allerta.as_string())
                server_allerta.quit()
                print("   ✅ Allerta e-mail inviata con successo in ufficio a Manuela!")
            except Exception as e_mail_allerta:
                print(f"   ❌ Impossibile spedire l'e-mail di allerta: {str(e_mail_allerta)}")

            browser.close()

if __name__ == "__main__":
    esegui_estrazione_gameslodi()
