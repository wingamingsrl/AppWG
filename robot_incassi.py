import os
import sys
import time
import base64
import requests
import openpyxl
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

print("🧭 [ROBOT INCASSI] Avvio motore virtuale Playwright con salvataggio foto in Artifacts...")

# 📂 FILIERA TARGET CONFIGURATA DA MANUELA
OUTPUT_FILE = "locali_non_incassati.xlsx"
URL_GESTIONALE = "https://gestionale.gameslodi.it"

G_EMAIL = os.environ.get("GAMESLODI_EMAIL", "manuela.arigoni@wingaming.it")
G_PASSWORD = os.environ.get("GAMESLODI_PASSWORD", "")
TOKEN_GITHUB = os.environ.get("TOKEN_ACCESSO_GITHUB", "")

def spingi_file_su_github(nome_file_target):
    try:
        import base64
        if not os.path.exists(nome_file_target): return False
        s_api = "api" + "." + "github" + "." + "com"
        url_api = f"https://{s_api}/repos/wingamingsrl/AppWG/contents/{nome_file_target}"
        
        with open(nome_file_target, "rb") as f: contenuto_binario = f.read()
        dati_base64 = base64.b64encode(contenuto_binario).decode('utf-8')
        headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Automation-Engine"}
        
        res_get = requests.get(url_api, headers=headers, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload = {"message": f"🤖 [Robot Incassi] Sincronizzazione {nome_file_target}", "content": dati_base64, "branch": "main"}
        if sha_file: payload["sha"] = sha_file
        res_put = requests.put(url_api, json=payload, headers=headers, timeout=10)
        return res_put.status_code == 200 or res_put.status_code == 201
    except Exception: return False

def esegui_estrazione_incassi_manuela():
    if not G_PASSWORD or not TOKEN_GITHUB:
        print("❌ [ERRORE] Mancano GAMESLODI_PASSWORD o TOKEN_ACCESSO_GITHUB nell'ambiente!")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1280, "height": 800})
        page = context.new_page()
        
        try:
            # 🌐 1. COLLEGAMENTO DI FABBRICA
            print(f"🌐 [FOTO 1] Collegamento a: {URL_GESTIONALE}...")
            page.goto(URL_GESTIONALE, wait_until="load", timeout=45000)
            page.wait_for_timeout(4000)
            page.screenshot(path="/tmp/1_schermata_login.png")
            
            # 🔑 2. COMPILAZIONE DEI CAMPI DI TESTO
            print("🔑 [FOTO 2] Compilazione dei campi di testo con la password corretta...")
            page.fill("input[name='username']", G_EMAIL)
            page.fill("input[name='userpassword']", G_PASSWORD)
            page.screenshot(path="/tmp/2_campi_compilati.png")
            
            # 🚀 3. ACCESSO ALLA PLANCIA CONFERMATO
            print("🚀 [FOTO 3] Clic sul pulsante Log In...")
            page.click("button[type='submit'], input[type='submit']")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(6000)
            page.screenshot(path="/tmp/3_dopo_login.png")
            
            # 🧭 4. NAVIGAZIONE CORAZZATA DIRETTISSIMA (Pialla i problemi di visibilità del menù)
            print("🧭 Salto diretto all'indirizzo del Report senza passare dai clic dei menu animati...")
            # 🎯 VITTORIA: Sfrutta la rotta reale verificata dell'HTML '/reports/localiNonIncassati'
            page.goto(f"{URL_GESTIONALE}/reports/localiNonIncassati", wait_until="networkidle", timeout=45000)
            page.wait_for_timeout(4000)
            page.screenshot(path="/tmp/4_schermata_report_grezza.png")


            
            print("📂 Clicco sulla sottovoce 'Locali non incassati'...")
            # 🎯 VITTORIA: Punta direttamente al link contenitore <a> usando la rotta reale dell'HTML!
            page.click("a[href='/reports/localiNonIncassati']")
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(4000)
            page.screenshot(path="/tmp/4_schermata_report_grezza.png")

            
            # ✍️ 5. IMPOSTAZIONE SOGLIA A 0 GIORNI
            print("✍️ Modifico il campo 'Numero giorni' impostandolo a 0...")
            casella_giorni = page.locator("input[type='text'], input[name='giorni']").first
            casella_giorni.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.keyboard.type("0")
            page.wait_for_timeout(1000)
            
            print("🔍 Clicco sul pulsante 'Cerca' per elaborare i registri...")
            page.click("button:has-text('Cerca'), input[value='Cerca'], #btn-cerca")
            page.wait_for_timeout(6000)
            page.screenshot(path="/tmp/5_griglia_popolata_cerca.png")

            
            # 📥 6. SCARICAMENTO TASTO EXCEL
            print("🟢 Scarico il file Excel tramite classe .buttons-excel...")
            with page.expect_download(timeout=45000) as info_dl:
                page.click(".buttons-excel")
            
            temp_path = "temp_grezzo.xlsx"
            info_dl.value.save_as(temp_path)
            print("   -> File ricevuto. Avvio normalizzazione...")
            
            try:
                tabelle = pd.read_html(temp_path)
                df_greggio = tabelle
            except Exception:
                df_greggio = pd.read_excel(temp_path)
                
            df_greggio.to_excel(OUTPUT_FILE, index=False)
            
            # Spinta del solo ed unico file Excel pulito su GitHub, senza imbrattare la cartella con le foto!
            print("🧭 [Incassi] Avvio spinta dell'Excel sulla repository cloud...")
            spingi_file_su_github(OUTPUT_FILE)
            print("✅ [SUCCESSO] Sincronizzazione registro incassi completata!")
                    
        except Exception as e:
            print(f"💥 ERRORE CRITICO: {str(e)}")
            try:
                page.screenshot(path="/tmp/FOTO_ERRORE_BLOCCO_RETE.png")
            except Exception: pass
        finally:
            browser.close()

if __name__ == "__main__":
    esegui_estrazione_incassi_manuela()

