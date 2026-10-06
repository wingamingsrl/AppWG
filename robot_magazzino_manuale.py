import os
import sys
import time
import base64
import requests
import openpyxl
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright

# 📂 FILIERA TARGET INVENTARIO WIN GAMING
FILE_GIACENZA_SCHEDE = "giacenza_schede.xlsx"
FILE_SLOT_PORTALI = "elenco_slot_portali.xlsx"
FILE_ELENCO_MAGAZZINI = "elenco_magazzini.xlsx"
FILE_GREZZO_SCHEDE = "schede_grezze_gameslodi.xlsx"

URL_GESTIONALE = "https://gestionale.gameslodi.it"
G_EMAIL = os.environ.get("GAMESLODI_EMAIL", "manuela.arigoni@wingaming.it")
G_PASSWORD = os.environ.get("GAMESLODI_PASSWORD", "")
TOKEN_GITHUB = os.environ.get("TOKEN_ACCESSO_GITHUB", "")

def spingi_file_su_github(nome_file_target):
    try:
        import base64
        if not os.path.exists(nome_file_target): return False
        s_api = "api" + "." + "github" + "." + "com"
        url_api = f"https://{s_api}/repos/wingamingsrl/AppWG-Test/contents/{nome_file_target}"
        
        with open(nome_file_target, "rb") as f: contenuto_binario = f.read()
        dati_base64 = base64.b64encode(contenuto_binario).decode('utf-8')
        headers = {"Authorization": f"token {TOKEN_GITHUB}", "Accept": "application/vnd.github+json", "User-Agent": "WinGaming-Automation-Engine"}
        
        res_get = requests.get(url_api, headers=headers, timeout=5)
        sha_file = res_get.json().get("sha", "") if res_get.status_code == 200 else ""
        
        payload = {"message": f"🤖 [Robot Magazzino] Sincronizzazione database {nome_file_target}", "content": dati_base64, "branch": "main"}
        if sha_file: payload["sha"] = sha_file
        res_put = requests.put(url_api, json=payload, headers=headers, timeout=10)
        return res_put.status_code == 200 or res_put.status_code == 201
    except Exception: return False

def esegui_estrazione_magazzino_puro():
    print(f"🚀 Avvio del robot Magazzino Isolato - Ore {datetime.now().strftime('%H:%M:%S')}")
    if not G_PASSWORD or not TOKEN_GITHUB: return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={"width": 1280, "height": 800})
        page = context.new_page()
        
        try:
            print(f"🌐 [FOTO 1] Collegamento a: {URL_GESTIONALE}...")
            page.goto(URL_GESTIONALE, wait_until="load", timeout=45000)
            page.wait_for_timeout(4000)
            
            print("🔑 [FOTO 2] Compilazione dei campi di testo...")
            page.fill("input[name='username']", G_EMAIL)
            page.fill("input[name='userpassword']", G_PASSWORD)
            
            print("🚀 [FOTO 3] Clic sul pulsante Log In...")
            page.click("button[type='submit'], input[type='submit']")
            page.wait_for_timeout(6000)
            
            print("📂 [FOTO 4] Spostamento in Area Logistica -> Elenco Schede...")
            page.click("text=Area Logistica")
            page.wait_for_timeout(1000)
            page.click("text=Awp - Comma 6")
            page.wait_for_timeout(1000)
            page.click("text=Elenco Schede")
            page.wait_for_timeout(4000)
            
            page.wait_for_selector("select[name='cabinet']", timeout=15000)
            page.select_option("select[name='cabinet']", value="")
            page.wait_for_timeout(1000)
            page.click("#btn-cerca")
            page.wait_for_timeout(4000)
            
            with page.expect_download(timeout=45000) as info_dl2:
                page.click("button:has-text('Genera Excel'), button[name='excel']")
            info_dl2.value.save_as(FILE_GREZZO_SCHEDE)
            
            df_grezzo = pd.read_excel(FILE_GREZZO_SCHEDE).fillna("")
            df_grezzo.to_excel(FILE_GIACENZA_SCHEDE, index=False)
            
            print("📂 [FOTO 5] Navigazione verso Portali -> Elenco Slot Portali...")
            page.click("text=Portali")
            page.wait_for_timeout(1000)
            page.click("text=Elenco slot portali")
            page.wait_for_timeout(3000)
            
            page.select_option("select#stato", value="")
            page.wait_for_timeout(500)
            page.click("#btn-cerca")
            page.wait_for_timeout(5000)
            
            with page.expect_download(timeout=45000) as info_dl3:
                page.click(".buttons-excel")
            info_dl3.value.save_as(FILE_SLOT_PORTALI)

            # ➡️ 🚀 GIRO 5 DI MANUELA: AREA LOGISTICA -> MAGAZZINI -> ELENCO MAGAZZINI
            print("📂 [GIRO 5] Navigazione verso Area Logistica -> Magazzini -> Elenco Magazzini...")
            page.click("text=Area Logistica")
            page.wait_for_timeout(1000)
            page.click("text=Magazzini")
            page.wait_for_timeout(1000)
            page.click("text=Elenco magazzini")
            page.wait_for_timeout(4000)
            page.screenshot(path="FOTO_07_elenco_magazzini.png")
            
            print("🟢 Scarico il terzo Excel dell'Elenco Magazzini tramite classe .buttons-excel...")
            with page.expect_download(timeout=45000) as info_dl4:
                # Clicca sul bottone speciale DataTables indicato da Manuela
                page.click(".buttons-excel")
            info_dl4.value.save_as(FILE_ELENCO_MAGAZZINI)
            print("   -> File Elenco Magazzini salvato in locale.")

            # Spinta dei 3 fogli ufficiali su GitHub
            print("🧭 [Magazzino] Avvio spinta dei file sulla repository cloud...")
            spingi_file_su_github(FILE_GIACENZA_SCHEDE)
            spingi_file_su_github(FILE_SLOT_PORTALI)
            spingi_file_su_github(FILE_ELENCO_MAGAZZINI)
            print("✅ [SUCCESSO] Sincronizzazione magazzino a comando completata!")
                    
        except Exception as e:
            print(f"💥 ERRORE CRITICO: {str(e)}")
            try:
                page.screenshot(path="FOTO_ERRORE_BLOCCO_RETE.png")
                spingi_file_su_github("FOTO_ERRORE_BLOCCO_RETE.png")
            except Exception: pass
        finally:
            browser.close()

if __name__ == "__main__":
    esegui_estrazione_magazzino_puro()
