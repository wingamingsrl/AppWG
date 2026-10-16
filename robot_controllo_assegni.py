import io
import base64
import requests
import pandas as pd
from datetime import datetime
import os

# 🎯 INDIRIZZAMENTI STRUTTURALI COMPLETI E CERTIFICATI
parte_api_tab = "api" + "." + "github" + "." + "com"
parte_repo_tab = "repos/wingamingsrl/AppWG"
URL_TABELLONE = f"https://{parte_api_tab}/{parte_repo_tab}/contents/registro_assegni.xlsx"
URL_STORICO = f"https://{parte_api_tab}/{parte_repo_tab}/contents/Storico_assegni.xlsx"

# Il token viene letto in automatico dall'ambiente sicuro di GitHub Actions
token_github = os.environ.get("TOKEN_ACCESSO_GITHUB", "").strip()
headers = {
    "Authorization": f"token {token_github}", 
    "Accept": "application/vnd.github+json", 
    "User-Agent": "WinGaming-Cloud-App"
}

def esegui_controllo_notturno():
    print("⏳ Avvio controllo notturno Assegni...")
    res = requests.get(URL_TABELLONE, headers=headers, timeout=15)
    if res.status_code != 200:
        print(f"❌ Impossibile leggere il registro (HTTP {res.status_code})")
        return
        
    dati_j = res.json()
    df = pd.read_excel(io.BytesIO(base64.b64decode(dati_j["content"])), dtype={"NUMERO_ASSEGNO": str}).fillna("")
    
    oggi_dt = datetime.now().date()
    da_aggiornare = False
    
    # ⏱️ 1. SMARCO AUTOMATICO IN VERSATO OK (SE PASSATO 1 GIORNO)
    for idx, riga in df.iterrows():
        if riga["STATO_VERSAMENTO"] == "In Portafoglio" and riga["DATA_ASSEGNO"]:
            try:
                dt_ass = datetime.strptime(str(riga["DATA_ASSEGNO"]).strip(), "%d/%m/%Y").date()
                if (oggi_dt - dt_ass).days >= 1:
                    df.at[idx, "STATO_VERSAMENTO"] = "Versato OK"
                    da_aggiornare = True
                    print(f"✅ Assegno N. {riga['NUMERO_ASSEGNO']} marcato come VERSATO OK")
            except Exception: pass

    # ⏱️ 2. TRAVASO NELLO STORICO (SE VERSATO OK DA PIÙ DI 100 GIORNI)
    indici_spurgo = []
    righe_spurgo = []
    for idx, riga in df.iterrows():
        if str(riga["STATO_VERSAMENTO"]).strip() == "Versato OK" and riga["DATA_ASSEGNO"]:
            try:
                dt_ass = datetime.strptime(str(riga["DATA_ASSEGNO"]).strip(), "%d/%m/%Y").date()
                if (oggi_dt - dt_ass).days >= 100:
                    indici_spurgo.append(idx)
                    righe_spurgo.append(riga)
            except: pass

    if indici_spurgo:
        print(f"📦 Sposto {len(indici_spurgo)} assegni vecchi nell'Archivio Storico...")
        df_st = pd.DataFrame()
        sha_st = ""
        res_st = requests.get(URL_STORICO, headers=headers, timeout=15)
        if res_st.status_code == 200:
            j_st = res_st.json()
            sha_st = j_st.get("sha", "")
            df_st = pd.read_excel(io.BytesIO(base64.b64decode(j_st["content"])), dtype={"NUMERO_ASSEGNO": str}).fillna("")
            
        df_st_nuovo = pd.concat([df_st, pd.DataFrame(righe_spurgo)], ignore_index=True).fillna("")
        
        out_st = io.BytesIO()
        with pd.ExcelWriter(out_st, engine='openpyxl') as wr_st:
            df_st_nuovo.to_excel(wr_st, index=False)
            wr_st.sheets["Sheet1"].views.sheetView[0].showGridLines = True
            
        payload_st = {"message": "🤖 [Archivio] Travaso automatico 100 giorni", "content": base64.b64encode(out_st.getvalue()).decode('utf-8'), "branch": "main"}
        if sha_st: payload_st["sha"] = sha_st
        requests.put(URL_STORICO, json=payload_st, headers=headers, timeout=15)
        
        df = df.drop(indici_spurgo).reset_index(drop=True)
        da_aggiornare = True

    # SALVATAGGIO CLOUD PERMANENTE SUL REGISTRO GIORNALIERO
    if da_aggiornare:
        out_c = io.BytesIO()
        with pd.ExcelWriter(out_c, engine='openpyxl') as wr_c:
            df.to_excel(wr_c, index=False)
            wr_c.sheets["Sheet1"].views.sheetView[0].showGridLines = True
            
        payload_c = {"message": "🤖 [Controllo Notturno] Aggiornamento Stati Assegni", "content": base64.b64encode(out_c.getvalue()).decode('utf-8'), "branch": "main", "sha": dati_j.get("sha", "")}
        res_put = requests.put(URL_TABELLONE, json=payload_c, headers=headers, timeout=15)
        
        # 🎯 SINTASSI CORRETTA (Sostituito il vecchio in:)
        if res_p_c.status_code == 200 or res_p_c.status_code == 201:
            print("💾 Registro aggiornato con successo su GitHub!")
        else:
            print(f"❌ Errore durante il salvataggio su GitHub: HTTP {res_put.status_code}")
    else:
        print("✨ Nessun assegno da modificare questa notte.")

if __name__ == "__main__":
    esegui_controllo_notturno()
