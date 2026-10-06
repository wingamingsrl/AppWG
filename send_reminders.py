import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pandas as pd

FILE_STORICO = "storico_ferie.xlsx"
FILE_TECNICI = "elenco_tecnici.xlsx"
EMAIL_MITTENTE_GMAIL = "wingamingsrl@gmail.com"
EMAIL_MANUELA_RICEVENTE = "manuela.arigoni@wingaming.it"

def invia_mail_promemoria_notturno(email_destinatario, oggetto, corpo, pass_gmail):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_MITTENTE_GMAIL
        msg['To'] = email_destinatario
        msg['Subject'] = oggetto
        msg.attach(MIMEText(corpo, 'plain'))
        
        # Connessione protetta IP diretto Google
        server = smtplib.SMTP_SSL('64.233.184.108', 465, timeout=10)
        server.login(EMAIL_MITTENTE_GMAIL, pass_gmail)
        server.sendmail(EMAIL_MITTENTE_GMAIL, [email_destinatario, EMAIL_MANUELA_RICEVENTE], msg.as_string())
        server.quit()
        print(f"📧 Promemoria inviato con successo a: {email_destinatario}")
    except Exception as e:
        print(f"❌ Errore invio mail a {email_destinatario}: {str(e)}")

def controlla_e_invia_promemoria():
    # Recupera la password dai Secrets di GitHub
    pass_gmail = os.environ.get("GMAIL_PASSWORD")
    if not pass_gmail:
        print("❌ Password Gmail non trovata nelle variabili d'ambiente.")
        return

    if not os.path.exists(FILE_STORICO) or not os.path.exists(FILE_TECNICI):
        print("❌ File storici o tecnici mancanti. Impossibile procedere.")
        return

    try:
        df_storico = pd.read_excel(FILE_STORICO).fillna("")
        df_tecnici = pd.read_excel(FILE_TECNICI).fillna("")
        
        # Mappa i nomi dei tecnici alle loro e-mail reali per sicurezza
        mappa_email_tecnici = {str(r["NOME"]).strip().lower(): str(r["EMAIL"]).strip().lower() for _, r in df_tecnici.iterrows()}
        
        oggi = datetime.now().date()
        
        for _, row in df_storico.iterrows():
            try:
                d_i = datetime.strptime(str(row["INIZIO_FERIE"]).strip(), "%d-%m-%Y").date()
                d_f = datetime.strptime(str(row["FINE_FERIE"]).strip(), "%d-%m-%Y").date()
                tecnico_nome = str(row["TECNICO"]).strip()
                locale = str(row["LOCALE"]).strip()
                
                # Cerca l'email del tecnico, altrimenti usa quella di Manuela come ruota di scorta
                email_tecnico = mappa_email_tecnici.get(tecnico_nome.lower(), EMAIL_MANUELA_RICEVENTE)
                
                # CASO 1: Mancano esattamente 3 giorni all'inizio della chiusura
                if d_i - oggi == timedelta(days=3):
                    oggetto = f"⚠️ PROMEMORIA: Chiusura Ferie tra 3 Giorni - {locale}"
                    corpo = f"Ciao {tecnico_nome},\n\nQuesto è un promemoria automatico logistico WinGaming.\n\n📍 Il locale: {locale}\n📅 Inizierà il periodo di chiusura il giorno: {row['INIZIO_FERIE']}\n\nSi prega di organizzare i giri logistici di conseguenza.\n\nWINGAMING SRL"
                    invia_mail_promemoria_notturno(email_tecnico, oggetto, corpo, pass_gmail)
                    
                # CASO 2: Mancano esattamente 3 giorni alla riapertura
                if d_f - oggi == timedelta(days=3):
                    oggetto = f"🚚 PROMEMORIA: Riapertura Locale tra 3 Giorni - {locale}"
                    corpo = f"Ciao {tecnico_nome},\n\nQuesto è un promemoria automatico logistico WinGaming.\n\n📍 Il locale: {locale}\n🚚 Riaprirà ufficialmente il giorno: {row['FINE_FERIE']}\n\nSi prega di verificare i giri logistici di rientro.\n\nWINGAMING SRL"
                    invia_mail_promemoria_notturno(email_tecnico, oggetto, corpo, pass_gmail)
                    
            except Exception:
                continue
    except Exception as e:
        print(f"Errore lettura tabelle: {str(e)}")

if __name__ == "__main__":
    controlla_e_invia_promemoria()
