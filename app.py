import pandas as pd
import numpy as np
import datetime
import calendar
import re
import requests
import smtplib
import io
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ==========================================
# 1. MAPOVÁNÍ E-MAILŮ REFERENTŮ (sloupec BD)
# ==========================================
EMAILY_REFERENTI = {
    "Jandošová Petra": "petra.jandosova@oktours.cz",
    "Matějková Ivona": "ivona.matejkova@oktours.cz",
    "Nekola Tomáš": "tomas.nekola@oktours.cz",
    "Třebický Tomáš": "tomas.trebicky@oktours.cz",
    # Sem můžete doplnit další referenty podle potřeby
}

# Přihlašovací údaje k e-mailu (načítají se bezpečně ze Secrets)
SMTP_SENDER = os.getenv("SMTP_SENDER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com") # Nebo smtp.office365.com


# ==========================================
# 2. NAČTENÍ DATA Z .IQY SOUBORU NEBO URL
# ==========================================
def nacti_data_z_iqy(iqy_file_path="dotaz.iqy"):
    """Načte URL z IQY souboru a stáshne aktuální data."""
    print("📥 Načítám čerstvá data z webového dotazu (.iqy)...")
    
    if not os.path.exists(iqy_file_path):
        raise FileNotFoundError(f"Soubor '{iqy_file_path}' nebyl nalezen.")

    with open(iqy_file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    urls = re.findall(r'https?://[^\s]+', content)
    if not urls:
        raise ValueError("V .iqy souboru nebyla nalezena platná URL adresa.")
    
    url = urls[0]
    response = requests.get(url)
    
    # Načtení dat do DataFramu
    try:
        df = pd.read_excel(io.BytesIO(response.content))
    except Exception:
        df = pd.read_html(io.BytesIO(response.content))[0]
        
    print(f"✅ Data úspěšně načtena ({len(df)} řádků).")
    return df


# ==========================================
# 3. FILTRACE MĚSÍCE A KONTROLA ČÍSEL ZAKÁZEK
# ==========================================
def zpracuj_a_zkonktroluj_vykaz(df):
    today = datetime.date.today()
    
    # Výpočet předchozího kalendářního měsíce
    first_day_curr = today.replace(day=1)
    last_day_prev = first_day_curr - datetime.timedelta(days=1)
    
    target_year = last_day_prev.year
    target_month = last_day_prev.month  # 1 až 12
    
    start_date = datetime.date(target_year, target_month, 1)
    end_date = datetime.date(target_year, target_month, calendar.monthrange(target_year, target_month)[1])
    
    print(f"🗓️ Vyhodnocuji měsíc: {target_month}/{target_year} ({start_date} až {end_date})")
    
    # A) Filtr podle datumu vytvoření
    df["Datum vytvoření"] = pd.to_datetime(df["Datum vytvoření"], errors='coerce')
    mask_datum = (df["Datum vytvoření"].dt.date >= start_date) & (df["Datum vytvoření"].dt.date <= end_date)
    df_mesic = df[mask_datum].copy()
    
    # B) Dynamická kontrola čísla zakázky pro daný měsíc (kód měsíce na 6. a 7. pozici)
    kod_mesice_str = f"{target_month:02d}"  # Např. "01", "02", ..., "10", "11", "12"
    df_mesic["Kod_Zakazky"] = df_mesic["Zakázka"].astype(str).str.strip().str[5:7]
    
    # Povoleny jsou zakázky se standardní řadou "00" NEBO s kódem daného měsíce
    platne_kody = ["00", kod_mesice_str]
    mask_platne = df_mesic["Kod_Zakazky"].isin(platne_kody)
    
    df_platne = df_mesic[mask_platne].copy()
    df_vyrazene = df_mesic[~mask_platne].copy()
    
    print(f"   • Celkem zakázek v měsíci: {len(df_mesic)}")
    print(f"   • Vyřazeno zakázek z jiných měsíců: {len(df_vyrazene)}")
    
    # C) Detekce chybějících nákladů
    mask_chybi_naklady = (
        df_platne["Náklady"].isna() | 
        (df_platne["Náklady"].astype(str).str.strip() == "") | 
        (df_platne["Náklady"] == 0)
    )
    
    df_chybi = df_platne[mask_chybi_naklady].copy()
    print(f"   • Nalezeno zakázek bez nákladů: {len(df_chybi)}")
    
    return df_chybi, target_month, target_year


# ==========================================
# 4. ROZESLÁNÍ UPOZORNĚNÍ / UPOMÍNEK REFERENTŮM
# ==========================================
def rozeslat_emaily(df_chybi, target_month, target_year):
    if df_chybi.empty:
        print("🎉 Všechny zakázky mají doplněné náklady! Žádné e-maily nebyly odeslány.")
        return

    today = datetime.date.today()
    
    # Rozlišení zda jde o první výzvu (5. den) nebo upomínku (7. den a později)
    je_upominka = today.day >= 7
    prefix_predmetu = "⚠️ UPOZORNĚNÍ (2. VÝZVA)" if je_upominka else "Upozornění"
    
    skupiny = df_chybi.groupby("BD")
    
    for referent, zakazky in skupiny:
        email_prijemce = EMAILY_REFERENTI.get(referent)
        
        if not email_prijemce:
            print(f"⚠️ Varování: Pro referenta '{referent}' nebyla nalezena e-mailová adresa.")
            continue
            
        # Generování HTML tabulky zakázek
        tabulka_html = zakazky[["Zakázka", "Datum vytvoření", "Název org."]].to_html(index=False)
        
        predmet = f"{prefix_predmetu}: Chybějící náklady u zakázek za {target_month}/{target_year}"
        
        upozorneni_text = (
            "<p style='color: red; font-weight: bold;'>Toto je opakovaná výzva! Náklady stále nebyly doplněny.</p>"
            if je_upominka else ""
        )
        
        html_obsah = f"""
        <html>
        <head>
            <style>
                table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
            </style>
        </head>
        <body>
            <p>Dobrý den, {referent},</p>
            {upozorneni_text}
            <p>při automatické kontrole uzávěrky za <b>{target_month}/{target_year}</b> byly u Vašich zakázek zjištěny <b>chybějící náklady</b>:</p>
            {tabulka_html}
            <p>Prosíme o jejich co nejrychlejší doplnění do systému.</p>
            <br>
            <p><i>Tato zpráva byla vygenerována automaticky.</i></p>
        </body>
        </html>
        """
        
        msg = MIMEMultipart()
        msg['From'] = SMTP_SENDER
        msg['To'] = email_prijemce
        msg['Subject'] = predmet
        msg.attach(MIMEText(html_obsah, 'html', 'utf-8'))
        
        try:
            port = 587
            server = smtplib.SMTP(SMTP_SERVER, port)
            server.starttls()
            server.login(SMTP_SENDER, SMTP_PASSWORD)
            server.sendmail(SMTP_SENDER, [email_prijemce], msg.as_string())
            server.quit()
            print(f"📧 E-mail doručen pro: {referent} ({email_prijemce})")
        except Exception as e:
            print(f"❌ Chyba při odesílání e-mailu pro {referent}: {e}")


# ==========================================
# HLAVNÍ SPUŠTĚNÍ
# ==========================================
if __name__ == "__main__":
    try:
        df_raw = nacti_data_z_iqy("dotaz.iqy")
        df_chybi, mesic, rok = zpracuj_a_zkonktroluj_vykaz(df_raw)
        rozeslat_emaily(df_chybi, mesic, rok)
        print("🎉 Automatická rutina dokončena.")
    except Exception as e:
        print(f"❌ Chyba během automatické rutiny: {e}")
