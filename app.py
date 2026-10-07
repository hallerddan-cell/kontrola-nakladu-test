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
    "Martínková Hana": "martinkova@oktours.cz",
    "Okrouhlíková Iva": "okrouhlikova@oktours.cz",
    "Novohradská Lucie": "lucie.novohradska@nemoletenky.cz",
    "Tlamichová Rita": "tlamichova@oktours.cz",
    "Haller Dan": "haller@oktours.cz",
    "Adamcová Marta": "adamcova@oktours.cz",
    "Gazdová Pavlína": "gazdova@oktours.cz",
    "Jandošová Petra": "jandosovap@gmail.com",
    "Matějková Ivona": "delta.letenky@gmail.com",
    "Pešatová Helena": "pesatova@oktours.cz",
    "Třebický Tomáš": "trebicky@oktours.cz",
    "Nekola Tomáš": "nekola@oktours.cz",
    "Vidermanova Sabina": "vidermanova@oktours.cz",
    "Flenerová Helena": "flenerova@oktours.cz",
    "Chumpitaz Pavlína": "Chumpitaz@oktours.cz",
    "Zuzánková Nikola": "zuzankova@oktours.cz",
    "Trimidal Joshua": "Trimidal@oktours.cz"
}

# Nastavení přihlášení k e-mailovému serveru ze Secrets
SMTP_SENDER = os.getenv("SMTP_SENDER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com") # Změňte na smtp.office365.com u M365


def najdi_email_referenta(jmeno_referenta):
    """Najde e-mail bez ohledu na pořadí Jméno/Příjmení."""
    if not jmeno_referenta or pd.isna(jmeno_referenta):
        return None
    
    jmeno_std = str(jmeno_referenta).strip()
    
    # 1. Přímá shoda (např. "Martínková Hana")
    if jmeno_std in EMAILY_REFERENTI:
        return EMAILY_REFERENTI[jmeno_std]
        
    # 2. Shoda po otočení jména a příjmení (např. "Hana Martínková" -> "Martínková Hana")
    casti = jmeno_std.split()
    if len(casti) >= 2:
        otocene = f"{casti[1]} {casti[0]}"
        if otocene in EMAILY_REFERENTI:
            return EMAILY_REFERENTI[otocene]
            
    # 3. Vyhledání podle příjmení
    for k, v in EMAILY_REFERENTI.items():
        if casti[0].lower() in k.lower():
            return v
            
    return None


# ==========================================
# 2. NAČTENÍ DATA Z .IQY SOUBORU NEBO URL
# ==========================================
def nacti_data_z_iqy(iqy_file_path="dotaz.iqy"):
    """Načte URL z .iqy souboru a stáhne čerstvá data ze systému."""
    print("📥 Načítám čerstvá data z webového dotazu (.iqy)...")
    
    if not os.path.exists(iqy_file_path):
        raise FileNotFoundError(f"Soubor '{iqy_file_path}' nebyl nalezen v repozitáři.")

    with open(iqy_file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    urls = re.findall(r'https?://[^\s]+', content)
    if not urls:
        raise ValueError("V .iqy souboru nebyla nalezena platná URL adresa.")
    
    url = urls[0]
    print(f"🔗 Připojuji se k systémové URL: {url[:60]}...")
    
    response = requests.get(url)
    response.raise_for_status()
    
    # Pokus o načtení Excelu nebo HTML tabulky
    try:
        df = pd.read_excel(io.BytesIO(response.content))
    except Exception:
        df = pd.read_html(io.BytesIO(response.content))[0]
        
    print(f"✅ Data úspěšně načtena ({len(df)} řádků).")
    return df


# ==========================================
# 3. FILTRACE MĚSÍCE A KONTROLA ČÍSEL ZAKÁZEK
# ==========================================
def zpracuj_a_zkonstruuj_vykaz(df):
    today = datetime.date.today()
    
    # Výpočet předchozího celého kalendářního měsíce (např. spuštění v listopadu -> vyhodnocuje říjen)
    first_day_curr = today.replace(day=1)
    last_day_prev = first_day_curr - datetime.timedelta(days=1)
    
    target_year = last_day_prev.year
    target_month = last_day_prev.month  # 1 až 12
    
    start_date = datetime.date(target_year, target_month, 1)
    end_date = datetime.date(target_year, target_month, calendar.monthrange(target_year, target_month)[1])
    
    print(f"🗓️ Vyhodnocuji období předchozího měsíce: {target_month}/{target_year} ({start_date} až {end_date})")
    
    # A) Filtr podle sloupce 'Datum vytvoření'
    df["Datum_dt"] = pd.to_datetime(df["Datum vytvoření"], errors='coerce')
    mask_datum = (df["Datum_dt"].dt.date >= start_date) & (df["Datum_dt"].dt.date <= end_date)
    df_mesic = df[mask_datum].copy()
    
    # B) Dynamická kontrola čísla zakázky pro daný měsíc (kód měsíce na 6. a 7. pozici)
    kod_mesice_str = f"{target_month:02d}"  # "01", "02", ..., "10", "11", "12"
    df_mesic["Kod_Zakazky"] = df_mesic["Zakázka"].astype(str).str.strip().str[5:7]
    
    # Povolené kódy: "00" (standardní řada s nulami) NEBO kód daného měsíce (např. "10" pro říjen)
    platne_kody = ["00", kod_mesice_str]
    mask_platne = df_mesic["Kod_Zakazky"].isin(platne_kody)
    
    df_platne = df_mesic[mask_platne].copy()
    df_vyrazene = df_mesic[~mask_platne].copy()
    
    print(f"   • Celkem zakázek v daném měsíci: {len(df_mesic)}")
    print(f"   • Platné zakázky zařazené do zpracování: {len(df_platne)}")
    print(f"   • Vyřazeno zakázek s kódem jiného měsíce: {len(df_vyrazene)}")
    
    # C) Detekce chybějících nákladů
    mask_chybi_naklady = (
        df_platne["Náklady"].isna() | 
        (df_platne["Náklady"].astype(str).str.strip() == "") | 
        (df_platne
