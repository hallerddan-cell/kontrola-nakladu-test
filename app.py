import streamlit as st
import pandas as pd
import numpy as np
import datetime
import calendar
import io
import re
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config(page_title="BTT: Kontrola a Očištění Zakázek", layout="wide")

# ==========================================
# 1. AUTOMATICKÉ SJEDNOCENÍ NÁZVŮ SLOUPCŮ
# ==========================================
def unifikuj_sloupce(df):
    """Sjednotí klíčové názvy sloupců a ponechá všechny ostatní sloupce nedotčené."""
    mapovani = {}
    
    # Číslo / Kód zakázky
    for c in ['Zakázka', 'Kód', 'Číslo zakázky']:
        if c in df.columns:
            mapovani[c] = 'Zakázka'
            break
            
    # Datum vytvoření
    for c in ['Datum vytvoření', 'Datum vytvoř.', 'Datum']:
        if c in df.columns:
            mapovani[c] = 'Datum vytvoření'
            break
            
    # Referent / BD
    for c in ['BD', 'Referent', 'Jméno referenta']:
        if c in df.columns:
            mapovani[c] = 'BD'
            break
            
    # Náklady
    for c in ['Náklady', 'Náklad']:
        if c in df.columns:
            mapovani[c] = 'Náklady'
            break

    # Název organizace / klienta
    for c in ['Název org.', 'Název org', 'Organizace', 'Klient']:
        if c in df.columns:
            mapovani[c] = 'Název org.'
            break

    return df.rename(columns=mapovani)

# ==========================================
# 2. MAPOVÁNÍ E-MAILŮ (JMÉNA + SYSTÉMOVÉ KÓDY)
# ==========================================
EMAILY_REFERENTI = {
    # Celá jména referentů
    "Martínková Hana": "martinkova@oktours.cz",
    "Okrouhlíková Iva": "okrouhlikova@oktours.cz",
    "Novohradská Lucie": "lucie.novohradska@nemoletenky.cz",
    "Tlamichová Rita": "tlamichova@oktours.cz",
    "Haller Dan": "haller@oktours.cz",
    "Daniel Haller": "haller@oktours.cz",
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
    "Trimidal Joshua": "Trimidal@oktours.cz",

    # Systémové zkratky a číselné kódy referentů
    "HM": "martinkova@oktours.cz",
    "IO": "okrouhlikova@oktours.cz",
    "Z02424": "tlamichova@oktours.cz",
    "Z12235": "pesatova@oktours.cz",
    "Z12228": "gazdova@oktours.cz",
    "Z56340": "vidermanova@oktours.cz",
    "Z56261": "trebicky@oktours.cz",
    "Z69805": "nekola@oktours.cz",
    "Z76612": "jandosovap@gmail.com",
    "Z12231": "delta.letenky@gmail.com",
    "Z76611": "adamcova@oktours.cz",
    "Z72348": "flenerova@oktours.cz",
    "Z12230": "zuzankova@oktours.cz",
    "Z12239": "Chumpitaz@oktours.cz",
    "ZC0003": "lucie.novohradska@nemoletenky.cz",
}

def najdi_email_referenta(identifikator):
    """Vyhledá e-mail podle systémového kódu, jména nebo prohozeného pořadí jména a příjmení."""
    if not identifikator or pd.isna(identifikator):
        return None
    val = str(identifikator).strip()
    
    # 1. Přímá shoda (kód např. 'IO', 'Z02424' nebo 'Martínková Hana')
    if val in EMAILY_REFERENTI:
        return EMAILY_REFERENTI[val]
        
    # 2. Shoda po otočení jména a příjmení (např. 'Hana Martínková')
    casti = val.split()
    if len(casti) >= 2:
        otocene = f"{casti[1]} {casti[0]}"
        if otocene in EMAILY_REFERENTI:
            return EMAILY_REFERENTI[otocene]
            
    # 3. Shoda podle příjmení
    if len(casti) >= 1:
        for k, v in EMAILY_REFERENTI.items():
            if casti[0].lower() in k.lower():
                return v
                
    return None

def nacisti_iqy_obsah(file_bytes_or_str):
    content = file_bytes_or_str.decode('utf-8', errors='ignore') if isinstance(file_bytes_or_str, bytes) else file_bytes_or_str
    urls = re.findall(r'https?://[^\s]+', content)
    if not urls:
        raise ValueError("V .iqy souboru nebyla nalezena platná URL adresa.")
    url = urls[0]
    res = requests.get(url)
    try:
        return pd.read_excel(io.BytesIO(res.content))
    except Exception:
        return pd.read_html(io.BytesIO(res.content))[0]

# ==========================================
# 3. HLAVNÍ WEBOVÉ ROZHRANÍ
# ==========================================
st.title("📊 BTT: Kontrola a Očištění Zakázek")

zdroj = st.radio(
    "Vyberte zdroj dat:", 
    ["Nahrajte ostrý soubor (.iqy nebo .xlsx)", "Použít soubor 'dotaz.iqy' z repozitáře", "Použít fiktivní testovací data (test_data.csv)"],
    horizontal=True
)

df = None

if zdroj == "Nahrajte ostrý soubor (.iqy nebo .xlsx)":
    uploaded = st.file_uploader("Vyberte soubor .iqy nebo .xlsx", type=["iqy", "xlsx"])
    if uploaded is not None:
        try:
            if uploaded.name.endswith(".iqy"):
                df = nacisti_iqy_obsah(uploaded.getvalue())
            else:
                df = pd.read_excel(uploaded)
            st.success(f"Soubor '{uploaded.name}' úspěšně načten ({len(df)} řádků).")
        except Exception as e:
            st.error(f"Chyba při zpracování souboru: {e}")

elif zdroj == "Použít soubor 'dotaz.iqy' z repozitáře":
    try:
        with open("dotaz.iqy", "r", encoding="utf-8") as f:
            df = nacisti_iqy_obsah(f.read())
        st.success(f"Data z 'dotaz.iqy' úspěšně načtena ({len(df)} řádků).")
    except Exception as e:
        st.error(f"Nelze načíst 'dotaz.iqy': {e}")

else:
    try:
        df = pd.read_csv("test_data.csv")
        st.success("Načtena fiktivní testovací data.")
    except Exception as e:
        st.error(f"Chyba při načítání test_data.csv: {e}")

if df is not None:
    # Aplikace automatického sjednocení názvů sloupců
    df = unifikuj_sloupce(df)

    st.divider()
    st.subheader("1. Nastavení vyhodnocovaného období")
    
    col1, col2 = st.columns(2)
    today = datetime.date.today()
    first_curr = today.replace(day=1)
    prev_month_date = first_curr - datetime.timedelta(days=1)
    
    with col1:
        vybrany_rok = st.number_input("Rok", min_value=2020, max_value=2030, value=prev_month_date.year)
    with col2:
        vybrany_mesic = st.selectbox("Měsíc", options=list(range(1, 13)), index=prev_month_date.month - 1, format_func=lambda m: f"{m}. měsíc")

    col_datum = "Datum vytvoření"
    col_zakazka = "Zakázka"
    col_naklady = "Náklady"
    col_referent = "BD"

    potrebne = [col_datum, col_zakazka, col_naklady, col_referent]
    chybi_cols = [c for c in potrebne if c not in df.columns]

    if chybi_cols:
        st.error(f"V souboru chybí tyto požadované sloupce: {chybi_cols}")
        st.info(f"Dostupné sloupce v souboru: {list(df.columns)}")
    else:
        # Časový filtr podle data vytvoření
        start_date = datetime.date(vybrany_rok, vybrany_mesic, 1)
        end_date = datetime.date(vybrany_rok, vybrany_mesic, calendar.monthrange(vybrany_rok, vybrany_mesic)[1])

        df["Datum_dt"] = pd.to_datetime(df[col_datum], errors='coerce')
        mask_datum = (df["Datum_dt"].dt.date >= start_date) & (df["Datum_dt"].dt.date <= end_date)
        df_mesic = df[mask_datum].copy()

        # Filtr číselné řady (kód "00" + kód daného měsíce)
        kod_mesice_str = f"{vybrany_mesic:02d}"
        df_mesic["Kod_Zakazky"] = df_mesic[col_zakazka].astype(str).str.strip().str[5:7]

        platne_kody = ["00", kod_mesice_str]
        mask_platne = df_mesic["Kod_Zakazky"].isin(platne_kody)
