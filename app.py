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
    for c in ['Datum vytvoření', 'Datum vytvoř.']:
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

    # Název org.
    for c in ['Název org.', 'Název org', 'Organizace', 'Klient']:
        if c in df.columns:
            mapovani[c] = 'Název org.'
            break

    return df.rename(columns=mapovani)

# ==========================================
# 2. MAPOVÁNÍ E-MAILŮ REFERENTŮ
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

def najdi_email_referenta(jmeno_referenta):
    if not jmeno_referenta or pd.isna(jmeno_referenta):
        return None
    jmeno_std = str(jmeno_referenta).strip()
    if jmeno_std in EMAILY_REFERENTI:
        return EMAILY_REFERENTI[jmeno_std]
    casti = jmeno_std.split()
    if len(casti) >= 2:
        otocene = f"{casti[1]} {casti[0]}"
        if otocene in EMAILY_REFERENTI:
            return EMAILY_REFERENTI[otocene]
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
    # Aplikujeme automatické sjednocení názvů klíčových sloupců
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
        # Časový filtr
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

        df_ocistene = df_mesic[mask_platne].copy()
        df_vyrazene = df_mesic[~mask_platne].copy()

        # Označení stavu nákladů
        mask_chybi = (
            df_ocistene[col_naklady].isna() | 
            (df_ocistene[col_naklady].astype(str).str.strip() == "") | 
            (df_ocistene[col_naklady].astype(str).str.strip() == "0") |
            (df_ocistene[col_naklady] == 0)
        )
        
        df_ocistene["Stav nákladů"] = np.where(mask_chybi, "❌ Chybí náklad", "✅ V pořádku")
        df_ocistene["Datum vytvoření"] = df_ocistene["Datum_dt"].dt.strftime('%d.%m.%Y')

        # Odstranění pomocných sloupců před exportem, aby data byla čistá
        df_export_ocistene = df_ocistene.drop(columns=["Datum_dt", "Kod_Zakazky"], errors="ignore")
        df_chybi = df_export_ocistene[mask_chybi].copy()

        st.divider()
        st.subheader(f"📊 Výsledky pro {vybrany_mesic}/{vybrany_rok}")

        m1, m2, m3 = st.columns(3)
        m1.metric("Celkem v měsíci", len(df_mesic))
        m2.metric("Vyřazené zakázky (jiný měsíc)", len(df_vyrazene))
        m3.metric("⚠️ Chybějící náklady", len(df_chybi))

        if not df_vyrazene.empty:
            with st.expander("ℹ️ Zobrazit vyřazené zakázky z jiných měsíců"):
                st.dataframe(df_vyrazene[[col_zakazka, col_datum, col_referent, "Kod_Zakazky"]], use_container_width=True)

        st.subheader("📋 Očištěná data (Kompletní tabulka se všemi sloupci)")
        st.dataframe(df_export_ocistene, use_container_width=True)

        # STAŽENÍ SOUBORŮ EXCEL (VŠECHNY SLOUPCE)
        st.divider()
        st.subheader("📥 Stažení výstupů v Excelu")
        
        col_d1, col_d2 = st.columns(2)
        
        # 1. Stažení VŠECH očištěných zakázek v PLNOU ŠÍŘI sloupců
        out_all = io.BytesIO()
        with pd.ExcelWriter(out_all, engine='openpyxl') as writer:
            df_export_ocistene.to_excel(writer, index=False, sheet_name=f"Ocistene_{vybrany_mesic}_{vybrany_rok}")
        
        with col_d1:
            st.download_button(
                label="🟢 Stáhnout VŠECHNY očištěné zakázky (.xlsx)",
                data=out_all.getvalue(),
                file_name=f"BTT_Ocistena_Data_Komplet_{vybrany_mesic}_{vybrany_rok}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 2. Stažení POUZE zakázek bez nákladů v PLNOU ŠÍŘI sloupců
        if not df_chybi.empty:
            out_chybi = io.BytesIO()
            with pd.ExcelWriter(out_chybi, engine='openpyxl') as writer:
                df_chybi.to_excel(writer, index=False, sheet_name="Chybejici_naklady")
            
            with col_d2:
                st.download_button(
                    label="🔴 Stáhnout POUZE zakázky CHYBĚJÍCÍ NÁKLAD (.xlsx)",
                    data=out_chybi.getvalue(),
                    file_name=f"BTT_Chybejici_Naklady_Komplet_{vybrany_mesic}_{vybrany_rok}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ROZESLÁNÍ EMAILŮ (VOLITELNÉ)
        st.divider()
        st.subheader("📧 Volitelné odeslání e-mailů referentům")
        st.caption("E-maily odcházejí pouze po stisknutí tlačítka níže.")

        muj_email = st.text_input("Zadejte e-mail pro testovací doručení (nebo nechte prázdné pro rozeslání referentům)", "")

        if st.button("🚀 Odeslat e-mailová upozornění"):
            if df_chybi.empty:
                st.success("Všechny zakázky mají vyplněné náklady, e-maily není potřeba odesílat.")
            else:
                st.info("Odesílám upozornění...")
                try:
                    sender = st.secrets["smtp"]["sender_email"]
                    pwd = st.secrets["smtp"]["password"]
                    server_name = st.secrets["smtp"].get("server", "smtp.gmail.com")
                except Exception:
                    st.error("Chybí nastavení SMTP v Secrets!")
                    sender = None

                if sender:
                    skupiny = df_chybi.groupby(col_referent)
                    for referent, zakazky in skupiny:
                        prijemce = muj_email if muj_email.strip() else najdi_email_referenta(referent)
                        if not prijemce:
                            st.warning(f"Chybí e-mail pro: {referent}")
                            continue

                        tab_html = zakazky[[col_zakazka, col_datum, "Název org."]].to_html(index=False)
                        msg_body = f"<h2>Upozornění: Chybějící náklady za {vybrany_mesic}/{vybrany_rok}</h2>"
                        msg_body += f"<p>Dobrý den, {referent}, u následujících zakázek chybí náklady:</p>{tab_html}"

                        msg = MIMEMultipart()
                        msg['From'] = sender
                        msg['To'] = prijemce
                        msg['Subject'] = f"Upozornění: Chybějící náklady za {vybrany_mesic}/{vybrany_rok}"
                        msg.attach(MIMEText(msg_body, 'html', 'utf-8'))

                        try:
                            s = smtplib.SMTP(server_name, 587)
                            s.starttls()
                            s.login(sender, pwd)
                            s.sendmail(sender, [prijemce], msg.as_string())
                            s.quit()
                            st.success(f"E-mail doručen na: {prijemce} ({referent})")
                        except Exception as ex:
                            st.error(f"Chyba při odesílání pro {referent}: {ex}")
