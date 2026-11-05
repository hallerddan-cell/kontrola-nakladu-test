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

st.set_page_config(page_title="BTT: Kontrola Nákladů a Route 1", layout="wide")

# ==========================================
# 1. AUTOMATICKÉ SJEDNOCENÍ NÁZVŮ SLOUPCŮ
# ==========================================
def unifikuj_sloupce(df):
    """Sjednotí klíčové sloupce a zachová všechny ostatní původní sloupce."""
    mapovani = {}
    
    # Kód zakázky
    for c in ['Zakázka', 'Kód', 'Číslo zakázky']:
        if c in df.columns:
            mapovani[c] = 'Zakázka'
            break
            
    # Datum vytvoření
    for c in ['Datum vytvoření', 'Datum vytvoř.', 'Datum']:
        if c in df.columns:
            mapovani[c] = 'Datum vytvoření'
            break
            
    # Referent - priorita pro skutečný kód/jméno tvůrce (nikoli 'Jméno referenta')
    for c in ['BD', 'Referent']:
        if c in df.columns:
            mapovani[c] = 'BD'
            break
    if 'BD' not in mapovani.values() and 'Jméno referenta' in df.columns:
        mapovani['Jméno referenta'] = 'BD'
            
    # Náklady
    for c in ['Náklady', 'Náklad']:
        if c in df.columns:
            mapovani[c] = 'Náklady'
            break

    # Route 1
    for c in ['Route 1', 'Route1', 'Rout 1', 'Trasa']:
        if c in df.columns:
            mapovani[c] = 'Route 1'
            break

    # Služba
    for c in ['Služba', 'Druh služby', 'Typ služby']:
        if c in df.columns:
            mapovani[c] = 'Služba'
            break

    # Název org.
    for c in ['Název org.', 'Název org', 'Organizace', 'Klient']:
        if c in df.columns:
            mapovani[c] = 'Název org.'
            break

    return df.rename(columns=mapovani)

# ==========================================
# 2. MAPOVÁNÍ E-MAILŮ (JMÉNA + SYSTÉMOVÉ KÓDY)
# ==========================================
EMAILY_REFERENTI = {
    # Celá jména
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

    # Systémové kódy referentů
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
    "Z02425": "haller@oktours.cz",
}

def najdi_email_referenta(identifikator):
    if not identifikator or pd.isna(identifikator):
        return None
    val = str(identifikator).strip()
    if val in EMAILY_REFERENTI:
        return EMAILY_REFERENTI[val]
    casti = val.split()
    if len(casti) >= 2:
        otocene = f"{casti[1]} {casti[0]}"
        if otocene in EMAILY_REFERENTI:
            return EMAILY_REFERENTI[otocene]
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
# 3. WEBOVÉ ROZHRANÍ
# ==========================================
st.title("📊 BTT: Kontrola Nákladů a Route 1")

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
    df = unifikuj_sloupce(df)

    col_datum = "Datum vytvoření"
    col_zakazka = "Zakázka"
    col_naklady = "Náklady"
    col_route = "Route 1"
    col_referent = "BD"
    col_sluzba = "Služba"

    potrebne = [col_datum, col_zakazka, col_naklady, col_referent]
    chybi_cols = [c for c in potrebne if c not in df.columns]

    if chybi_cols:
        st.error(f"V souboru chybí tyto povinné sloupce: {chybi_cols}")
        st.info(f"Dostupné sloupce v souboru: {list(df.columns)}")
    else:
        # Odstranění prázdných řádků a systémových patiček
        df = df.dropna(subset=[col_zakazka, col_datum]).copy()
        df = df[~df[col_zakazka].astype(str).str.startswith("Počet")].copy()

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

        start_date = datetime.date(vybrany_rok, vybrany_mesic, 1)
        end_date = datetime.date(vybrany_rok, vybrany_mesic, calendar.monthrange(vybrany_rok, vybrany_mesic)[1])

        # Časový filtr s českým formátem DD.MM.YYYY
        df["Datum_dt"] = pd.to_datetime(df[col_datum], dayfirst=True, errors='coerce')
        mask_datum = (df["Datum_dt"].dt.date >= start_date) & (df["Datum_dt"].dt.date <= end_date)
        df_mesic = df[mask_datum].copy()

        # Filtr číselné řady (kód "00" + kód vybraného měsíce)
        kod_mesice_str = f"{vybrany_mesic:02d}"
        df_mesic["Kod_Zakazky"] = df_mesic[col_zakazka].astype(str).str.strip().str[5:7]

        platne_kody = ["00", kod_mesice_str]
        mask_platne = df_mesic["Kod_Zakazky"].isin(platne_kody)

        df_ocistene = df_mesic[mask_platne].copy()
        df_vyrazene = df_mesic[~mask_platne].copy()

        # ==========================================
        # KONTROLA NÁKLADŮ A ROUTE 1
        # ==========================================
        # 1. Kontrola nákladů
        mask_chybi_naklady = (
            df_ocistene[col_naklady].isna() | 
            (df_ocistene[col_naklady].astype(str).str.strip() == "") | 
            (df_ocistene[col_naklady].astype(str).str.strip() == "0") |
            (df_ocistene[col_naklady].astype(str).str.strip() == "0.0") |
            (df_ocistene[col_naklady] == 0)
        )

        # 2. Kontrola Route 1
        if col_route in df_ocistene.columns:
            mask_chybi_route = (
                df_ocistene[col_route].isna() | 
                (df_ocistene[col_route].astype(str).str.strip() == "") | 
                (df_ocistene[col_route].astype(str).str.strip() == "-")
            )
        else:
            mask_chybi_route = pd.Series(False, index=df_ocistene.index)

        # Výjimka pro HELPDESK (u helpdesku se nevyžaduje náklad ani route)
        if col_sluzba in df_ocistene.columns:
            is_helpdesk = df_ocistene[col_sluzba].astype(str).str.strip().str.upper() == "HELPDESK"
            mask_chybi_naklady = mask_chybi_naklady & (~is_helpdesk)
            mask_chybi_route = mask_chybi_route & (~is_helpdesk)

        # Klasifikace výsledného stavu chyby
        conditions = [
            mask_chybi_naklady & mask_chybi_route,
            mask_chybi_naklady & (~mask_chybi_route),
            (~mask_chybi_naklady) & mask_chybi_route
        ]
        choices = [
            "❌ Chybí Náklady + Route 1",
            "❌ Chybí Náklady",
            "❌ Chybí Route 1"
        ]
        
        df_ocistene["Chyba"] = np.select(conditions, choices, default="✅ V pořádku")
        df_ocistene["Datum vytvoření"] = df_ocistene["Datum_dt"].dt.strftime('%d.%m.%Y')

        df_export_ocistene = df_ocistene.drop(columns=["Datum_dt", "Kod_Zakazky"], errors="ignore")
        
        # Filtr na jakoukoliv chybu
        mask_jaka_koliv_chyba = mask_chybi_naklady | mask_chybi_route
        df_chybi = df_export_ocistene[mask_jaka_koliv_chyba].copy()

        # Zobrazení výsledků
        st.divider()
        st.subheader(f"📊 Výsledky kontroly pro {vybrany_mesic}/{vybrany_rok}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Celkem v měsíci", len(df_mesic))
        m2.metric("Vyřazené (jiný měsíc)", len(df_vyrazene))
        m3.metric("⚠️ Zakázky s chybou celkem", len(df_chybi))
        m4.metric("Rozpad chyb", f"N:{mask_chybi_naklady.sum()} | R1:{mask_chybi_route.sum()}")

        if not df_vyrazene.empty:
            with st.expander("ℹ️ Zobrazit vyřazené zakázky z jiných měsíců"):
                st.dataframe(df_vyrazene[[col_zakazka, col_datum, col_referent, "Kod_Zakazky"]], use_container_width=True)

        st.subheader("⚠️ Zakázky vyžadující doplnění (Náklady / Route 1)")
        if df_chybi.empty:
            st.success("🎉 Skvělé! Všechny zakázky mají vyplněné náklady i Route 1.")
        else:
            nahled_cols = [c for c in [col_zakazka, col_datum, col_referent, "Název org.", col_naklady, col_route, "Chyba"] if c in df_chybi.columns]
            st.dataframe(df_chybi[nahled_cols], use_container_width=True)

        st.subheader("📋 Kompletní očištěná data (Všechny zakázky měsíce)")
        st.dataframe(df_export_ocistene, use_container_width=True)

        # STAŽENÍ SOUBORŮ EXCEL
        st.divider()
        st.subheader("📥 Stažení výstupů v Excelu")
        
        col_d1, col_d2 = st.columns(2)
        
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

        if not df_chybi.empty:
            out_chybi = io.BytesIO()
            with pd.ExcelWriter(out_chybi, engine='openpyxl') as writer:
                df_chybi.to_excel(writer, index=False, sheet_name="Chybne_zakazky")
            
            with col_d2:
                st.download_button(
                    label="🔴 Stáhnout POUZE zakázky S CHYBOU (.xlsx)",
                    data=out_chybi.getvalue(),
                    file_name=f"BTT_Zakazky_s_Chybou_{vybrany_mesic}_{vybrany_rok}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # ROZESLÁNÍ EMAILŮ (VOLITELNÉ)
        st.divider()
        st.subheader("📧 Volitelné odeslání e-mailů referentům")
        st.caption("E-maily odcházejí pouze po stisknutí tlačítka níže.")

        muj_email = st.text_input("Zadejte e-mail pro testovací doručení (nechte prázdné pro ostré rozeslání referentům)", "")

        if st.button("🚀 Odeslat e-mailová upozornění"):
            if df_chybi.empty:
                st.success("Žádné zakázky nemají chybu, e-maily není potřeba odesílat.")
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
                            st.warning(f"Chybí e-mail pro referenta/kód: '{referent}'")
                            continue

                        tab_cols = [c for c in [col_zakazka, col_datum, "Název org.", col_naklady, col_route, "Chyba"] if c in zakazky.columns]
                        tab_html = zakazky[tab_cols].to_html(index=False)
                        msg_body = f"""
                        <html>
                        <head>
                            <style>
                                table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
                                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                                th {{ background-color: #f2f2f2; font-weight: bold; }}
                            </style>
                        </head>
                        <body>
                            <p>Dobrý den,</p>
                            <p>při kontrole zakázek za období <b>{vybrany_mesic}/{vybrany_rok}</b> byly u Vašich zakázek zjištěny následující nedostatky (chybějící Náklady nebo Route 1):</p>
                            {tab_html}
                            <p>Prosíme o co nejrychlejší doplnění chybějících údajů do systému.</p>
                        </body>
                        </html>
                        """

                        msg = MIMEMultipart()
                        msg['From'] = sender
                        msg['To'] = prijemce
                        msg['Subject'] = f"Upozornění: Chybějící údaje zakázek za {vybrany_mesic}/{vybrany_rok}"
                        msg.attach(MIMEText(msg_body, 'html', 'utf-8'))

                        try:
                            s = smtplib.SMTP(server_name, 587)
                            s.starttls()
                            s.login(sender, pwd)
                            s.sendmail(sender, [prijemce], msg.as_string())
                            s.quit()
                            st.success(f"E-mail doručen na: {prijemce} (pro '{referent}')")
                        except Exception as ex:
                            st.error(f"Chyba při odesílání pro '{referent}': {ex}")
