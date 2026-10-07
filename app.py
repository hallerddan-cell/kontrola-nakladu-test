import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config(page_title="TEST: Kontrola Nákladů", layout="wide")

# --- FUNKCE PRO ODESLÁNÍ EMAILU PŘES OUTLOOK ---
def odeslat_testovaci_email(muj_email, predmet, html_obsah):
    try:
        sender_email = st.secrets["smtp"]["sender_email"]
        sender_password = st.secrets["smtp"]["password"]
    except Exception:
        return False, "Chybí nastavení SMTP v Secrets (sender_email a password)."

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = muj_email
    msg['Subject'] = f"[TEST] {predmet}"
    msg.attach(MIMEText(html_obsah, 'html', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [muj_email], msg.as_string())
        server.quit()
        return True, f"✅ Testovací e-mail úspěšně odeslán na VÁŠ e-mail: {muj_email}"
    except Exception as e:
        return False, f"❌ Chyba při odesílání: {str(e)}"

# --- HLAVNÍ APLIKACE ---
st.title("🧪 TESTOVACÍ MÓD: Kontrola Chybějících Nákladů")
st.info("Všechny e-maily v tomto testovacím režimu odcházejí VÝHRADNĚ na váš zadaný e-mail!")

# Volba zdroje dat
zdroj_dat = st.radio("Vyberte zdroj dat pro test:", ["Použít fiktivní testovací data (test_data.csv)", "Nahrát vlastní Excel (.xlsx)"])

if zdroj_dat == "Použít fiktivní testovací data (test_data.csv)":
    try:
        df = pd.read_csv("test_data.csv")
        st.success("Načtena fiktivní testovací data.")
    except Exception as e:
        st.error(f"Chyba při načítání test_data.csv: {e}")
        df = None
else:
    uploaded_file = st.file_uploader("Nahrajte testovací Excel", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.success("Excel byl úspěšně nahrán.")
    else:
        df = None

if df is not None:
    st.divider()
    st.subheader("1. Výběr kontrolovaného měsíce")
    
    col1, col2 = st.columns(2)
    with col1:
        vybrany_rok = st.number_input("Rok", min_value=2020, max_value=2030, value=2026)
    with col2:
        vybrany_mesic = st.selectbox("Měsíc", options=list(range(1, 13)), index=8, format_func=lambda m: f"{m}. měsíc (Září)")

    col_datum = "Datum vytvoření"
    col_zakazka = "Zakázka"
    col_naklady = "Náklady"
    col_referent = "BD"

    # Kontrola měsíce
    start_date = datetime.date(vybrany_rok, vybrany_mesic, 1)
    end_date = datetime.date(vybrany_rok, vybrany_mesic, calendar.monthrange(vybrany_rok, vybrany_mesic)[1])

    df[col_datum] = pd.to_datetime(df[col_datum], errors='coerce')
    mask_datum = (df[col_datum].dt.date >= start_date) & (df[col_datum].dt.date <= end_date)
    df_mesic = df[mask_datum].copy()

    # Kontrola čísel zakázek (kód měsíce na 6. a 7. pozici)
    kod_mesice_str = f"{vybrany_mesic:02d}"
    df_mesic["Kod_Zakazky"] = df_mesic[col_zakazka].astype(str).str.strip().str[5:7]

    platne_kody = ["00", kod_mesice_str]
    mask_platne = df_mesic["Kod_Zakazky"].isin(platne_kody)

    df_platne = df_mesic[mask_platne].copy()
    df_vyrazene = df_mesic[~mask_platne].copy()

    # Kontrola chybějících nákladů
    mask_chybi = (
        df_platne[col_naklady].isna() | 
        (df_platne[col_naklady].astype(str).str.strip() == "") | 
        (df_platne[col_naklady] == 0)
    )
    df_chybi = df_platne[mask_chybi].copy()

    st.divider()
    st.subheader("📊 Výsledky testovací kontroly")

    m1, m2, m3 = st.columns(3)
    m1.metric("Načtené zakázky pro měsíc", len(df_mesic))
    m2.metric("Vyřazené zakázky (kód jiného měsíce)", len(df_vyrazene))
    m3.metric("⚠️ Nalezeno zakázek CHYBĚJÍCÍ NÁKLAD", len(df_chybi))

    if not df_vyrazene.empty:
        with st.expander("ℹ️ Správně vyřazená zakázka (nepatří do tohoto měsíce)"):
            st.dataframe(df_vyrazene[[col_zakazka, col_datum, col_referent, "Kod_Zakazky"]], use_container_width=True)

    if not df_chybi.empty:
        st.warning("Zakázky vyžadující doplnění nákladů:")
        st.dataframe(df_chybi[[col_zakazka, col_datum, col_referent, col_naklady]], use_container_width=True)

        st.divider()
        st.subheader("📧 Testovací odeslání e-mailového výkazu")
        
        muj_osobni_email = st.text_input("Zadejte VÁŠ e-mail pro doručení testovacího výkazu:", "vás.email@firma.cz")
        
        if st.button("🚀 Odeslat testovací výkaz na MŮJ e-mail"):
            # Vytvoření souhrnu pro všechny referenty
            html_obsah = f"<h2>[TEST] Přehled chybějících nákladů za {vybrany_mesic}/{vybrany_rok}</h2>"
            html_obsah += "<p>Tato zpráva je testovací a simuluje e-maily, které dostanou jednotliví referenti.</p>"
            
            skupiny = df_chybi.groupby(col_referent)
            for referent, zakazky in skupiny:
                html_obsah += f"<h3>Referent: {referent}</h3>"
                html_obsah += zakazky[[col_zakazka, col_datum, "Název org."]].to_html(index=False)
                html_obsah += "<hr>"

            uspech, sprava = odeslat_testovaci_email(muj_osobni_email, f"Testovací výkaz chybějících nákladů {vybrany_mesic}/{vybrany_rok}", html_obsah)
            if uspech:
                st.success(sprava)
            else:
                st.error(sprava)
