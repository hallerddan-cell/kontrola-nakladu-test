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

    # Zjištění sloupce pro službu
    col_sluzba = "Služba" if "Služba" in df_ocistene.columns else None

    # Označení stavu nákladů s výjimkou pro HELPDESK
    mask_chybi_zaklad = (
        df_ocistene[col_naklady].isna() | 
        (df_ocistene[col_naklady].astype(str).str.strip() == "") | 
        (df_ocistene[col_naklady].astype(str).str.strip() == "0") |
        (df_ocistene[col_naklady] == 0)
    )

    if col_sluzba:
        is_helpdesk = df_ocistene[col_sluzba].astype(str).str.strip().str.upper() == "HELPDESK"
        mask_chybi = mask_chybi_zaklad & (~is_helpdesk)
    else:
        mask_chybi = mask_chybi_zaklad
        
    df_ocistene["Stav nákladů"] = np.where(mask_chybi, "❌ Chybí náklad", "✅ V pořádku")
    df_ocistene["Datum vytvoření"] = df_ocistene["Datum_dt"].dt.strftime('%d.%m.%Y')

    # Odstranění pomocných sloupců před exportem, aby data byla čistá
    df_export_ocistene = df_ocistene.drop(columns=["Datum_dt", "Kod_Zakazky"], errors="ignore")
    df_chybi = df_export_ocistene[mask_chybi].copy()
