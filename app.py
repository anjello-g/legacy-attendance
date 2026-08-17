for f in hc_files:
    # Use sheet name "Master"
    df = pd.read_excel(f, sheet_name="Master")

    # Rename columns to internal expected names
    rename_map = {
        "Client": "Project",
        "Subprocess": "Sub-Process",
        "Separation Date": "Date of Separation",
    }
    df = df.rename(columns=rename_map)

    if "ECN" not in df.columns or "DOJ Knack" not in df.columns:
        st.error(f"Headcount **{f.name}** missing ECN or DOJ Knack"); st.stop()

    export_dt = extract_headcount_date(f.name, df)
    dt_key = export_dt.strftime("%Y-%m-%d") if export_dt else f.name

    df["ECN"] = pd.to_numeric(df["ECN"], errors="coerce").dropna().astype(int)
    for _, row in df.iterrows():
        doj = parse_doj_knack(row.get("DOJ Knack"))
        if doj:
            doj_lookup[int(row["ECN"])] = doj
    hc_cache[dt_key] = df
    hc_dates.append(export_dt)
