from __future__ import annotations
import getpass
from pathlib import Path
import pandas as pd

def run() -> pd.DataFrame:
    # --- Paths ---
    USER_NAME = getpass.getuser()
    DATA_DIR  = Path(fr"C:\Users\{USER_NAME}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat")

    # --- Pick newest rpt_MIS_Stats CSV ---
    mis_files = sorted(DATA_DIR.glob("rpt_MIS_Stats*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not mis_files:
        raise FileNotFoundError(f"No 'rpt_MIS_Stats*.csv' found in {DATA_DIR}")

    mis_path = mis_files[0]
    print(f"[MIS] Using: {mis_path.name}")

    # --- Read (preserve original columns; no transforms) ---
    mis_df = pd.read_csv(mis_path, encoding="utf-8-sig", keep_default_na=False)

    mis_df = mis_df[mis_df['Category'].isin([
        'Service Provider Interactions',
        'Service Provider Group Interactions',
        'Attendance Days Face to Face',
        'Attendance Days Non Face to Face',
        'Resident Days',
        'Group Participants',
        'Group Session',
    ])]

    # 1) Drop unneeded columns
    mis_df2 = mis_df.drop(columns=["FCC", "MIS Code", "Description"], errors="ignore").copy()

    # 2) Identify month columns in FY order (Apr..Mar) that actually exist in your data
    FY_MONTHS = ["Apr","May","June","July","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
    month_cols = [c for c in FY_MONTHS if c in mis_df2.columns]
    print(f"[MIS] Month columns detected: {month_cols}")

    # 3) Coerce months to numeric
    mis_df2[month_cols] = mis_df2[month_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    # 4) Sum per Category × Program (months only)
    combined = (
        mis_df2.groupby(["Category", "Program"], as_index=False)[month_cols]
               .sum()
    )

    # 5) Build "Total until end of last month" (Apr → last completed month)
    first_of_this_month = pd.Timestamp.today().replace(day=1)
    last_month_num = (first_of_this_month - pd.DateOffset(days=1)).month
    num_to_lbl = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"June",7:"July",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
    last_lbl = num_to_lbl[last_month_num]

    # months to include = Apr .. last_lbl (only those present in the dataframe)
    cutoff_idx = FY_MONTHS.index(last_lbl)
    months_to_sum = [m for m in FY_MONTHS[:cutoff_idx+1] if m in combined.columns]
    combined["Total until end of last month"] = combined[months_to_sum].sum(axis=1)
    print(f"[MIS] Summing FY-to-date through: {last_lbl}  -> {months_to_sum}")

    # Keep only needed columns
    keep = ["Category", "Program", "Total until end of last month"]
    tbl = combined[keep].copy()

    # Pivot: rows = Program, columns = Category, values = "Total until end of last month"
    pivot = (
        tbl.pivot_table(
            index="Program",
            columns="Category",
            values="Total until end of last month",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    # Optional: sort category columns and tidy header
    category_cols = sorted([c for c in pivot.columns if c != "Program"])
    pivot = pivot[["Program"] + category_cols]
    pivot.columns.name = None

    # --- Pick the newest rpt_ProgStaffFinance.csv ---
    cands = sorted(DATA_DIR.glob("rpt_ProgStaffFinance.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not cands:
        raise FileNotFoundError("rpt_ProgStaffFinance.csv found in Clean Data\\Treat")

    p = cands[0]
    print(f"[ProgStaffFinance] Using: {p.name}")
    df_progress = pd.read_csv(p, encoding="utf-8-sig", dtype={"ID": "string"})

    # Parse the date column (MM/DD/YYYY in your sample), with a safe fallback
    d = pd.to_datetime(df_progress["Note Date"], format="%m/%d/%Y", errors="coerce")
    d = d.fillna(pd.to_datetime(df_progress["Note Date"], errors="coerce"))

    # Fiscal window: Apr 1 of active FY → last day of last month
    today = pd.Timestamp.today().normalize()
    apr1_this_year = pd.Timestamp(today.year, 4, 1)
    fy_start = apr1_this_year if today >= apr1_this_year else pd.Timestamp(today.year - 1, 4, 1)
    first_of_this_month = today.replace(day=1)
    eom_last_month = first_of_this_month - pd.Timedelta(days=1)

    # Filter
    mask = (d >= fy_start) & (d <= eom_last_month)
    df_progress_fy = df_progress.loc[mask].copy()

    # (optional) keep the parsed datetime back in the filtered frame
    df_progress_fy["Note Date"] = d.loc[df_progress_fy.index]

    print(f"[ProgStaffFinance] Fiscal window: {fy_start.date()} → {eom_last_month.date()} | kept {len(df_progress_fy)} rows")

    # Unique individuals served per Program (FY window already applied)
    df_prog_counts = (
        df_progress_fy
          .assign(ID_num=pd.to_numeric(df_progress_fy["ID"], errors="coerce"))
          .dropna(subset=["ID_num"])
          .groupby("Program", dropna=False)["ID_num"]
          .nunique()
          .reset_index(name="Individuals Served")
          .sort_values("Program", kind="stable")
    )

    # Visits per Program (dedupe by client+date+visit type+program)
    tmp_date = pd.to_datetime(df_progress_fy["Interaction Start Date"], errors="coerce").dt.normalize()
    df_progress_fy = df_progress_fy.assign(_date_only=tmp_date)

    dedupe_cols = ["ClientName", "_date_only", "Visit Type", "Program"]
    df_progress_fy = (
        df_progress_fy
        .sort_values(dedupe_cols)
        .drop_duplicates(subset=dedupe_cols, keep="first")
        .drop(columns="_date_only")
        .reset_index(drop=True)
    )

    df_prog_visits = (
        df_progress_fy
          .groupby("Program", dropna=False)
          .size()
          .reset_index(name="Visits")
          .sort_values("Program", kind="stable")
    )

    # Merge to final
    main_table = (
        pd.merge(pivot, df_prog_counts, how="left", on="Program")
          .merge(df_prog_visits, how="left", on="Program")
    )

    # Nicer ints
    for c in ["Individuals Served", "Visits"]:
        if c in main_table.columns:
            main_table[c] = main_table[c].fillna(0).astype("int64")

    print(f"[OUTPUT] Rows={len(main_table)}, Cols={len(main_table.columns)}")
    # Show a peek in the log (your main captures stdout; see patch below)
    with pd.option_context("display.max_rows", 20, "display.max_columns", 20, "display.width", 140):
        print(main_table.head(10).to_string(index=False))


    main_table = main_table.rename(columns={"Attendance Days Face to Face": "Attendance Days F2F", "Attendance Days Non Face to Face": "Attendance Days NF2F","Service Provider Group Interactions":"SPGI"})

    return main_table


if __name__ == "__main__":
    df = run()