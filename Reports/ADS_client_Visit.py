import os
import pandas as pd
from pathlib import Path


def run():
    # =========================================================
    # 1. Auto-detect Windows username
    # =========================================================
    username = os.getlogin()

    base_path = Path(
        rf"C:\Users\{username}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare"
    )

    downloads_path = Path.home() / "Downloads"

    adt_path = base_path / "ADT with Zone and Reference.csv"
    visits_path = base_path / "Visits.csv"

    # =========================================================
    # 2. Read files
    # =========================================================
    adt_df = pd.read_csv(adt_path, low_memory=False)
    visits_df = pd.read_csv(visits_path, low_memory=False)

    print("Username:", username)
    print("ADT shape:", adt_df.shape)
    print("Visits shape:", visits_df.shape)

    # =========================================================
    # 3. Clean column names
    # =========================================================
    adt_df.columns = adt_df.columns.str.strip()
    visits_df.columns = visits_df.columns.str.strip()

    # =========================================================
    # 4. Convert Visit Date to datetime
    # =========================================================
    visits_df["Date"] = pd.to_datetime(
        visits_df["Date"],
        errors="coerce"
    )

    visits_df["VisitDate"] = visits_df["Date"].dt.normalize()

    # =========================================================
    # 5. Calculate previous full month
    # =========================================================
    today = pd.Timestamp.today().normalize()

    current_month_start = today.replace(day=1)
    previous_month_start = (
        current_month_start - pd.DateOffset(months=1)
    )
    previous_month_end = (
        current_month_start - pd.Timedelta(days=1)
    )

    print("\nPrevious month:")
    print("Start:", previous_month_start.date())
    print("End:", previous_month_end.date())

    # =========================================================
    # 6. Filter ADT
    # =========================================================
    adt_filtered = adt_df[
        (
            adt_df["ServiceCodeDept"]
            .astype(str)
            .str.strip()
            .str.lower()
            == "adult day program"
        )
        &
        (
            adt_df["Status"]
            .astype(str)
            .str.strip()
            .str.lower()
            == "active"
        )
    ].copy()

    # =========================================================
    # 7. Clean ADT ClientID
    # =========================================================
    adt_filtered["ClientID"] = (
        adt_filtered["ClientID"]
        .astype(str)
        .str.strip()
    )

    adt_filtered = adt_filtered[
        ~adt_filtered["ClientID"]
        .str.lower()
        .isin(["", "nan", "none"])
    ].copy()

    active_ads_ids = set(
        adt_filtered["ClientID"].unique()
    )

    # =========================================================
    # 8. Filter Visits
    # =========================================================
    visits_filtered = visits_df[
        (
            visits_df["Service Code Department Name"]
            .astype(str)
            .str.strip()
            .str.lower()
            == "adult day program"
        )
        &
        (visits_df["VisitDate"] >= previous_month_start)
        &
        (visits_df["VisitDate"] < current_month_start)
    ].copy()

    # =========================================================
    # 9. Clean AlaycareID
    # =========================================================
    visits_filtered["AlaycareID"] = (
        visits_filtered["AlaycareID"]
        .astype(str)
        .str.strip()
    )

    visits_filtered = visits_filtered[
        ~visits_filtered["AlaycareID"]
        .str.lower()
        .isin(["", "nan", "none"])
    ].copy()

    # Keep active ADS clients only
    visits_filtered = visits_filtered[
        visits_filtered["AlaycareID"]
        .isin(active_ads_ids)
    ].copy()

    # =========================================================
    # 10. Randomize rows before duplicate removal
    # =========================================================
    visits_filtered = (
        visits_filtered
        .sample(
            frac=1,
            random_state=42
        )
        .reset_index(drop=True)
    )

    # =========================================================
    # 11. Same client + same date = one visit
    # =========================================================
    visits_unique_daily = (
        visits_filtered
        .drop_duplicates(
            subset=[
                "AlaycareID",
                "VisitDate"
            ],
            keep="first"
        )
        .copy()
    )

    # =========================================================
    # 12. Monthly visit count
    # =========================================================
    visit_counts = (
        visits_unique_daily
        .groupby("AlaycareID")
        .size()
        .rename("Visit Count of the Month")
    )

    visits_unique_daily[
        "Visit Count of the Month"
    ] = (
        visits_unique_daily["AlaycareID"]
        .map(visit_counts)
    )

    # =========================================================
    # 13. Visit Details tab
    # =========================================================
    visit_detail_columns = [
        "AlaycareID",
        "ClientFullName",
        "Date",
        "ServiceCode",
        "Service Code Department Name",
        "Staff",
        "Visit Count of the Month"
    ]

    visit_details = (
        visits_unique_daily[
            visit_detail_columns
        ]
        .drop_duplicates()
        .sort_values(
            by=[
                "AlaycareID",
                "Date"
            ]
        )
        .reset_index(drop=True)
    )

    # =========================================================
    # 14. Client Summary tab
    # =========================================================
    client_summary = (
        visit_details[
            [
                "AlaycareID",
                "ClientFullName",
                "Visit Count of the Month"
            ]
        ]
        .drop_duplicates(
            subset=["AlaycareID"]
        )
        .sort_values(
            by=[
                "ClientFullName",
                "AlaycareID"
            ]
        )
        .reset_index(drop=True)
    )

    # =========================================================
    # 15. Summary
    # =========================================================
    clients_with_visits = (
        client_summary["AlaycareID"]
        .nunique()
    )

    total_unique_visit_days = len(
        visit_details
    )

    print("\n=========================================")
    print("Adult Day Program Visit Summary")
    print("=========================================")

    print(
        f"Month reviewed: "
        f"{previous_month_start.strftime('%B %Y')}"
    )

    print(
        f"Active ADS clients with visits: "
        f"{clients_with_visits}"
    )

    print(
        f"Total unique client visit-days: "
        f"{total_unique_visit_days}"
    )

    # =========================================================
    # 16. Output filename
    # =========================================================
    base_filename = (
        f"ADS_Visits_"
        f"{previous_month_start.strftime('%Y_%m')}.xlsx"
    )

    output_path = (
        downloads_path / base_filename
    )

    counter = 1

    while output_path.exists():
        output_path = downloads_path / (
            f"ADS_Visits_"
            f"{previous_month_start.strftime('%Y_%m')}_"
            f"{counter}.xlsx"
        )

        counter += 1

    # =========================================================
    # 17. Export Excel
    # =========================================================
    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        visit_details.to_excel(
            writer,
            sheet_name="Visit Details",
            index=False
        )

        client_summary.to_excel(
            writer,
            sheet_name="Client Summary",
            index=False
        )

    print("\n=========================================")
    print("Export completed")
    print("=========================================")
    print(output_path)

    # Return useful objects in case your main script needs them
    return {
        "visit_details": visit_details,
        "client_summary": client_summary,
        "output_path": output_path
    }


def main():
    return run()


if __name__ == "__main__":
    main()