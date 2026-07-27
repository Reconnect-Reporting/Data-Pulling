####################################### TCH + TSH Reports ################################################

"""
TCH + TSH Health & Wellness Report

TCH is monthly and TSH is quarterly.
"""

from pathlib import Path
import calendar
from datetime import datetime

import pandas as pd


def get_last_month():
    """Return the year and month for the previous calendar month."""
    now = datetime.now()

    if now.month == 1:
        return now.year - 1, 12

    return now.year, now.month - 1


def main():
    # =========================
    # Paths
    # =========================
    alayacare_folder = (
        Path.home()
        / "OneDrive - Reconnect Community Health Services (1)"
        / "data"
        / "Clean Data"
        / "AlayaCare"
    )

    visits_path = alayacare_folder / "Visits.csv"
    adt_path = alayacare_folder / "ADT with Zone and Reference.csv"

    output_path = Path.home() / "Downloads" / "Visits_TCH_TSH.xlsx"

    # =========================
    # Load data
    # =========================
    visits_df = pd.read_csv(visits_path)
    adt_df = pd.read_csv(adt_path)

    # =========================
    # Date range - last month
    # =========================
    year, report_month = get_last_month()

    start_date = datetime(year, report_month, 1)
    _, last_day = calendar.monthrange(year, report_month)
    end_date = datetime(year, report_month, last_day)

    # Convert Date to datetime once
    visits_df["Date"] = pd.to_datetime(
        visits_df["Date"],
        errors="coerce"
    )

    # =========================
    # TCH + TSH H&W report
    # =========================
    visits_df = visits_df.rename(
        columns={"AlaycareID": "ClientID"}
    )

    visits_df = visits_df[
        [
            "ClientID",
            "Date",
            "Service Code Department Name",
            "Facility Name",
            "Staff",
        ]
    ].copy()

    visits_df = visits_df[
        (visits_df["Date"] >= start_date)
        & (visits_df["Date"] <= end_date)
        & (visits_df["Service Code Department Name"] == "Health and Wellness")
    ].copy()

    visits_df = pd.merge(
        visits_df,
        adt_df[["ClientID", "ClientAddress"]],
        on="ClientID",
        how="left",
    )

    # =========================
    # TCH
    # =========================
    tch_facilities = [
        "100 High Park",
        "Roselawn",
        "Eagle Manor",
    ]

    TCH = visits_df[
        visits_df["Facility Name"].isin(tch_facilities)
        & (visits_df["Date"] >= pd.Timestamp("2026-04-01"))
        & (visits_df["Date"] <= pd.Timestamp("2026-06-30"))
    ].copy()

    TCH = (
        TCH.sort_values("Date", ascending=False)
        .drop_duplicates(subset=["ClientID", "Date"], keep="first")
    )

    TCH["Date"] = TCH["Date"].dt.date

    # =========================
    # TSH
    # =========================
    tsh_facilities = [
        "Outlook Manor",
        "41 Mabelle Ave",
        "193 Wilson Ave",
        "Silverthorne",
        "98/99 Cavell",
        "Doug Saunders",
        "340 Royal York",
        "Louis Tower",
    ]

    TSH = visits_df[
        visits_df["Facility Name"].isin(tsh_facilities)
    ].copy()

    TSH = (
        TSH.sort_values("Date", ascending=False)
        .drop_duplicates(subset=["ClientID", "Date"], keep="first")
    )

    TSH["Date"] = TSH["Date"].dt.date

    # =========================
    # Export
    # =========================
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        TCH.to_excel(writer, sheet_name="TCH", index=False)
        TSH.to_excel(writer, sheet_name="TSH", index=False)

    print(f"✅ TCH + TSH report saved to: {output_path}")

    return output_path


def run():
    """Entry point used by main.py."""
    return main()


if __name__ == "__main__":
    run()