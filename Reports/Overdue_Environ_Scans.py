
import pandas as pd
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

USER_HOME = Path.home()

DATA_PATH = (
    USER_HOME
    / "OneDrive - Reconnect Community Health Services (1)"
    / "data"
    / "Clean Data"
)

OUTPUT_FILE = (
    USER_HOME / "Downloads" / "Overdue_Environmental_Scans.xlsx"
)


# ============================================================
# ALAYACARE - ENVIRONMENTAL SCANS (SFS)
# ============================================================

def process_sfs():

    print("Processing AlayaCare Environmental Scans...")

    folder = DATA_PATH / "AlayaCare"

    # Load files
    adt = pd.read_csv(folder / "ADT with Zone and Reference.csv")
    forms = pd.read_csv(folder / "Form Report.csv")
    cm = pd.read_csv(folder / "CM Supervisors.csv")

    # Standardize client IDs
    adt["ClientID"] = adt["ClientID"].astype("string").str.strip()
    forms["ClientID"] = forms["ClientID"].astype("string").str.strip()
    cm["AlayaCare Client ID"] = (
        cm["AlayaCare Client ID"].astype("string").str.strip()
    )

    # Convert dates
    adt["ServiceStartDate"] = pd.to_datetime(
        adt["ServiceStartDate"], errors="coerce"
    )

    adt["Status (Effective) Date"] = pd.to_datetime(
        adt["Status (Effective) Date"], errors="coerce"
    )

    forms["SubmissionDate"] = pd.to_datetime(
        forms["SubmissionDate"], errors="coerce"
    )

    # Filter Environmental Scan forms
    forms = forms.loc[
        forms["FormName"] == "Environmental Scan",
        ["ClientID", "SubmissionDate", "FormName"]
    ].drop_duplicates()

    # Filter active clients
    excluded_departments = [
        "Health and Wellness",
        "Meals on Wheels",
        "Transportation",
        "York W",
        ""
    ]

    adt = adt.loc[
        (adt["Status"] == "active") &
        (~adt["ServiceCodeDept"].isin(excluded_departments)),
        [
            "ClientID",
            "ServiceCode",
            "ServiceCodeDept",
            "ServiceStartDate",
            "Status (Effective) Date",
            "Status",
            "Zone"
        ]
    ].copy()

    # Retain one record per client
    adt = adt.drop_duplicates(
        subset=["ClientID"],
        keep="first"
    )

    # Determine completion status
    adt["EnvironScan Completed?"] = (
        adt["ClientID"]
        .isin(forms["ClientID"])
        .map({
            True: "Completed",
            False: "Overdue"
        })
    )

    # Add supervisor information
    cm = cm[
        ["AlayaCare Client ID", "Employee Full Name"]
    ].drop_duplicates(subset=["AlayaCare Client ID"])

    adt = adt.merge(
        cm,
        how="left",
        left_on="ClientID",
        right_on="AlayaCare Client ID",
        validate="many_to_one"
    )

    # Keep overdue clients with valid service start dates
    overdue_sfs = adt.loc[
        (adt["EnvironScan Completed?"] == "Overdue") &
        adt["ServiceStartDate"].notna()
    ].copy()

    print(f"AlayaCare overdue scans: {len(overdue_sfs):,}")

    return overdue_sfs


# ============================================================
# TREAT - ENVIRONMENTAL SCANS (MHA)
# ============================================================

def process_mha():

    print("Processing TREAT Environmental Scans...")

    folder = DATA_PATH / "Treat"

    # Load files
    adt = pd.read_csv(folder / "ADTCensusCAN.csv")
    external_docs = pd.read_csv(
        folder / "External_Documents_Report.csv"
    )
    prog_notes = pd.read_csv(
        folder / "rpt_ProgStaffFinance.csv"
    )

    # Standardize IDs
    for df in [adt, external_docs, prog_notes]:
        df["ID"] = pd.to_numeric(
            df["ID"], errors="coerce"
        ).astype("Int64")

    # Convert dates
    adt["AdmitDate"] = pd.to_datetime(
        adt["AdmitDate"], errors="coerce"
    )

    external_docs["DOCUMENTDATE"] = pd.to_datetime(
        external_docs["DOCUMENTDATE"], errors="coerce"
    )

    prog_notes["Interaction Start Date"] = pd.to_datetime(
        prog_notes["Interaction Start Date"], errors="coerce"
    )

    # Excluded programs
    excluded_programs = [
        "ACT Social Group",
        "Addiction Focused Psycho-Education group",
        "Art Therapy",
        "ASH/COPA Skill Group",
        "Behavioural Therapy Individual Sessions",
        "CBT Adult Group",
        "CDBT Skills Group",
        "Central Intake",
        "Clinical Nutrition Services",
        "COPA SMART Recovery Group",
        "DBT 101 Skills Group",
        "DBT Groups",
        "DBT Intake",
        "Dog assisted Therapy Group",
        "Drama Therapy",
        "FAMEkids Family Caregive Support Services Group",
        "FAMEkids Family Caregive Support Services",
        "Family and Caregiver Support Services",
        "Family and Caregiver Support Services Group",
        "FOCUS",
        "Mindfulness Group",
        "Music Therapy",
        "Short-Term Crisis Beds",
        "SOTO Social Group",
        "South Toronto Link",
        "Streets 2 Home",
        "WeTo Social Group"
    ]

    # Filter active clients
    adt = adt.loc[
        adt["DischargeDate"].isna() &
        (~adt["Program"].isin(excluded_programs)),
        [
            "ClientName",
            "ID",
            "Staff",
            "Program",
            "AdmitDate",
            "DischargeDate"
        ]
    ].copy()

    # Identify Environmental Scan external documents
    external_docs = external_docs.loc[
        external_docs["FILENAME"].str.contains(
            "Environment",
            case=True,
            na=False
        ),
        ["ID", "DOCUMENTDATE", "FILENAME"]
    ]

    # Identify Environmental Scan progress notes
    prog_notes = prog_notes.loc[
        prog_notes["Task"] == "Environmental Scan"
    ]

    # Retain one record per client
    adt = adt.drop_duplicates(
        subset=["ID"],
        keep="first"
    )

    # Combine completed scans from both sources
    completed_ids = pd.concat(
        [
            external_docs["ID"],
            prog_notes["ID"]
        ],
        ignore_index=True
    ).dropna().drop_duplicates()

    # Determine completion status
    adt["EnvironScan Completed?"] = (
        adt["ID"]
        .isin(completed_ids)
        .map({
            True: "Completed",
            False: "Overdue"
        })
    )

    # Keep overdue clients
    overdue_mha = adt.loc[
        (adt["EnvironScan Completed?"] == "Overdue") &
        adt["Program"].notna() &
        (adt["Program"] != "")
    ].copy()

    print(f"TREAT overdue scans: {len(overdue_mha):,}")

    return overdue_mha


# ============================================================
# RUN REPORT
# ============================================================

def run():

    # Process both systems
    overdue_mha = process_mha()
    overdue_sfs = process_sfs()

    # Ensure Downloads folder exists
    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    # Export to one Excel workbook
    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        date_format="yyyy-mm-dd"
    ) as writer:

        overdue_mha.to_excel(
            writer,
            sheet_name="Overdue MHA",
            index=False
        )

        overdue_sfs.to_excel(
            writer,
            sheet_name="Overdue SFS",
            index=False
        )

        # Format worksheet headers and column widths
        for sheet_name, df in [
            ("Overdue MHA", overdue_mha),
            ("Overdue SFS", overdue_sfs)
        ]:

            worksheet = writer.sheets[sheet_name]

            worksheet.freeze_panes(1, 0)

            worksheet.autofilter(
                0, 0,
                len(df),
                len(df.columns) - 1
            )

            for col_num, column in enumerate(df.columns):

                width = min(
                    max(len(str(column)) + 3, 16),
                    35
                )

                worksheet.set_column(
                    col_num,
                    col_num,
                    width
                )

    # Final summary
    print("\n" + "=" * 50)
    print("ENVIRONMENTAL SCAN REPORT COMPLETE")
    print("=" * 50)
    print(f"TREAT overdue:     {len(overdue_mha):,}")
    print(f"AlayaCare overdue: {len(overdue_sfs):,}")
    print(f"Total overdue:     {len(overdue_mha) + len(overdue_sfs):,}")
    print(f"\nSaved to:\n{OUTPUT_FILE}")


# ============================================================
# MAIN
# ============================================================

def main():
    run()


if __name__ == "__main__":
    main()
