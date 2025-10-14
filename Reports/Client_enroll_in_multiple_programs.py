import os
from pathlib import Path
import pandas as pd


def main(username: str | None = None,
         save_dir: Path | None = None) -> Path:
    """
    Build 'updated_Clients Enroll in Multiple programs.xlsx' and return its path.
    - username: override Windows username (defaults to os.getlogin()).
    - save_dir: override output folder (defaults to AlayaCare clean-data folder).
    """
    user = username or os.getlogin()

    # ---- Paths ----
    referrals_dir = Path(
        fr"C:\Users\{user}\Reconnect Community Health Services\Health Services - Documents\Decision Support\Decision Support Data\Reports\Data\TREAT\Historical\Referrals"
    )
    treat_clean_dir = Path(
        fr"C:\Users\{user}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat"
    )
    alaya_dir = Path(
        fr"C:\Users\{user}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare"
    )

    out_dir = save_dir or alaya_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "updated_Clients Enroll in Multiple programs.xlsx"

    # ---- Referrals (historical) ----
    referrals_path = referrals_dir / "rpt_Referrals_total.csv"
    referrals = pd.read_csv(referrals_path, dtype={"MRN": str}, encoding="utf-8-sig")
    if "ReferralDate" in referrals.columns:
        referrals = referrals.drop(columns=["ReferralDate"])

    # ---- Treat: rpt_Census (active only) ----
    census_path = treat_clean_dir / "rpt_Census.csv"
    rpt_Census = pd.read_csv(census_path, dtype={"ID": str}, encoding="utf-8-sig")
    rpt_Census = rpt_Census[["ClientName", "ID", "Program", "DischargeDate"]]
    rpt_Census = rpt_Census[rpt_Census["Program"] != "Central Intake"]

    test_ids = ["1", "10314", "12675", "7941", "7942", "9383", "9384", "11633", "8239"]
    rpt_Census = rpt_Census[~rpt_Census["ID"].isin(test_ids)]

    rpt_Census_active = rpt_Census[rpt_Census["DischargeDate"].isna()].copy()
    rpt_Census_active = rpt_Census_active.rename(columns={"ID": "Treat ID"}).drop(columns=["DischargeDate"])

    rpt_Census_active = pd.merge(
        rpt_Census_active,
        referrals,
        how="left",
        left_on=["Treat ID", "Program"],
        right_on=["MRN", "Program"],
        copy=False,
    )

    # Fill + fold referral source into Program
    rpt_Census_active["ReferralSource"] = rpt_Census_active.get("ReferralSource", pd.Series(index=rpt_Census_active.index))
    rpt_Census_active["ReferralSource"].fillna("Unknown", inplace=True)
    rpt_Census_active["Program"] = rpt_Census_active["Program"] + " (" + rpt_Census_active["ReferralSource"] + ")"
    rpt_Census_active = rpt_Census_active.drop(columns=["ReferralSource", "MRN"], errors="ignore")

    # ---- AlayaCare: ADT ----
    adt_path = alaya_dir / "ADT with Zone and Reference.csv"
    adt = pd.read_csv(adt_path, encoding="utf-8-sig")

    adt = adt[["ClientFullName", "ClientID", "ServiceCodeDept", "Referral Source", "Status", "Zone"]].copy()
    adt["Referral Source"] = adt["Referral Source"].fillna("Unknown")
    adt = adt[adt["Status"] == "active"].copy()

    adt = adt.rename(
        columns={
            "ClientID": "AlayaCare ID",
            "ServiceCodeDept": "Program",
            "ClientFullName": "ClientName",
        }
    ).drop(columns=["Status"])

    adt["Program"] = adt["Program"] + " (" + adt["Referral Source"] + ")"
    adt = adt.drop(columns=["Referral Source"])

    # ---- Combine & keep multi-program clients ----
    client_list = pd.concat([adt, rpt_Census_active], ignore_index=True, copy=False)
    client_list["Count"] = client_list.groupby("ClientName")["Program"].transform("nunique")
    client_list = client_list[client_list["Count"] >= 2].copy()
    client_list = client_list.drop_duplicates()

    # ---- Group & explode Programs into Program + Referral Source ----
    client_list_grouped = (
        client_list.groupby("ClientName", as_index=False)
        .agg(
            Programs=("Program", lambda x: ", ".join(x.astype(str))),
            AlayaCare_ID=("AlayaCare ID", "first"),
            Treat_ID=("Treat ID", "first"),
            Zone=("Zone", "first"),
            Count=("Count", "first"),
        )
    )

    rows = []
    for _, row in client_list_grouped.iterrows():
        name = row["ClientName"]
        programs = str(row["Programs"]).split(",")
        for p in programs:
            p = p.strip()
            if "(" in p and ")" in p:
                program_name = p[: p.index("(")].strip()
                source = p[p.index("(") + 1 : p.index(")")].strip()
                rows.append(
                    {
                        "ClientName": name,
                        "Program": program_name,
                        "Referral Source": source,
                        "Zone": row["Zone"],
                    }
                )

    updated_client_list_grouped = pd.DataFrame(rows)
    updated_client_list_grouped = updated_client_list_grouped.merge(
        client_list_grouped[["ClientName", "AlayaCare_ID", "Treat_ID"]],
        on="ClientName",
        how="left",
        copy=False,
    )
    updated_client_list_grouped["Zone"] = updated_client_list_grouped["Zone"].fillna("Unknown")

    # ---- Save ----
    updated_client_list_grouped.to_excel(out_path, index=False)
    print(f"Saved: {out_path}")
    return out_path


def run() -> Path:
    """Zero-arg wrapper for job runners."""
    return main()


if __name__ == "__main__":
    main()
