# ocan_overdue.py
import os
from datetime import date, datetime
from pathlib import Path
import pandas as pd


class OCANOverdue:
    def __init__(self, base_path: str | None = None):
        """
        base_path: folder that contains rpt_Census.csv and Assessment.csv.
        If None, will try common OneDrive paths.
        """
        self.today = date.today()
        self.user = os.getlogin()

        if base_path is None:
            # Try with and without the "(1)" suffix
            candidates = [
                rf"C:\Users\{self.user}\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat",
                rf"C:\Users\{self.user}\OneDrive - Reconnect Community Health Services\data\Clean Data\Treat",
            ]
            for p in candidates:
                if Path(p).exists():
                    base_path = p
                    break
            if base_path is None:
                raise FileNotFoundError("Could not find Treat folder; please pass base_path explicitly.")
        self.base_path = base_path

        # Downloads folder
        self.downloads = str(Path.home() / "Downloads")

    def run(self) -> pd.DataFrame:
        # --- Read CSVs
        census = pd.read_csv(os.path.join(self.base_path, "rpt_Census.csv"), dtype={'ID': str})
        assessment = pd.read_csv(os.path.join(self.base_path, "Assessment.csv"), dtype={'ID': str})

        # --- Active only + keep needed cols, drop FAME
        census_active = census[census['DischargeDate'].isna()].copy()
        census_active = census_active[['ClientName', 'ID', 'Program', 'Staff', 'AdmitDate']].copy()
        census_active = census_active[~census_active['Program'].str.contains('FAME', case=False, na=False)].copy()

        # --- Merge multiple staff into one cell (normalized casing)
        census_active = census_active.groupby(
            ["ClientName", "ID", "Program", "AdmitDate"], as_index=False
        ).agg({"Staff": lambda x: ", ".join(sorted(set(x)))})

        census_active["Staff"] = census_active["Staff"].apply(
            lambda names: ", ".join(" ".join(w.capitalize() for w in name.split()) for name in names.split(", "))
        )

        # --- OCAN-only assessments; keep ASSESSMENTDATE as datetime for arithmetic
        assessment_OCAN = assessment[assessment['TOOLNAME'] == 'OCAN'].copy()
        assessment_OCAN['ASSESSMENTDATE'] = pd.to_datetime(assessment_OCAN['ASSESSMENTDATE'], errors='coerce')

        # --- Merge on ID
        census_with_assessment = pd.merge(census_active, assessment_OCAN, how='left', on='ID')

        # ---- Keep one row per (ClientName, ID, Program, AdmitDate, Staff):
        keys = ["ClientName", "ID", "Program", "AdmitDate", "Staff"]
        nonempty = census_with_assessment.dropna(subset=["ASSESSMENTDATE"]).copy()
        empty    = census_with_assessment[census_with_assessment["ASSESSMENTDATE"].isna()].copy()

        # Keep the MOST RECENT assessment per group
        latest = (
            nonempty.sort_values("ASSESSMENTDATE", ascending=False)
                    .groupby(keys, as_index=False)
                    .head(1)
        )

        # Keep empty rows only for groups that had no non-empty counterpart
        empty_keep = empty.merge(latest[keys], on=keys, how="left", indicator=True)
        empty_keep = empty_keep[empty_keep["_merge"] == "left_only"].drop(columns="_merge")

        # Combine back
        census_with_assessment = pd.concat([latest, empty_keep], ignore_index=True)

        # OCAN status: if no date -> "OCAN not Found"; else days since last assessment
        today_ts = pd.Timestamp(self.today)
        census_with_assessment["Days Since Last OCAN"] = census_with_assessment["ASSESSMENTDATE"].apply(
            lambda d: "OCAN not Found" if pd.isna(d) else (today_ts - d.normalize()).days
        )

        # Optional tidy column order
        cols_first = ["ClientName", "ID", "Program", "AdmitDate", "Staff",
                      "FULLNAME", "TOOLNAME", "ASSESSMENTTYPE", "ASSESSMENTDATE", "Days Since Last OCAN"]
        existing = [c for c in cols_first if c in census_with_assessment.columns]
        census_with_assessment = census_with_assessment[existing + [
            c for c in census_with_assessment.columns if c not in existing
        ]]


        col = "Days Since Last OCAN"

        # Convert numeric values where possible, non-numeric (like "OCAN not Found") become NaN
        census_with_assessment[col + "_num"] = pd.to_numeric(census_with_assessment[col], errors="coerce")

        # Keep rows where it's "OCAN not Found" OR numeric > 183
        census_with_assessment = census_with_assessment[(census_with_assessment[col] == "OCAN not Found") | (census_with_assessment[col + "_num"] > 183)]

        # Drop helper column
        census_with_assessment = census_with_assessment.drop(columns=[col + "_num"])

        census_with_assessment = census_with_assessment.drop(columns = ['FULLNAME'])
        # Save to Downloads
        ts = datetime.now().strftime("%Y%m%d")
        out_path = os.path.join(self.downloads, f"Outstanding OCANs_{ts}.csv")
        #census_with_assessment.to_csv(out_path, index=False)

        # return df for further use in the app
        return census_with_assessment

def run():
    job = OCANOverdue()
    df  = job.run()

    downloads = Path.home() / "Downloads"   # points to C:\Users\<you>\Downloads
    out_file = downloads / "Overdue_OCAN List.csv"

    df.to_csv(out_file, index=False, encoding="utf-8")
    print(f"[Overdue_OCAN_List] Saved: {out_file}")

# ----- Minimal example usage -----
if __name__ == "__main__":
    job = OCANOverdue()  # or OCANOverdue(base_path=r"...\Treat")
    df_result = job.run()
    print("Saved to Downloads and returned DataFrame with shape:", df_result.shape)
