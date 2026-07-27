import sys, os, tempfile, shutil
from pathlib import Path
import pandas as pd
import win32com.client as win32

# ---- Fixed CSV input path ----
ONE_DRIVE = Path.home() / "OneDrive - Reconnect Community Health Services (1)"

INPUT_PATH = ONE_DRIVE / "data" / "Clean Data" / "AlayaCare" / "HHRI(VHA).csv"

if not INPUT_PATH.exists():
    raise FileNotFoundError(f"Could not find file at: {INPUT_PATH}")

print("Resolved INPUT_PATH:", INPUT_PATH)

# ---- Output (local Downloads) ----
DOWNLOADS = Path.home() / "Downloads"
ATTACHMENT_DISPLAY_NAME = "HHRI Hours.xlsx"
OPEN_PASSWORD = "clients"  # password to open the workbook


def clean_hhri_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # Drop the extra index column if it exists
    # Example: blank column or "Unnamed: 0"
    first_col = df.columns[0]
    if first_col == "" or first_col.lower().startswith("unnamed"):
        df = df.drop(columns=[first_col])

    # Drop Completed column if present
    completed_cols = [
        col for col in df.columns
        if col.strip().lower() in ["completed (yes / no)", "completed"]
    ]

    if completed_cols:
        df = df.drop(columns=completed_cols)

    # Rename columns by actual source column names, not by position
    rename_map = {
        "Start Date": "Date of Visit",
        "Scheduled Duration": "Duration",
        "Name": "Program",
        "Full Name": "Worker",
        "Code": "Bill Code",
        "Start Time": "Time of Visit",
    }

    df = df.rename(columns=rename_map)

    # Minimal validation
    required_cols = [
        "Last Name",
        "First Name",
        "Date of Visit",
        "Time of Visit",
        "Duration",
        "Program",
        "Worker",
        "Bill Code",
        "Zone",
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        raise ValueError(
            f"Missing required columns: {missing_cols}. Got columns: {list(df.columns)}"
        )

    # Format Date of Visit
    df["Date of Visit"] = pd.to_datetime(
        df["Date of Visit"],
        errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    # Format Time of Visit from original Start Time
    # Example: 2026-05-08 8:45 -> 08:45 AM
    df["Time of Visit"] = pd.to_datetime(
        df["Time of Visit"],
        errors="coerce"
    ).dt.strftime("%I:%M %p")

    # Clean Worker column
    df["Worker"] = df["Worker"].replace({"null null": "No Worker"})

    # Final column order
    df = df[
        [
            "Last Name",
            "First Name",
            "Date of Visit",
            "Time of Visit",
            "Duration",
            "Program",
            "Worker",
            "Bill Code",
            "Zone",
        ]
    ]

    return df


def excel_csv_to_encrypted_xlsx(csv_path: Path, xlsx_path: Path, password: str):
    """Open CSV in Excel and SaveAs password-protected .xlsx."""
    excel = None
    wb = None

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(csv_path))

        # FileFormat=51 means .xlsx
        wb.SaveAs(str(xlsx_path), FileFormat=51, Password=password)

        wb.Close(SaveChanges=False)
        wb = None

    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


def main():
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Input not found: {INPUT_PATH}")

    # 1) Clean data from CSV
    df = clean_hhri_csv(INPUT_PATH)

    # 2) Use temp workspace to let Excel do the encryption
    tmpdir = Path(tempfile.mkdtemp(prefix="hhri_"))

    try:
        csv_clean = tmpdir / "hhri_clean.csv"
        xlsx_enc_tmp = tmpdir / "hhri_encrypted.xlsx"

        # Save cleaned dataframe back to CSV
        df.to_csv(csv_clean, index=False, encoding="utf-8-sig")

        # 3) Excel converts CSV to encrypted XLSX
        excel_csv_to_encrypted_xlsx(csv_clean, xlsx_enc_tmp, OPEN_PASSWORD)

        # 4) Copy encrypted file to Downloads
        DOWNLOADS.mkdir(parents=True, exist_ok=True)
        out_path = DOWNLOADS / ATTACHMENT_DISPLAY_NAME
        shutil.copy2(xlsx_enc_tmp, out_path)

        print(f"[INFO] Saved to: {out_path}")

    finally:
        # 5) Cleanup temp workspace
        for p in tmpdir.glob("*"):
            try:
                os.remove(p)
            except Exception:
                pass

        try:
            os.rmdir(tmpdir)
        except Exception:
            pass

        print("[INFO] Cleaned up temporary files.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)