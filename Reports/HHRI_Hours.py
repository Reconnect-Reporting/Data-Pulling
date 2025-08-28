import sys, os, tempfile, shutil
from pathlib import Path
from datetime import datetime
import pandas as pd
import win32com.client as win32

# ---- Fixed CSV input path ----
ONE_DRIVE = Path(
    os.environ.get("OneDriveCommercial")
    or os.environ.get("OneDrive")
    or (Path.home() / "OneDrive - Reconnect Community Health Services (1)")
)
INPUT_PATH = ONE_DRIVE / "data" / "Clean Data" / "AlayaCare" / "HHRI(VHA).csv"

# ---- Output (local Downloads) ----
DOWNLOADS = Path.home() / "Downloads"
ATTACHMENT_DISPLAY_NAME = "HHRI Hours.xlsx"
OPEN_PASSWORD = "clients"  # password to open the workbook

def clean_hhri_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    # Drop columns 1 and 10 (1-based) if present
    drop_idxs = []
    if df.shape[1] >= 1:  drop_idxs.append(0)
    if df.shape[1] >= 10: drop_idxs.append(9)
    if drop_idxs:
        df = df.drop(df.columns[drop_idxs], axis=1)

    # Standardize column names (up to available columns)
    target_cols = ["Last Name","First Name","Date","Duration","Program","Worker","Bill Code","Zone"]
    df = df.rename(columns={old: new for old, new in zip(df.columns.tolist(), target_cols)})

    # Minimal validation/cleanup
    if "Date" not in df.columns or "Worker" not in df.columns:
        raise ValueError(f"Missing required columns. Got: {list(df.columns)}")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Worker"] = df["Worker"].replace({"null null": "No Worker"})
    return df

def excel_csv_to_encrypted_xlsx(csv_path: Path, xlsx_path: Path, password: str):
    """Open CSV in Excel and SaveAs password-protected .xlsx (password to open)."""
    excel = None
    wb = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(csv_path))
        # FileFormat=51 -> .xlsx; Password -> open password
        wb.SaveAs(str(xlsx_path), FileFormat=51, Password=password)
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()

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

        # Save cleaned dataframe back to CSV (simple + reliable for Excel)
        df.to_csv(csv_clean, index=False, encoding="utf-8-sig")

        # 3) Excel converts CSV -> encrypted XLSX
        excel_csv_to_encrypted_xlsx(csv_clean, xlsx_enc_tmp, OPEN_PASSWORD)

        # 4) Copy the encrypted file to local Downloads with friendly name
        DOWNLOADS.mkdir(parents=True, exist_ok=True)
        out_path = DOWNLOADS / ATTACHMENT_DISPLAY_NAME
        shutil.copy2(xlsx_enc_tmp, out_path)
        print(f"[INFO] Saved to: {out_path}")

    finally:
        # 5) Cleanup temp workspace
        for p in tmpdir.glob("*"):
            try: os.remove(p)
            except Exception: pass
        try: os.rmdir(tmpdir)
        except Exception: pass
        print("[INFO] Cleaned up temporary files.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
