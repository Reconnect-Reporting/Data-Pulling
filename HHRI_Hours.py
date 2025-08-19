import sys, os, tempfile
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


# ---- Email config ----
RECIPIENTS = "recreporting@reconnect.on.ca"
SUBJECT = f"HHRI Hours (encrypted) - {datetime.now():%Y-%m-%d}"
BODY = (
    "Hi team,\n\n"
    "Please find the HHRI Hours report attached. Could you please confirm the hours. Thanks!\n\n"
    "Best,\nAutomated Sender"
)
OPEN_PASSWORD = "clients"  # password to open the workbook
ATTACHMENT_DISPLAY_NAME = "HHRI Hours.xlsx"

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

def send_email_with_attachment(to_addrs: str, subject: str, body: str, attachment_path: Path, display_name: str):
    """Attach file and rename it in the email to `display_name`."""
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addrs
    mail.Subject = subject
    mail.Body = body
    att = mail.Attachments.Add(str(attachment_path))
    # PR_DISPLAY_NAME (0x3001, PT_UNICODE 0x001F)
    att.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3001001F", display_name)
    # mail.Display()  # uncomment to preview before sending
    mail.Send()

def main():
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Input not found: {INPUT_PATH}")

    # 1) Clean data from CSV
    df = clean_hhri_csv(INPUT_PATH)

    # 2) Use temp workspace; no permanent save
    tmpdir = Path(tempfile.mkdtemp(prefix="hhri_"))
    try:
        csv_clean = tmpdir / "hhri_clean.csv"
        xlsx_enc = tmpdir / "hhri_encrypted.xlsx"

        # Save cleaned dataframe back to CSV (simple + reliable for Excel)
        df.to_csv(csv_clean, index=False, encoding="utf-8-sig")

        # 3) Excel converts CSV -> encrypted XLSX
        excel_csv_to_encrypted_xlsx(csv_clean, xlsx_enc, OPEN_PASSWORD)

        # 4) Email it as "HHRI Hours.xlsx"
        send_email_with_attachment(RECIPIENTS, SUBJECT, BODY, xlsx_enc, ATTACHMENT_DISPLAY_NAME)
        print(f"[INFO] Email sent to {RECIPIENTS}")
    finally:
        # 5) Cleanup everything
        for p in [*tmpdir.glob("*")]:
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
