# combine_to_passworded_excel.py
from pathlib import Path
import os
import pandas as pd

# --- helpers ---
def downloads_dir() -> Path:
    d = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Downloads"
    if not d.exists():
        for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
            base = os.environ.get(var)
            if base and (Path(base) / "Downloads").exists():
                return Path(base) / "Downloads"
    d.mkdir(parents=True, exist_ok=True)
    return d

# --- source files ---
base_ac = Path.home() / r"OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare"
base_tr = Path.home() / r"OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat"

adt = pd.read_csv(base_ac / "ADT with Zone and Reference.csv")
adt_active = adt[adt["Status"].astype(str).str.lower() == "active"][
    ["ClientFullName", "ClientID", "ServiceCode", "ServiceCodeDept", "Status", "Zone", "ServiceStartDate"]
]

rpt = pd.read_csv(base_tr / "rpt_Census.csv", dtype={"ID": str})
rpt = rpt[["ClientName", "ID", "Program", "Staff", "AdmitDate", "DischargeDate"]]
rpt_active = rpt[rpt["DischargeDate"].isna()]

# --- write to Excel in Downloads ---
dl = downloads_dir()
tmp_xlsx = dl / "Active_Combined_tmp.xlsx"
final_xlsx = dl / "Active_Combined.xlsx"

with pd.ExcelWriter(tmp_xlsx, engine="xlsxwriter") as writer:
    adt_active.to_excel(writer, sheet_name="AlayaCare", index=False)
    rpt_active.to_excel(writer, sheet_name="Treat", index=False)

# --- password-protect workbook (requires Excel + pywin32) ---
try:
    import win32com.client as win32  # pip install pywin32
    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(tmp_xlsx))
    # 51 = xlsx; Password sets the *open* password
    wb.SaveAs(str(final_xlsx), FileFormat=51, Password="clients")
    wb.Close(SaveChanges=False)
    excel.Quit()
    os.remove(tmp_xlsx)
    print(f"✅ Saved password-protected workbook: {final_xlsx}")
except Exception as e:
    # Fallback: keep unprotected file so you still get an output
    os.replace(tmp_xlsx, final_xlsx)
    print("⚠️ Could not password-protect (need Excel + pywin32). Saved without password.")
    print(f"Error: {e}\nFile: {final_xlsx}")
