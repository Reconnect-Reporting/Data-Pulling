import os
from datetime import datetime
from pathlib import Path
from typing import Tuple

import pandas as pd


# =========================
#         DEFAULT PATHS
# =========================
user = os.environ.get("UserName") or os.environ.get("USERNAME") or os.getlogin()

DEFAULT_TREAT_DIR = Path(r"C:\Users\=\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\Treat".replace("=", user))
DEFAULT_ALAYA_DIR = Path(r"C:\Users\=\OneDrive - Reconnect Community Health Services (1)\data\Clean Data\AlayaCare".replace("=", user))


# =========================
#        HELPERS
# =========================
def _downloads_dir(fallback: Path) -> Path:
    """Try to resolve the user's Downloads folder (Windows friendly)."""
    candidates = [
        Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Downloads",
        Path.home() / "Downloads",
        Path.home() / "OneDrive" / "Downloads",
    ]
    for p in candidates:
        if p.exists():
            return p
    return fallback  # fallback if none exists

def newest_csv(prefix: str, base: Path) -> Path:
    files = sorted(base.glob(f"{prefix}*.csv"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No '{prefix}*.csv' found in {base}")
    return files[0]

def newest_any(prefix: str, base: Path, exts=(".csv", ".xlsx", ".xls")) -> Path:
    files = []
    for ext in exts:
        files.extend(base.glob(f"{prefix}*{ext}"))
    files = sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No '{prefix}*{exts}' found in {base}")
    return files[0]

def read_any(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "cp1252", "latin1"):
            try:
                return pd.read_csv(path, encoding=enc, keep_default_na=False)
            except UnicodeDecodeError:
                continue
        raise UnicodeDecodeError("Could not decode CSV with utf-8-sig/cp1252/latin1.")
    return pd.read_excel(path, dtype=str)

def _pick_col(df: pd.DataFrame, candidates):
    cols = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().strip()
        if key in cols:
            return cols[key]
    # loose contains match
    for k, orig in cols.items():
        if any(c.lower().strip() in k for c in candidates):
            return orig
    raise KeyError(f"None of {candidates} found in columns: {list(df.columns)}")

def _nonempty_series(s: pd.Series) -> pd.Series:
    return s.notna() & s.astype(str).str.strip().ne("")

def _choose_excel_engine() -> str:
    try:
        import openpyxl  # noqa: F401
        return "openpyxl"
    except Exception:
        try:
            import xlsxwriter  # noqa: F401
            return "xlsxwriter"
        except Exception:
            return "openpyxl"  # let pandas try default and error if missing


# =========================
#      CORE PIPELINE
# =========================
def build_missing_mha(treat_dir: Path) -> pd.DataFrame:
    """Block 1 (TREAT) → missing_df_mha"""
    census_path = newest_csv("rpt_Census", treat_dir)
    print(f"Reading Census: {census_path}")
    census = read_any(census_path)

    discharge_col = _pick_col(census, ["DischargeDate", "Discharge Date"])
    census_active = census[census[discharge_col].astype(str).str.strip().eq("")].copy()

    id_col      = _pick_col(census_active, ["ID", "Client ID", "ClientID"])
    name_col    = _pick_col(census_active, ["ClientName", "Client Name"])
    program_col = _pick_col(census_active, ["Program"])
    staff_col   = _pick_col(census_active, ["Staff", "PrimaryWorker", "Primary Worker", "Staff Name", "Staff Name:"])

    assess_path = newest_csv("Assessment", treat_dir)
    print(f"Reading Assessment: {assess_path}")
    assess = read_any(assess_path)

    tool_col       = _pick_col(assess, ["TOOLNAME", "Tool Name"])
    assess_id_col  = _pick_col(assess, ["ID", "Client ID", "ClientID"])
    assess_filtered = assess[assess[tool_col].astype(str).str.strip().eq("Reconnect Consent Assessment Tool")].copy()

    active_ids  = set(census_active[id_col].astype(str).str.strip())
    assess_ids  = set(assess_filtered[assess_id_col].astype(str).str.strip())
    missing_ids = active_ids - assess_ids

    missing_df_mha = (
        census_active[census_active[id_col].astype(str).str.strip().isin(missing_ids)]
        .loc[:, [name_col, id_col, staff_col, program_col]]
        .rename(columns={
            name_col: "ClientName",
            id_col: "ID",
            staff_col: "Staff",
            program_col: "Program"
        })
        .sort_values(["Program", "Staff", "ClientName"], kind="stable")
        .reset_index(drop=True)
    )
    missing_df_mha = missing_df_mha.drop_duplicates()
    print(f"[MHA] Active clients in Census: {len(active_ids)}")
    print(f"[MHA] With 'Reconnect Consent Assessment Tool' in Assessment: {len(assess_ids)}")
    print(f"[MHA] Missing in Assessment: {missing_df_mha.shape[0]}")
    return missing_df_mha


def build_missing_sfs(alaya_dir: Path) -> pd.DataFrame:
    """Block 2 (ALAYACARE) → missing_df_sfs (no 'Test')"""
    adt_path = newest_any("ADT with Zone and Reference", alaya_dir)
    print(f"Reading ADT: {adt_path}")
    adt = read_any(adt_path)

    status_col      = _pick_col(adt, ["Status"])
    client_col_a    = _pick_col(adt, ["ClientID", "Client ID", "ID"])
    name_col_a      = _pick_col(adt, ["ClientFullName", "Client Name", "ClientName"])
    program_col_a   = _pick_col(adt, ["ServiceCodeDept", "Service Code Department Name"])
    zone_col        = _pick_col(adt, ["Zone"])
    servicecode_col = _pick_col(adt, ["ServiceCode", "Service Code"])

    adt_active = adt[
        _nonempty_series(adt[status_col]) &
        adt[status_col].astype(str).str.strip().str.lower().eq("active")
    ].copy()
    adt_active = adt_active[_nonempty_series(adt_active[client_col_a])].copy()  # drop empty ClientID

    form_path = newest_any("Form Report", alaya_dir)
    print(f"Reading Form Report: {form_path}")
    forms = read_any(form_path)

    formname_col = _pick_col(forms, ["FormName", "Form Name"])
    client_col_f = _pick_col(forms, ["ClientID", "Client ID", "ID"])

    consent_forms = forms[
        _nonempty_series(forms[formname_col]) &
        forms[formname_col].astype(str).str.strip().eq("Consent Form")
    ].copy()
    consent_forms = consent_forms[_nonempty_series(consent_forms[client_col_f])].copy()

    adt_ids     = set(adt_active[client_col_a].astype(str).str.strip())
    consent_ids = set(consent_forms[client_col_f].astype(str).str.strip())
    missing_ids_sfs = adt_ids - consent_ids

    keep_cols = [client_col_a, name_col_a, program_col_a, zone_col, status_col, servicecode_col]
    rename_map = {
        client_col_a: "ClientID",
        name_col_a: "ClientName",
        program_col_a: "ServiceCodeDept",
        zone_col: "Zone",
        status_col: "Status",
        servicecode_col: "ServiceCode",
    }

    missing_df_sfs = (
        adt_active[adt_active[client_col_a].astype(str).str.strip().isin(missing_ids_sfs)]
        .loc[:, keep_cols]
        .rename(columns=rename_map)
        .sort_values(["ServiceCodeDept", "Zone", "ClientName"], kind="stable")
        .reset_index(drop=True)
    )
    missing_df_sfs["Consent Status"] = "Missing"
    missing_df_sfs = missing_df_sfs.drop_duplicates()
    # Remove 'Test' clients
    missing_df_sfs = missing_df_sfs[~missing_df_sfs["ClientName"].astype(str).str.contains("Test", na=False)].copy()

    print(f"[SFS] ADT active clients (non-empty ClientID): {len(adt_ids)}")
    print(f"[SFS] Clients with Consent Form (non-empty ClientID): {len(consent_ids)}")
    print(f"[SFS] Missing (after removing 'Test'): {len(missing_df_sfs)}")
    return missing_df_sfs


def export_to_excel(mha_df: pd.DataFrame, sfs_df: pd.DataFrame, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    engine = _choose_excel_engine()
    with pd.ExcelWriter(out_path, engine=engine) as writer:
        mha_df.to_excel(writer, sheet_name="MHA", index=False)
        sfs_df.to_excel(writer, sheet_name="SFS", index=False)
    print(f"Saved Excel: {out_path}")
    return out_path


# =========================
#           MAIN
# =========================
def main(treat_dir: Path = DEFAULT_TREAT_DIR,
         alaya_dir: Path = DEFAULT_ALAYA_DIR,
         out_path: Path | None = None) -> Tuple[Path, pd.DataFrame, pd.DataFrame]:
    """
    Build MHA and SFS missing-consent reports and export to Excel.
    If out_path is None, save to user's Downloads folder with a timestamped name.
    Returns (out_path, missing_df_mha, missing_df_sfs).
    """
    # Resolve default output location → Downloads
    if out_path is None:
        downloads = _downloads_dir(fallback=treat_dir)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = downloads / f"Consent_Missing_MHA_SFS_{ts}.xlsx"

    missing_df_mha = build_missing_mha(treat_dir)
    missing_df_sfs = build_missing_sfs(alaya_dir)
    out_path = export_to_excel(missing_df_mha, missing_df_sfs, out_path)
    return out_path, missing_df_mha, missing_df_sfs


if __name__ == "__main__":
    main()
