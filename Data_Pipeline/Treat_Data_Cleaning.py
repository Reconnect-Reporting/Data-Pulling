# Treat_Data_Cleaning.py
from __future__ import annotations
import csv
import os
import re
from pathlib import Path
from typing import Iterable, Optional, List, Tuple

import pandas as pd  # NEW: for Excel handling

# -------- Folders --------
def raw_dir() -> Path:
    return Path.home() / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Raw Data" / "Treat"

def clean_dir() -> Path:
    return Path.home() / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Clean Data" / "Treat"

# -------- Test IDs to remove --------
TEST_IDS = {"1","10314","12675","7941","7942","9383","9384","11633","8239"}

# ============================================================
# =============== CSV HELPERS (existing logic) ===============
# ============================================================
def _open_with_encodings(p: Path, encodings=("utf-8-sig","utf-8","cp1252","latin-1")) -> Tuple[str, str]:
    for enc in encodings:
        try:
            with p.open("r", encoding=enc, errors="strict") as f:
                return f.read(8192), enc
        except Exception:
            pass
    with p.open("r", encoding="utf-8", errors="ignore") as f:
        return f.read(8192), "utf-8"

def _detect_delimiter(sample: str) -> str:
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", "\t", ";", "|"])
        return dialect.delimiter
    except Exception:
        pass
    cands = [",", "\t", ";", "|"]
    counts = {d: sample.count(d) for d in cands}
    return max(counts, key=count(s) if (count := counts.get(d, 0)) else 0) if any(counts.values()) else ","

def _find_header_row_by_colA(p: Path, enc: str, delim: str, needles: Iterable[str]) -> Optional[int]:
    needles_l = [n.lower() for n in needles]
    with p.open("r", encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=delim)
        for i, row in enumerate(reader):
            if not row:
                continue
            a = (row[0] or "").strip().lower()
            if any(n in a for n in needles_l):
                return i
    return None

def _read_after_header(p: Path, enc: str, delim: str, header_row_idx: int) -> Tuple[List[str], List[List[str]]]:
    with p.open("r", encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=delim)
        header = []
        for _ in range(header_row_idx+1):
            header = next(reader, [])
        rows = [row for row in reader]
    return [c if c is not None else "" for c in header], rows

def _write_csv_atomic(dest: Path, delim: str, header: List[str], rows: List[List[str]]):
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + "._tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as fout:
        w = csv.writer(fout, delimiter=delim, lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
        w.writerow(header)
        for r in rows:
            w.writerow(r)
    os.replace(tmp, dest)

def _case_index(header: List[str]) -> dict:
    m = {}
    for i, h in enumerate(header):
        key = (h or "").strip().lower()
        if key and key not in m:
            m[key] = i
    return m

def _rename_header_ci(header: List[str], mapping_ci: dict[str,str] | None) -> List[str]:
    if not mapping_ci:
        return header[:]
    out = []
    for h in header:
        repl = None
        for old, new in mapping_ci.items():
            if (h or "").strip().lower() == old.lower():
                repl = new
                break
        out.append(repl if repl is not None else (h or ""))
    return out

def _drop_cols_ci(header: List[str], drop_list: Iterable[str]) -> Tuple[List[str], List[int]]:
    """Return (new_header, keep_idxs) where keep_idxs refer to the ORIGINAL header."""
    drops = {d.lower() for d in drop_list}
    keep_idxs = [i for i, h in enumerate(header) if (h or "").strip().lower() not in drops]
    new_header = [header[i] for i in keep_idxs]
    return new_header, keep_idxs

def _project_rows(rows: List[List[str]], keep_idxs: List[int]) -> List[List[str]]:
    """Project rows only (use when you already computed the new header)."""
    out = []
    for r in rows:
        out.append([r[i] if i < len(r) else "" for i in keep_idxs])
    return out

# -------- Filters / transforms (existing) --------
def _filter_out_test_ids(header: List[str], rows: List[List[str]], id_col_name: str) -> Tuple[List[str], List[List[str]]]:
    idx_map = _case_index(header)
    i = idx_map.get(id_col_name.lower())
    if i is None:
        return header, rows
    kept = []
    for r in rows:
        val = (r[i] if i < len(r) else "").strip()
        if val not in TEST_IDS:
            kept.append(r)
    return header, kept

def _adtcensuscan_transform(header: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    header = _rename_header_ci(header, {"textbox2": "Staff", "patient": "ClientName"})
    new_header, keep_idxs = _drop_cols_ci(header, ["NewAdmissionstoOrg","DischargedfromOrg","ClientsActiveinOrg","Lead_Health_Home"])
    rows = _project_rows(rows, keep_idxs)
    header = new_header
    header, rows = _filter_out_test_ids(header, rows, "ID")
    return header, rows

def _rpt_census_transform(header: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    drop_cols = [
        "AddtionalAddressType","AdditionalAddressType",
        "HealthNumberOrMedicaid",
        "ProvinceLabel","AddProvinceLabel",
        "SSN","Site",
        "LastClaimDate","LastClaimD","LastClaimDt"
    ]
    new_header, keep_idxs = _drop_cols_ci(header, drop_cols)
    rows = _project_rows(rows, keep_idxs)
    header = new_header
    header, rows = _filter_out_test_ids(header, rows, "ID")
    return header, rows

def _rpt_progstafffinance_transform(header: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    idx_map = _case_index(header)
    i_client = idx_map.get("textbox55")
    if i_client is not None:
        new_rows = []
        for r in rows:
            txt = (r[i_client] if i_client < len(r) else "").strip()
            m = re.search(r"\((\d+)\)\s*$", txt)
            cid = m.group(1) if m else ""
            name = re.sub(r"\s*\(\d+\)\s*$", "", txt).strip()
            if i_client >= len(r):
                r = r + [""]*(i_client - len(r) + 1)
            r[i_client] = name
            r.append(cid)
            new_rows.append(r)
        rows = new_rows
        header = header[:] + ["__EXTRACTED_ID__"]

    header = _rename_header_ci(header, {
        "textbox55": "ClientName",
        "__EXTRACTED_ID__": "ID",
        "textbox57": "Interaction Start Date",
        "textbox101": "Start Time",
        "textbox102": "End Time",
        "textbox37": "Staff",
        "textbox96": "Task",
        "textbox60": "Program",
        "textbox32": "Visit Type",
        "textbox50": "Note Date",
    })

    new_header, keep_idxs = _drop_cols_ci(header, ["textbox92", "ClaimID"])
    rows = _project_rows(rows, keep_idxs)
    header = new_header

    header, rows = _filter_out_test_ids(header, rows, "ID")
    return header, rows

def _rpt_mis_stats_transform(header: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    # Final mapping
    name_map = {
        "textbox55": "Program",
        "textbox24": "FCC",
        "textbox21": "MIS Code",
        "textbox18": "Description",
        "textbox15": "Apr",
        "textbox9":  "May",
        "textbox12": "June",
        "textbox30": "Q1",
        "textbox2":  "July",
        "textbox3":  "Aug",
        "textbox4":  "Sep",
        "textbox5":  "Q2",
        "textbox6":  "Oct",
        "textbox7":  "Nov",
        "textbox8":  "Dec",
        "textbox10": "Q3",
        "textbox11": "Jan",
        "textbox13": "Feb",
        "textbox16": "Mar",
        "textbox17": "Q4",   # second April source preserved distinctly
        "textbox19": "Total",      # corrected per your note
    }

    # Apply rename (case-insensitive)
    header = _rename_header_ci(header, name_map)

    # Keep only these columns, in this order
    target = [
        "Program", "FCC", "MIS Code", "Description",
        "Apr", "May", "June", "Q1",
        "July", "Aug", "Sep", "Q2",
        "Oct", "Nov", "Dec", "Q3",
        "Jan", "Feb", "Mar", "Apr_2",
        "Q4",
    ]

    hmap = _case_index(header)  # lowercased header -> index
    keep_idxs = [hmap[t.lower()] for t in target if t.lower() in hmap]
    final_header = [t for t in target if t.lower() in hmap]

    rows = _project_rows(rows, keep_idxs)
    header = final_header
    return header, rows


# ============================================================
# =============== EXCEL HELPERS (new logic) =================
# ============================================================
_EXCEL_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")

def _find_latest_excel_by_prefix(folder: Path, prefix: str) -> Optional[Path]:
    """Return newest Excel file whose *stem* starts with prefix (case-insensitive)."""
    prefix_l = prefix.lower()
    cands = [p for p in folder.iterdir()
             if p.is_file() and p.suffix.lower() in _EXCEL_EXTS and p.stem.lower().startswith(prefix_l)]
    if not cands:
        return None
    return max(cands, key=lambda p: p.stat().st_mtime)

def _read_excel_any(p: Path) -> pd.DataFrame:
    """Read Excel into DataFrame with no header first; handle engines as needed."""
    try:
        return pd.read_excel(p, header=None, dtype=str)
    except Exception:
        ext = p.suffix.lower()
        if ext in (".xlsx", ".xlsm"):
            return pd.read_excel(p, header=None, dtype=str, engine="openpyxl")
        if ext == ".xls":
            return pd.read_excel(p, header=None, dtype=str, engine="xlrd")
        if ext == ".xlsb":
            return pd.read_excel(p, header=None, dtype=str, engine="pyxlsb")
        raise

def _detect_header_row_excel(df: pd.DataFrame, needles: Iterable[str]) -> int:
    """Find first row where any cell == any needle (case-insensitive); fallback to first non-empty."""
    target = {str(n).strip().lower() for n in needles}
    for i in range(len(df)):
        row_vals = df.iloc[i].astype(str).str.strip().str.lower()
        if any(v in target for v in row_vals):
            return i
    for i in range(len(df)):
        if df.iloc[i].notna().any():
            return i
    return 0

def _materialize_from_header_excel(df_raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    cols = df_raw.iloc[header_row].astype(str).str.strip().tolist()
    body = df_raw.iloc[header_row+1:].copy()
    body.columns = cols
    body = body.dropna(how="all")
    body.columns = [c.strip() if isinstance(c, str) else c for c in body.columns]
    return body

def _rename_mrn_to_id_df(df: pd.DataFrame) -> pd.DataFrame:
    ren = {c: "ID" for c in df.columns if isinstance(c, str) and c.strip().lower() == "mrn"}
    return df.rename(columns=ren) if ren else df

def _filldown_all_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.ffill()

def _export_df_to_csv(df: pd.DataFrame, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(dest, index=False, encoding="utf-8-sig", lineterminator="\n")

def process_excel(prefix: str, *, dest_name: str, needles=("MRN",), filldown: bool = False) -> Optional[Path]:
    """Find newest Excel file by prefix, do headerize + (optional) filldown + MRN→ID, write CSV to Clean."""
    src_folder = raw_dir()
    dst_folder = clean_dir()
    p = _find_latest_excel_by_prefix(src_folder, prefix)
    if not p:
        print(f"[SKIP] {prefix}*.xlsx: not found in {src_folder}")
        return None

    try:
        df_raw = _read_excel_any(p)
        hdr = _detect_header_row_excel(df_raw, needles=needles)
        df = _materialize_from_header_excel(df_raw, hdr)
        if filldown:
            df = _filldown_all_df(df)
        df = _rename_mrn_to_id_df(df)
        out = dst_folder / dest_name
        _export_df_to_csv(df, out)
        print(f"[OK] {p.name} -> {out.name} (Excel rows={len(df_raw)}, header row={hdr})")
        return out
    except Exception as e:
        print(f"[WARN] {p.name}: failed to process Excel -> CSV ({e})")
        return None

# ============================================================
# =============== GENERIC CSV PROCESSOR ======================
# ============================================================
def _find_src_by_stem(folder: Path, stem: str) -> Optional[Path]:
    target = stem.lower()
    for p in folder.iterdir():
        if p.is_file() and p.stem.lower() == target:
            return p
    p = folder / f"{stem}.csv"
    return p if p.exists() else None

def process_file(stem: str,
                 needles: List[str],
                 *,
                 dest_name_override: str | None = None,
                 per_file_transform = None):
    src_folder = raw_dir()
    dst_folder = clean_dir()
    src = _find_src_by_stem(src_folder, stem)
    if not src:
        print(f"[SKIP] {stem}: not found in {src_folder}")
        return

    sample, enc = _open_with_encodings(src)
    delim = _detect_delimiter(sample)
    hdr_idx = _find_header_row_by_colA(src, enc, delim, needles)
    if hdr_idx is None:
        print(f"[WARN] {src.name}: no header row found in column A for {needles}")
        return

    header, rows = _read_after_header(src, enc, delim, hdr_idx)

    print(f"-> Processing {src.name}")
    if per_file_transform:
        header, rows = per_file_transform(header, rows)

    dest = dst_folder / (dest_name_override if dest_name_override else src.name)
    _write_csv_atomic(dest, delim, header, rows)
    print(f"[OK] {src.name} -> {dest.name} (row {hdr_idx}, sep '{delim}', enc {enc})")

# ============================================================
# ====================== ORCHESTRATOR ========================
# ============================================================
def run():
    raw = raw_dir()
    clean = clean_dir()
    if not raw.exists():
        raise RuntimeError(f"Raw folder not found: {raw}")
    clean.mkdir(parents=True, exist_ok=True)

    # ---- Existing CSV flows ----
    process_file("ADTCensusCAN", ["NewAdmissionstoOrg"], per_file_transform=_adtcensuscan_transform)
    process_file("RECONNECTWORKProofingByClinician", ["textbox121"])
    process_file("rpt_Census", ["AddtionalAddressType","AdditionalAddressType"], per_file_transform=_rpt_census_transform)
    process_file("rpt_MIS_Stats", ["textbox1"], per_file_transform=_rpt_mis_stats_transform)
    process_file("rpt_ProgStaffFinance", ["textbox55"], per_file_transform=_rpt_progstafffinance_transform)

    # rpt_Referrals_DT_waitlisted (save as fixed name)
    for stem in ["rpt_Referrals_DT (1)", "rpt_Referrals_DT_waitlisted"]:
        if _find_src_by_stem(raw, stem):
            process_file(stem, ["Textbox1"], dest_name_override="rpt_Referrals_DT_waitlisted.csv")
            break
    else:
        print("[INFO] No 'rpt_Referrals_DT (1)' or 'rpt_Referrals_DT_waitlisted' found in Raw.")

    process_file("rpt_Referrals_DT", ["Textbox1"])

    # ---- NEW: Excel flows (convert + clean → CSV in Clean folder) ----
    # Assessment: filldown all columns, MRN→ID
    process_excel("Assessment", dest_name="Assessment.csv", needles=("MRN",), filldown=True)

    # Demographics: MRN→ID only
    process_excel("Demographics", dest_name="Demographics.csv", needles=("MRN",), filldown=False)

    # External_Documents_Report: filldown all columns, MRN→ID
    process_excel("External_Documents_Report", dest_name="External_Documents_Report.csv", needles=("MRN",), filldown=True)

# Optional alias
def main():
    run()

if __name__ == "__main__":
    run()
