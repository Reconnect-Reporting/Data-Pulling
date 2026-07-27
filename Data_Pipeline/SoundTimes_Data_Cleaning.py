# SoundTimes_Data_Cleaning.py
from __future__ import annotations

import csv
import os
import re
from pathlib import Path
from typing import Iterable, Optional, List, Tuple


# ===================== FOLDERS =====================
def raw_dir() -> Path:
    return (
        Path.home()
        / "OneDrive - Reconnect Community Health Services (1)"
        / "data"
        / "Raw Data"
        / "Sound Times"
    )


def clean_dir() -> Path:
    return (
        Path.home()
        / "OneDrive - Reconnect Community Health Services (1)"
        / "data"
        / "Clean Data"
        / "Sound Times"
    )


# ===================== TEST IDS TO REMOVE =====================
TEST_IDS = {
    "1", "10314", "12675", "7941", "7942",
    "9383", "9384", "11633", "8239", "7945"
}


# ===================== CSV HELPERS =====================
def _open_with_encodings(
    p: Path,
    encodings=("utf-8-sig", "utf-8", "cp1252", "latin-1")
) -> Tuple[str, str]:
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

    candidates = [",", "\t", ";", "|"]
    counts = {d: sample.count(d) for d in candidates}

    if any(counts.values()):
        return max(counts, key=counts.get)

    return ","


def _find_header_row_by_colA(
    p: Path,
    enc: str,
    delim: str,
    needles: Iterable[str]
) -> Optional[int]:
    """
    Find header row when Column A contains the target text.
    """
    needles_l = [n.lower() for n in needles]

    with p.open("r", encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=delim)

        for i, row in enumerate(reader):
            if not row:
                continue

            col_a = (row[0] or "").strip().lower()

            if any(n in col_a for n in needles_l):
                return i

    return None


def _find_header_row_any_col(
    p: Path,
    enc: str,
    delim: str,
    needles: Iterable[str]
) -> Optional[int]:
    """
    Find header row if any cell contains one of the target text values.
    Used for STimesWORKProofingByClinician_Stimes because its header contains Clinician.
    """
    needles_l = [n.lower() for n in needles]

    with p.open("r", encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=delim)

        for i, row in enumerate(reader):
            if not row:
                continue

            row_values = [
                (cell or "").strip().lower()
                for cell in row
            ]

            if any(any(n in cell for n in needles_l) for cell in row_values):
                return i

    return None


def _read_after_header(
    p: Path,
    enc: str,
    delim: str,
    header_row_idx: int
) -> Tuple[List[str], List[List[str]]]:
    with p.open("r", encoding=enc, newline="") as f:
        reader = csv.reader(f, delimiter=delim)

        header = []

        for _ in range(header_row_idx + 1):
            header = next(reader, [])

        rows = [row for row in reader]

    header = [c if c is not None else "" for c in header]

    return header, rows


def _write_csv_atomic(
    dest: Path,
    delim: str,
    header: List[str],
    rows: List[List[str]]
):
    """
    Write CSV and replace existing file with the same name.
    """
    dest.parent.mkdir(parents=True, exist_ok=True)

    tmp = dest.with_suffix(dest.suffix + "._tmp")

    with tmp.open("w", encoding="utf-8-sig", newline="") as fout:
        writer = csv.writer(
            fout,
            delimiter=delim,
            lineterminator="\n",
            quoting=csv.QUOTE_MINIMAL
        )

        writer.writerow(header)

        for row in rows:
            writer.writerow(row)

    # This replaces the file if the same name already exists.
    os.replace(tmp, dest)


def _case_index(header: List[str]) -> dict:
    mapping = {}

    for i, h in enumerate(header):
        key = (h or "").strip().lower()

        if key and key not in mapping:
            mapping[key] = i

    return mapping


def _rename_header_ci(
    header: List[str],
    mapping_ci: dict[str, str] | None
) -> List[str]:
    if not mapping_ci:
        return header[:]

    output = []

    for h in header:
        replacement = None

        for old, new in mapping_ci.items():
            if (h or "").strip().lower() == old.lower():
                replacement = new
                break

        output.append(replacement if replacement is not None else (h or ""))

    return output


def _drop_cols_ci(
    header: List[str],
    drop_list: Iterable[str]
) -> Tuple[List[str], List[int]]:
    drops = {d.lower() for d in drop_list}

    keep_idxs = [
        i for i, h in enumerate(header)
        if (h or "").strip().lower() not in drops
    ]

    new_header = [header[i] for i in keep_idxs]

    return new_header, keep_idxs


def _project_rows(
    rows: List[List[str]],
    keep_idxs: List[int]
) -> List[List[str]]:
    output = []

    for row in rows:
        output.append([
            row[i] if i < len(row) else ""
            for i in keep_idxs
        ])

    return output


def _filter_out_test_ids(
    header: List[str],
    rows: List[List[str]],
    id_col_name: str
) -> Tuple[List[str], List[List[str]]]:
    idx_map = _case_index(header)
    id_idx = idx_map.get(id_col_name.lower())

    if id_idx is None:
        return header, rows

    kept_rows = []

    for row in rows:
        value = (row[id_idx] if id_idx < len(row) else "").strip()

        if value not in TEST_IDS:
            kept_rows.append(row)

    return header, kept_rows


# ===================== TRANSFORMS =====================
def _rpt_progstafffinance_transform(
    header: List[str],
    rows: List[List[str]]
) -> Tuple[List[str], List[List[str]]]:
    idx_map = _case_index(header)
    i_client = idx_map.get("textbox55")

    if i_client is not None:
        new_rows = []

        for row in rows:
            text = (row[i_client] if i_client < len(row) else "").strip()

            match = re.search(r"\((\d+)\)\s*$", text)

            client_id = match.group(1) if match else ""
            client_name = re.sub(r"\s*\(\d+\)\s*$", "", text).strip()

            if i_client >= len(row):
                row = row + [""] * (i_client - len(row) + 1)

            row[i_client] = client_name
            row.append(client_id)

            new_rows.append(row)

        rows = new_rows
        header = header[:] + ["__EXTRACTED_ID__"]

    header = _rename_header_ci(
        header,
        {
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
        }
    )

    new_header, keep_idxs = _drop_cols_ci(
        header,
        ["textbox92", "ClaimID"]
    )

    rows = _project_rows(rows, keep_idxs)
    header = new_header

    header, rows = _filter_out_test_ids(header, rows, "ID")

    return header, rows


def _rpt_mis_stats_transform(
    header: List[str],
    rows: List[List[str]]
) -> Tuple[List[str], List[List[str]]]:
    name_map = {
        "textbox1": "Category",
        "textbox55": "Program",
        "textbox24": "FCC",
        "textbox21": "MIS Code",
        "textbox18": "Description",
        "textbox15": "Apr",
        "textbox9": "May",
        "textbox12": "June",
        "textbox30": "Q1",
        "textbox2": "July",
        "textbox3": "Aug",
        "textbox4": "Sep",
        "textbox5": "Q2",
        "textbox6": "Oct",
        "textbox7": "Nov",
        "textbox8": "Dec",
        "textbox10": "Q3",
        "textbox11": "Jan",
        "textbox13": "Feb",
        "textbox16": "Mar",
        "textbox17": "Q4",
        "textbox19": "Total",
    }

    header = _rename_header_ci(header, name_map)

    target = [
        "Category",
        "Program",
        "FCC",
        "MIS Code",
        "Description",
        "Apr",
        "May",
        "June",
        "Q1",
        "July",
        "Aug",
        "Sep",
        "Q2",
        "Oct",
        "Nov",
        "Dec",
        "Q3",
        "Jan",
        "Feb",
        "Mar",
        "Apr_2",
        "Q4",
        "Total",
    ]

    hmap = _case_index(header)

    keep_idxs = [
        hmap[col.lower()]
        for col in target
        if col.lower() in hmap
    ]

    final_header = [
        col
        for col in target
        if col.lower() in hmap
    ]

    rows = _project_rows(rows, keep_idxs)
    header = final_header

    return header, rows


def _no_rename_transform(
    header: List[str],
    rows: List[List[str]]
) -> Tuple[List[str], List[List[str]]]:
    """
    For files where we only need to remove report rows above the header
    and keep the original columns exactly as-is.
    """
    return header, rows


# ===================== FILE FINDING / PROCESSOR =====================
def _find_src_by_stem(folder: Path, stem: str) -> Optional[Path]:
    """
    Finds exact stem first.
    If not found, finds newest file whose stem starts with the stem.

    Supports files like:
      rpt_MIS_Stats_Stimes.csv
      rpt_MIS_Stats_Stimes_1.csv
      rpt_Referrals_DT_Stimes.csv
      rpt_Referrals_DT_Stimes1.csv
      ADTCensusCAN_Stimes.csv
      ADTCensusCAN_Stimes_1.csv
    """
    target = stem.lower()

    exact_candidates = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".csv"
        and p.stem.lower() == target
    ]

    if exact_candidates:
        return max(exact_candidates, key=lambda p: p.stat().st_mtime)

    prefix_candidates = [
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".csv"
        and p.stem.lower().startswith(target)
    ]

    if prefix_candidates:
        return max(prefix_candidates, key=lambda p: p.stat().st_mtime)

    return None


def process_file(
    stem: str,
    needles: List[str],
    *,
    dest_name_override: str | None = None,
    per_file_transform=None,
    search_any_column: bool = False
):
    src_folder = raw_dir()
    dst_folder = clean_dir()

    src = _find_src_by_stem(src_folder, stem)

    if not src:
        print(f"[SKIP] {stem}: not found in {src_folder}")
        return None

    sample, enc = _open_with_encodings(src)
    delim = _detect_delimiter(sample)

    if search_any_column:
        header_idx = _find_header_row_any_col(src, enc, delim, needles)
    else:
        header_idx = _find_header_row_by_colA(src, enc, delim, needles)

    if header_idx is None:
        print(f"[WARN] {src.name}: no header row found for {needles}")
        return None

    header, rows = _read_after_header(src, enc, delim, header_idx)

    print(f"-> Processing {src.name}")

    if per_file_transform:
        header, rows = per_file_transform(header, rows)

    dest_name = dest_name_override if dest_name_override else src.name
    dest = dst_folder / dest_name

    _write_csv_atomic(dest, delim, header, rows)

    print(
        f"[OK] {src.name} -> {dest.name} "
        f"(header row {header_idx}, sep '{delim}', enc {enc})"
    )

    return dest


# ===================== ORCHESTRATOR =====================
def run():
    raw = raw_dir()
    clean = clean_dir()

    if not raw.exists():
        raise RuntimeError(f"Raw folder not found: {raw}")

    clean.mkdir(parents=True, exist_ok=True)

    print(f"Raw folder:   {raw}")
    print(f"Clean folder: {clean}")

    # 1. Same cleaning as Treat rpt_MIS_Stats
    process_file(
        "rpt_MIS_Stats_Stimes",
        ["textbox1"],
        dest_name_override="rpt_MIS_Stats_Stimes.csv",
        per_file_transform=_rpt_mis_stats_transform
    )

    # 2. Same cleaning as Treat rpt_ProgStaffFinance
    process_file(
        "rpt_ProgStaffFinance_Stimes",
        ["textbox55"],
        dest_name_override="rpt_ProgStaffFinance_Stimes.csv",
        per_file_transform=_rpt_progstafffinance_transform
    )

    # 3. STimes WorkProofing:
    # Detect header by "Clinician" and do not rename columns.
    process_file(
        "STimesWORKProofingByClinician_Stimes",
        ["Clinician"],
        dest_name_override="WORKProofingByClinician_Stimes.csv",
        per_file_transform=_no_rename_transform,
        search_any_column=True
    )

    # 4. Referral Summary - New Referrals:
    # Detect Textbox1, no column renaming.
    process_file(
        "rpt_Referrals_DT_Stimes",
        ["Textbox1"],
        dest_name_override="rpt_Referrals_DT_Stimes.csv",
        per_file_transform=_no_rename_transform
    )

    # 5. Referral Summary - Waitlisted:
    # Source: rpt_Referrals_DT_Stimes_1
    # Output: rpt_Referrals_DT_waitlisted_Stimes.csv
    process_file(
        "rpt_Referrals_DT_Stimes_1",
        ["Textbox1"],
        dest_name_override="rpt_Referrals_DT_waitlisted_Stimes.csv",
        per_file_transform=_no_rename_transform
    )

    # 6. ADT Census CAN:
    # Source: ADTCensusCAN_Stimes.csv in Raw Data/Sound Times
    # Detect "NewAdmissionstoOrg" and use that row as the header.
    # No other cleaning / renaming needed.
    process_file(
        "ADTCensusCAN_Stimes",
        ["NewAdmissionstoOrg"],
        dest_name_override="ADTCensusCAN_Stimes.csv",
        per_file_transform=_no_rename_transform
    )

    print("Sound Times Treat cleaning completed.")


def main():
    run()


if __name__ == "__main__":
    main()