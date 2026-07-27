from pathlib import Path
from datetime import datetime
import re

import pandas as pd
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import Pt


def iter_block_items(parent):
    """
    Yield each paragraph and table child within parent, in document order.
    """
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def copy_paragraph(dest_doc, para):
    """
    Copy one paragraph into destination doc.
    Keeps text only, with basic run formatting.
    """
    if not para.text.strip():
        dest_doc.add_paragraph("")
        return

    new_para = dest_doc.add_paragraph()
    try:
        new_para.style = para.style
    except Exception:
        pass

    for run in para.runs:
        new_run = new_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        if run.font.size:
            new_run.font.size = run.font.size
        if run.font.name:
            new_run.font.name = run.font.name


def copy_table(dest_doc, table):
    """
    Copy a table into destination doc.
    Copies text only.
    """
    rows = len(table.rows)
    cols = len(table.columns)
    new_table = dest_doc.add_table(rows=rows, cols=cols)

    try:
        new_table.style = table.style
    except Exception:
        pass

    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            new_table.cell(i, j).text = cell.text


def append_doc_content(dest_doc, source_path):
    """
    Append paragraphs and tables from a source .docx into dest_doc
    without copying document-level section/body XML.
    """
    src = Document(source_path)

    for block in iter_block_items(src):
        if isinstance(block, Paragraph):
            copy_paragraph(dest_doc, block)
        elif isinstance(block, Table):
            copy_table(dest_doc, block)


def get_newest_word_file(folder_path):
    """
    Return newest .docx or .doc file in folder.
    """
    files = list(folder_path.glob("*.docx")) + list(folder_path.glob("*.doc"))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)


def add_big_title(doc, text, size=36):
    """
    Add a large title paragraph.
    """
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    return p


def extract_all_text_from_docx(doc_path):
    """
    Extract text from paragraphs and tables for validation.
    """
    doc = Document(doc_path)
    parts = []

    for para in doc.paragraphs:
        if para.text:
            parts.append(para.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text:
                    parts.append(cell.text)

    return "\n".join(parts)


def validate_doc_dates(doc_path):
    """
    Validate:
    1) date in docx filename == today
    2) date in 'Source file: ...csv' == today
    3) date in 'Generated: YYYY-MM-DD ...' == today
    """
    today_file = datetime.today().strftime("%Y%m%d")
    today_text = datetime.today().strftime("%Y-%m-%d")

    filename_match = re.search(r"(\d{8})", doc_path.stem)
    filename_date = filename_match.group(1) if filename_match else None

    text = extract_all_text_from_docx(doc_path)

    source_match = re.search(
        r"Source file:\s*.*?(\d{8})_\d{6}\.csv",
        text,
        flags=re.IGNORECASE
    )
    source_date = source_match.group(1) if source_match else None

    generated_match = re.search(
        r"Generated:\s*(\d{4}-\d{2}-\d{2})",
        text,
        flags=re.IGNORECASE
    )
    generated_date = generated_match.group(1) if generated_match else None

    details = {
        "filename_date": filename_date,
        "source_date": source_date,
        "generated_date": generated_date,
        "today_file": today_file,
        "today_text": today_text
    }

    is_valid = (
        filename_date == today_file and
        source_date == today_file and
        generated_date == today_text
    )

    return is_valid, details


def find_adt_file():
    """
    Find the ADT with Zone and Reference file from the common AlayaCare clean data folder.
    Supports csv/xlsx/xls.
    """
    home = Path.home()

    candidate_dirs = [
        home / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Clean Data" / "AlayaCare",
        home / "OneDrive - Reconnect Community Health Services" / "data" / "Clean Data" / "AlayaCare",
    ]

    patterns = [
        "ADT with Zone and Reference.csv",
        "ADT with Zone and Reference.xlsx",
        "ADT with Zone and Reference.xls",
        "*ADT with Zone and Reference*.csv",
        "*ADT with Zone and Reference*.xlsx",
        "*ADT with Zone and Reference*.xls",
    ]

    for folder in candidate_dirs:
        if not folder.exists():
            continue
        for pattern in patterns:
            matches = list(folder.glob(pattern))
            if matches:
                return max(matches, key=lambda f: f.stat().st_mtime)

    return None


def read_csv_fallback(file_path):
    """
    Read csv with common encoding fallbacks.
    """
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin1"]
    last_err = None

    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc)
        except Exception as e:
            last_err = e

    raise last_err


def load_adt_zone_map():
    """
    Load ADT with Zone and Reference and return:
    - zone_map: {ClientID: Zone}
    - file_path: source file path

    Keeps:
    - Zone1 ~ Zone9
    - TCU
    - Z-5A
    - Z-5B

    Ignores all other values.
    """
    adt_file = find_adt_file()
    if adt_file is None:
        return {}, None

    suffix = adt_file.suffix.lower()

    if suffix == ".csv":
        df = read_csv_fallback(adt_file)
    elif suffix in [".xlsx", ".xls"]:
        df = pd.read_excel(adt_file, dtype=str)
    else:
        return {}, adt_file

    df.columns = [str(c).strip() for c in df.columns]

    if "ClientID" not in df.columns or "Zone" not in df.columns:
        return {}, adt_file

    df["ClientID"] = df["ClientID"].astype(str).str.strip()
    df["Zone"] = df["Zone"].astype(str).str.strip()

    # Keep Zone1 to Zone9, TCU, Z-5A, Z-5B
    valid_zone_pattern = r"^(Zone[1-9]|TCU|Z-5A|Z-5B)$"

    df = df[
        df["ClientID"].str.startswith("AC", na=False) &
        df["Zone"].notna() &
        (df["Zone"] != "") &
        (df["Zone"].str.lower() != "nan") &
        (df["Zone"].str.match(valid_zone_pattern, na=False))
    ].copy()

    # If duplicated IDs exist, keep the first valid Zone
    df = df.drop_duplicates(subset=["ClientID"], keep="first")

    zone_map = dict(zip(df["ClientID"], df["Zone"]))
    return zone_map, adt_file


def annotate_text_with_zone(text, zone_map):
    """
    Replace AC ClientID with 'ClientID [Zone: X]' if found in zone_map.
    Skip if the text already contains '[Zone:' immediately after the same ID.
    """
    if not text:
        return text

    pattern = re.compile(r"\b(AC\d+)\b")

    def repl(match):
        client_id = match.group(1)
        zone = zone_map.get(client_id)

        if not zone:
            return client_id

        remaining = text[match.end(): match.end() + 25]
        if "[Zone:" in remaining:
            return client_id

        return f"{client_id} [Zone: {zone}]"

    return pattern.sub(repl, text)


def annotate_doc_with_zones(doc, zone_map):
    """
    Go through all paragraphs and tables in the document and append Zone info
    to ClientIDs starting with AC.
    """
    updated_ids = set()

    # Paragraphs
    for para in doc.paragraphs:
        original_text = para.text
        new_text = annotate_text_with_zone(original_text, zone_map)

        if new_text != original_text:
            found_ids = re.findall(r"\b(AC\d+)\b", original_text)
            for cid in found_ids:
                if cid in zone_map:
                    updated_ids.add(cid)

            # Replace paragraph text
            para.clear()
            para.add_run(new_text)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                original_text = cell.text
                new_text = annotate_text_with_zone(original_text, zone_map)

                if new_text != original_text:
                    found_ids = re.findall(r"\b(AC\d+)\b", original_text)
                    for cid in found_ids:
                        if cid in zone_map:
                            updated_ids.add(cid)

                    cell.text = new_text

    return sorted(updated_ids)


def run():
    home = Path.home()

    base_docs = (
        home
        / "OneDrive - Reconnect Community Health Services (1)"
        / "data"
        / "risk_id_programs"
        / "docs"
    )

    downloads = home / "Downloads"

    target_folders = ["act", "als", "copa", "ehts", "mha_cm", "sfscm"]

    display_names = {
        "act": "ACT",
        "als": "ALS",
        "copa": "COPA",
        "ehts": "EHTS",
        "mha_cm": "MHA CM",
        "sfscm": "SFSCM"
    }

    newest_files = {}
    skipped_doc_files = []
    stale_or_invalid_files = []
    checked_messages = []

    for folder_name in target_folders:
        folder_path = base_docs / folder_name

        if not folder_path.exists():
            newest_files[folder_name] = None
            checked_messages.append(f"✗ {display_names[folder_name]} folder not found")
            continue

        newest_file = get_newest_word_file(folder_path)

        if newest_file is None:
            newest_files[folder_name] = None
            checked_messages.append(f"✗ {display_names[folder_name]} no Word file found")
            continue

        newest_files[folder_name] = newest_file

        if newest_file.suffix.lower() == ".doc":
            skipped_doc_files.append(str(newest_file))
            checked_messages.append(f"✗ {display_names[folder_name]} skipped (.doc file)")
            continue

        is_valid, _ = validate_doc_dates(newest_file)

        if is_valid:
            checked_messages.append(f"✓ {display_names[folder_name]} checked")
        else:
            checked_messages.append(f"✗ {display_names[folder_name]} date check failed")
            stale_or_invalid_files.append(str(newest_file))

    master_doc = Document()
    added_any = False

    for idx, folder_name in enumerate(target_folders):
        file_path = newest_files.get(folder_name)

        if file_path is None:
            continue

        if file_path.suffix.lower() == ".doc":
            continue

        is_valid, _ = validate_doc_dates(file_path)
        if not is_valid:
            continue

        add_big_title(master_doc, display_names[folder_name], size=36)
        append_doc_content(master_doc, file_path)

        remaining_valid_after_this = []
        for later_folder in target_folders[idx + 1:]:
            later_file = newest_files.get(later_folder)
            if later_file and later_file.suffix.lower() == ".docx":
                later_valid, _ = validate_doc_dates(later_file)
                if later_valid:
                    remaining_valid_after_this.append(later_folder)

        if remaining_valid_after_this:
            master_doc.add_page_break()

        added_any = True

    # Load ADT zone map and annotate combined doc
    annotated_client_ids = []
    adt_file_used = None

    if added_any:
        zone_map, adt_file_used = load_adt_zone_map()

        if adt_file_used is None:
            checked_messages.append("✗ ADT with Zone and Reference file not found")
        elif not zone_map:
            checked_messages.append(
                f"✗ ADT file found but no usable ClientID/Zone1-9 mapping: {adt_file_used}"
            )
        else:
            annotated_client_ids = annotate_doc_with_zones(master_doc, zone_map)
            checked_messages.append(f"✓ Zone lookup completed from: {adt_file_used}")
            checked_messages.append(
                f"✓ ClientIDs updated with valid Zone1-9 info: {len(annotated_client_ids)}"
            )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = downloads / f"combined_latest_risk_docs_{timestamp}.docx"

    if added_any:
        master_doc.save(output_file)
        checked_messages.append(f"Combined document saved to: {output_file}")
    else:
        checked_messages.append("No valid .docx files with today's date were found to combine.")

    return {
        "success": added_any,
        "output_file": str(output_file) if added_any else None,
        "checked_messages": checked_messages,
        "skipped_doc_files": skipped_doc_files,
        "stale_or_invalid_files": stale_or_invalid_files,
        "adt_file_used": str(adt_file_used) if adt_file_used else None,
        "annotated_client_ids": annotated_client_ids
    }


if __name__ == "__main__":
    result = run()
    for msg in result["checked_messages"]:
        print(msg)

    if result.get("annotated_client_ids"):
        print("\nUpdated ClientIDs with valid Zone1-9:")
        for cid in result["annotated_client_ids"]:
            print(f" - {cid}")