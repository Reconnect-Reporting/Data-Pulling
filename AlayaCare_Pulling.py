# -*- coding: utf-8 -*-
import os
import re
from pathlib import Path
from glob import glob
from datetime import datetime, timedelta
import win32com.client as win32

# ========= CONFIG =========
OUTLOOK_ACCOUNT    = "recreporting@reconnect.on.ca"
SENDER_EMAIL       = "no-reply@alayamail.com"
INCLUDE_SUBFOLDERS = False  # Only Inbox

def _onedrive_root() -> Path:
    # Prefer OneDrive env vars (Business -> OneDriveCommercial)
    for var in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        p = os.environ.get(var)
        if p and Path(p).exists():
            return Path(p)
    # Fallback: first "OneDrive*" folder in the user profile
    home = Path.home()
    for child in home.iterdir():
        if child.is_dir() and child.name.startswith("OneDrive"):
            return child
    # Last resort
    return home / "OneDrive"

OD = _onedrive_root()

# Destinations (OneDrive AlayaCare)
RAW_BASE          = OD / "data" / "Raw Data" / "AlayaCare"
BASE_DIR          = str(RAW_BASE)             # ADT + Form + CM reports: replace behavior
ADT_DIR           = str(RAW_BASE)
NOTE_DIR          = str(RAW_BASE / "Note")
VISITS_DIR        = str(RAW_BASE / "Visits")
CLIENT_CALLS_DIR  = str(RAW_BASE / "Client Calls")

# Clean Data destination for HHRI (VHA)
CLEAN_DIR         = str(OD / "data" / "Clean Data" / "AlayaCare")

# ========= HELPERS =========
def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)

def outlook_dt_string(dt: datetime) -> str:
    # Outlook Restrict expects US format with a 12-hour clock
    return dt.strftime("%m/%d/%Y %I:%M %p")

def get_inbox_store_by_smtp(namespace, smtp_address: str):
    # Try to find a store that matches your account display; fall back to default
    for store in namespace.Stores:
        try:
            if store.DisplayName.lower() == smtp_address.lower():
                return store
        except Exception:
            pass
    return namespace.Stores.Item(1)

def iter_folder_and_children(folder):
    yield folder
    for f in folder.Folders:
        yield from iter_folder_and_children(f)

def get_ext_from_attachment(att) -> str:
    fn = att.FileName or ""
    ext = os.path.splitext(fn)[1]
    return ext if ext else ".dat"

def remove_existing_by_stem(dest_dir: str, stem: str):
    """
    Remove any previous files for this stem, including undated and dated variants:
      {stem}.*
      {stem}_*.*
    """
    patterns = [
        os.path.join(dest_dir, f"{stem}.*"),
        os.path.join(dest_dir, f"{stem}_*.*"),
    ]
    for pattern in patterns:
        for path in glob(pattern):
            try:
                os.remove(path)
            except PermissionError:
                try:
                    os.replace(path, path + ".bak")
                except Exception:
                    pass

# ---- Type detectors ----------------------------------------------------------
def _lname(fn: str) -> str:
    return (fn or "").strip().lower()

def is_adt_candidate(filename: str) -> bool:
    n = _lname(filename)
    return n.startswith("adt") or " adt" in n or "_adt" in n

def looks_like_zone_reference(filename: str) -> bool:
    n = _lname(filename)
    return ("zone" in n) and ("reference" in n)

def is_form_report(filename: str) -> bool:
    n = _lname(filename)
    return ("form" in n and "report" in n)

def is_note_report(filename: str) -> bool:
    n = _lname(filename)
    return n.startswith("note") or " note" in n or "notes" in n

def is_visits_report(filename: str) -> bool:
    n = _lname(filename)
    return n.startswith("visits") or " visits" in n or "visit " in n or " visit." in n

# --- Client Calls detector (robust) ---
_CLIENT_CALLS_RE = re.compile(r'^client[\s_-]*calls\b', re.IGNORECASE)
def is_client_calls_report(filename: str) -> bool:
    n = (filename or "").strip()
    return bool(_CLIENT_CALLS_RE.search(n))

# --- CM reports detectors (robust) ---
_CM_SUP_RE = re.compile(r'^cm[\s_-]*supervisors(\b|[_-])', re.IGNORECASE)
_CM_SUP_DISCH_RE = re.compile(r'^cm[\s_-]*supervisors[\s_-]*discharged(\b|[_-])', re.IGNORECASE)

def is_cm_supervisors_report(filename: str) -> bool:
    n = (filename or "").strip()
    return bool(_CM_SUP_RE.search(n)) and not bool(_CM_SUP_DISCH_RE.search(n))

def is_cm_supervisors_discharged_report(filename: str) -> bool:
    n = (filename or "").strip()
    return bool(_CM_SUP_DISCH_RE.search(n))

# --- HHRI(VHA) detector (robust: HHRI(VHA), HHRI VHA, HHRI-VHA, HHRI_VHA) ---
_HHRI_VHA_RE = re.compile(r'^hhri[\s_-]*\(?\s*vha\s*\)?', re.IGNORECASE)
def is_hhri_vha_report(filename: str) -> bool:
    n = (filename or "").strip()
    return bool(_HHRI_VHA_RE.search(n))

def pick_best_adt_attachment(attachments):
    """
    From a list of COM Attachment objects, choose the ADT file.
    Preference order:
      1) Name contains both 'zone' and 'reference'
      2) Otherwise, the largest by size
    """
    adt_atts = []
    for att in attachments:
        try:
            fn = (att.FileName or "").strip()
        except Exception:
            fn = ""
        if is_adt_candidate(fn):
            adt_atts.append(att)

    if not adt_atts:
        return None

    zr = [a for a in adt_atts if looks_like_zone_reference(a.FileName or "")]
    if zr:
        return max(zr, key=lambda a: getattr(a, "Size", 0))
    return max(adt_atts, key=lambda a: getattr(a, "Size", 0))

# ---- Save helpers ------------------------------------------------------------
def save_replace(dest_dir: str, stem: str, att) -> str:
    """Replace existing files for a given stem and save the attachment."""
    ensure_dir(dest_dir)
    remove_existing_by_stem(dest_dir, stem)
    ext = get_ext_from_attachment(att)
    dest_path = os.path.join(dest_dir, sanitize_filename(f"{stem}{ext}"))
    try:
        att.SaveAsFile(dest_path)
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest_path = os.path.join(dest_dir, sanitize_filename(f"{stem}_{ts}{ext}"))
        att.SaveAsFile(dest_path)
    return dest_path

def save_dated_replace_same_day(dest_dir: str, stem: str, att, day_str: str) -> str:
    """
    Replace any file for the SAME DATE only.
    Filename: {stem}_{YYYY-MM-DD}{ext}
    """
    ensure_dir(dest_dir)
    ext = get_ext_from_attachment(att)
    # remove files that match this exact date (keep other dates)
    remove_existing_by_stem(dest_dir, f"{stem}_{day_str}")
    dest_path = os.path.join(dest_dir, sanitize_filename(f"{stem}_{day_str}{ext}"))
    att.SaveAsFile(dest_path)
    return dest_path

# ========= MAIN =========
def main():
    # Ensure folders exist
    for p in (ADT_DIR, NOTE_DIR, VISITS_DIR, CLIENT_CALLS_DIR, CLEAN_DIR):
        ensure_dir(p)

    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = get_inbox_store_by_smtp(outlook, OUTLOOK_ACCOUNT)
    inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox

    # Rolling last 24 hours
    now = datetime.now()
    start_time = now - timedelta(days=1)
    day_str = now.strftime("%Y-%m-%d")  # use today's date for naming

    sender_filter = f"[SenderEmailAddress] = '{SENDER_EMAIL}'"
    start_filter  = f"[ReceivedTime] >= '{outlook_dt_string(start_time)}'"
    end_filter    = f"[ReceivedTime] <  '{outlook_dt_string(now)}'"

    folders = [inbox] if not INCLUDE_SUBFOLDERS else list(iter_folder_and_children(inbox))

    # Only save one ADT in the last 24h; same for HHRI(VHA)
    adt_done_last24h = False
    hhri_done_last24h = False

    for f in folders:
        items = f.Items
        items.Sort("[ReceivedTime]", True)  # newest first
        items = items.Restrict(sender_filter).Restrict(start_filter).Restrict(end_filter)

        for mail in list(items):
            try:
                if getattr(mail, "Class", None) != 43:  # 43 = MailItem
                    continue
                if mail.Attachments.Count == 0:
                    continue

                # ADT (replace) once
                if not adt_done_last24h:
                    best_adt = pick_best_adt_attachment(mail.Attachments)
                    if best_adt is not None:
                        saved = save_replace(ADT_DIR, "ADT with Zone and Reference", best_adt)
                        print(f"Saved ADT -> {saved}")
                        adt_done_last24h = True

                # Process all attachments for other types
                for i in range(1, mail.Attachments.Count + 1):
                    att = mail.Attachments.Item(i)
                    fn = (att.FileName or "").strip()
                    if not fn:
                        continue

                    # Skip ADT already handled
                    if adt_done_last24h and is_adt_candidate(fn):
                        continue

                    # --- HHRI(VHA): save newest once to Clean Data (replace fixed name) ---
                    if (not hhri_done_last24h) and is_hhri_vha_report(fn):
                        saved = save_replace(CLEAN_DIR, "HHRI(VHA)", att)
                        print(f"Saved HHRI(VHA) -> {saved}")
                        hhri_done_last24h = True
                        continue

                    # --- CM_Supervisors (replace, fixed name) ---
                    if is_cm_supervisors_report(fn):
                        saved = save_replace(ADT_DIR, "CM_Supervisors", att)
                        print(f"Saved CM_Supervisors -> {saved}")
                        continue

                    # --- CM_Supervisors_Discharged (replace, fixed name) ---
                    if is_cm_supervisors_discharged_report(fn):
                        saved = save_replace(ADT_DIR, "CM_Supervisors_Discharged", att)
                        print(f"Saved CM_Supervisors_Discharged -> {saved}")
                        continue

                    # Form Report: replace existing
                    if is_form_report(fn):
                        saved = save_replace(ADT_DIR, "Form Report", att)
                        print(f"Saved Form Report -> {saved}")
                        continue

                    # Note: replace previous SAME-DAY file, keep other days
                    if is_note_report(fn):
                        saved = save_dated_replace_same_day(NOTE_DIR, "Note", att, day_str)
                        print(f"Saved Note -> {saved}")
                        continue

                    # Visits: replace previous SAME-DAY file, keep other days
                    if is_visits_report(fn):
                        saved = save_dated_replace_same_day(VISITS_DIR, "Visits", att, day_str)
                        print(f"Saved Visits -> {saved}")
                        continue

                    # Client Calls: replace previous SAME-DAY file, keep other days
                    if is_client_calls_report(fn):
                        saved = save_dated_replace_same_day(CLIENT_CALLS_DIR, "Client Calls", att, day_str)
                        print(f"Saved Client Calls -> {saved}")
                        continue

            except Exception as e:
                print(f"Skipped one item due to error: {e}")

    if not adt_done_last24h:
        print("No ADT attachment found in the last 24 hours from the specified sender (that's okay if none were sent).")
    if not hhri_done_last24h:
        print("No HHRI(VHA) attachment found in the last 24 hours (that's okay if none were sent).")

if __name__ == "__main__":
    main()
