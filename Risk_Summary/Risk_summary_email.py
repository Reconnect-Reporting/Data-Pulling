# Risk_Summary/send_latest_risk_doc_email.py

from __future__ import annotations

import json
import os
from pathlib import Path
from datetime import datetime
import pythoncom
import win32com.client


SETTINGS_FILE = Path(__file__).resolve().parent.parent / "user_settings.json"


def _load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _get_default_downloads_folder() -> Path:
    return Path.home() / "Downloads"


def _find_newest_risk_doc(downloads_folder: Path | None = None) -> Path:
    if downloads_folder:
        folders_to_check = [Path(downloads_folder)]
    else:
        folders_to_check = [
            Path.home() / "Downloads",
            Path.home() / "OneDrive - Reconnect Community Health Services" / "Downloads",
            Path.home() / "OneDrive - Reconnect Community Health Services (1)" / "Downloads",
        ]

    patterns = [
        "combined_lastest_risk_docs*.docx",
        "combined_lastest_risk_docs*.doc",
        "combined_latest_risk_docs*.docx",
        "combined_latest_risk_docs*.doc",
    ]

    matches = []
    for folder in folders_to_check:
        if not folder.exists():
            continue
        for pattern in patterns:
            matches.extend(folder.glob(pattern))

    if not matches:
        searched = "\n".join(str(f) for f in folders_to_check)
        raise FileNotFoundError(
            "No file found. Searched folders:\n"
            f"{searched}\n"
            "Expected names starting with 'combined_lastest_risk_docs' or "
            "'combined_latest_risk_docs' and ending with .doc or .docx"
        )

    newest = max(matches, key=lambda p: p.stat().st_mtime)
    return newest


def _get_outlook_account(session, sender_email: str):
    sender_email = (sender_email or "").strip().lower()
    for account in session.Accounts:
        smtp = ""
        try:
            smtp = (account.SmtpAddress or "").strip().lower()
        except Exception:
            pass

        if smtp == sender_email:
            return account
    return None


def send_latest_risk_doc_email(
    to_email: str | None = None,
    cc_email: str | None = None,
    sender_email: str = "recreporting@reconnect.on.ca",
    subject: str | None = None,
    body: str | None = None,
    downloads_folder: str | Path | None = None,
    send_now: bool = True,
):
    """
    Finds the newest Downloads file starting with combined_lastest_risk_docs
    and sends it through Outlook.

    Important:
    The sender mailbox must exist in the current Outlook profile and you must
    have permission to send from it.
    """
    settings = _load_settings()

    if not to_email:
        to_email = settings.get("RISK_SUMMARY_EMAIL_TO", "").strip()

    if not cc_email:
        cc_email = settings.get("RISK_SUMMARY_EMAIL_CC", "").strip()

    if not to_email:
        raise ValueError("Recipient email is empty. Please set RISK_SUMMARY_EMAIL_TO first.")

    newest_file = _find_newest_risk_doc(downloads_folder)

    if subject is None:
        subject = f"Latest Risk Summary Document - {datetime.now():%Y-%m-%d}"

    if body is None:
        body = (
            "Good morning,\n\n"
            "Please find attached your daily risk summary report. All programs are combined into one report, and you can navigate to your program(s) using the titles.\n\n"
            "Thank you!"
        )

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session

        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = to_email
        mail.CC = cc_email or ""
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(str(newest_file))

        account = _get_outlook_account(session, sender_email)
        if account is None:
            raise RuntimeError(
                f"Outlook account '{sender_email}' was not found in the current Outlook profile. "
                f"Make sure this mailbox is added and you have send permission."
            )

        # Send from specific mailbox
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # SendUsingAccount

        if send_now:
            mail.Send()
            action = "sent"
        else:
            mail.Display()
            action = "displayed"

        return {
            "status": action,
            "to": to_email,
            "cc": cc_email or "",
            "from": sender_email,
            "attachment": str(newest_file),
        }
    finally:
        pythoncom.CoUninitialize()


def run():
    """
    Keeps compatibility with your pipeline runner that expects run() or main().
    """
    return send_latest_risk_doc_email()


def main():
    result = send_latest_risk_doc_email()
    print(f"Email {result['status']} successfully.")
    print(f"To: {result['to']}")
    print(f"CC: {result['cc']}")
    print(f"From: {result['from']}")
    print(f"Attachment: {result['attachment']}")


if __name__ == "__main__":
    main()