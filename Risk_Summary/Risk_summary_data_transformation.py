import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
import time


def main():
    # ================= CONFIG =================
    BASE = Path.home() / "OneDrive - Reconnect Community Health Services (1)" / "data"

    ALAYA = BASE / "Clean Data" / "AlayaCare"
    TREAT = BASE / "Clean Data" / "Treat"

    OUT_RISK = BASE / "risk_id"
    OUT_PROGRAMS = BASE / "risk_id_programs"

    time.sleep(5)  # wait for OneDrive

    # ================= LOAD =================
    progress = pd.read_csv(ALAYA / "notes_progress_look_daily.csv", dtype=str)
    visit = pd.read_csv(ALAYA / "notes_visit_look_daily.csv", dtype=str)
    client = pd.read_csv(ALAYA / "notes_client_look_daily.csv", dtype=str)

    # ================= STANDARDIZE =================
    progress = progress[[
        "AlayaCare Client ID",
        "Create Date",
        "Progress Note"
    ]]

    visit = visit[[
        "AlayaCare Client ID",
        "Note Created At Date",
        "Note"
    ]].rename(columns={
        "Note Created At Date": "Create Date",
        "Note": "Progress Note"
    })

    client = client[[
        "AlayaCare Client ID",
        "Creation Date",
        "Content"
    ]].rename(columns={
        "Creation Date": "Create Date",
        "Content": "Progress Note"
    })

    combined = pd.concat([progress, visit, client], ignore_index=True)

    # ================= ADT =================
    adt = pd.read_csv(ALAYA / "ADT with Zone and Reference.csv", dtype=str)
    adt.columns = adt.columns.str.strip()

    adt = adt[adt["Status"].str.lower() == "active"]

    adt = adt[[
        "ClientID",
        "ServiceCodeDept"
    ]]

    adt = adt.rename(columns={"ClientID": "AlayaCare Client ID"})

    # ================= MERGE =================
    x = combined.merge(adt, on="AlayaCare Client ID", how="left")

    x = x.drop_duplicates()

    x = x[~x["Progress Note"].str.contains(
        "Email communication regarding client",
        na=False
    )]

    # ================= SPLIT =================
    psw = x[x["ServiceCodeDept"] == "Personal Support Worker"]
    ads = x[x["ServiceCodeDept"] == "Adult Day Program"]
    sfs_cm = x[x["ServiceCodeDept"] == "Case Management"]
    ehts = x[x["ServiceCodeDept"] == "EHTS"]

    # ================= SAVE =================
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    def save(df, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(path, index=False)

    save(ehts, OUT_RISK / f"ehts_{timestamp}.csv")
    time.sleep(5)

    save(psw, OUT_PROGRAMS / "als" / f"assistedliving_{timestamp}.csv")
    time.sleep(5)

    # ================= TREAT DATA =================
    note_content = pd.read_csv(TREAT / "notes_content.csv", dtype=str)
    note_content.columns = note_content.columns.str.strip()

    note_content = note_content.rename(columns={
        note_content.columns[5]: "Program"
    })

    note_content["ID"] = note_content["ID"].ffill()

    note_content["InteractionDate\nY-M-D"] = pd.to_datetime(
        note_content["InteractionDate\nY-M-D"],
        errors="coerce"
    )

    yesterday = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)
    note_content = note_content[
        note_content["InteractionDate\nY-M-D"] >= yesterday
    ]

    note_content = note_content[note_content["ID"].notna()]

    # ================= PROGRAM SPLITS =================
    act = note_content[note_content["Program"].isin([
        "ACTT", "PG ACTT", "WeTo", "SoTo"
    ])]

    cm = note_content[note_content["Program"].isin([
        "Case Management", "COPA", "Virtual Case Management",
        "MHJ CM", "Streets 2 Home", "Continuing Service",
        "Transitional Aged Youth", "Mild to Moderate CM"
    ])]

    copa = note_content[note_content["Program"] == "COPA"]

    # ================= SAVE =================
    save(act, OUT_PROGRAMS / "actt" / f"actt_{timestamp}.csv")
    time.sleep(5)

    save(cm, OUT_PROGRAMS / "cm" / f"cm_{timestamp}.csv")
    time.sleep(5)

    save(copa, OUT_PROGRAMS / "copa" / f"copa_{timestamp}.csv")
    time.sleep(5)

    print("All files generated successfully.")


def run():
    main()


if __name__ == "__main__":
    run()