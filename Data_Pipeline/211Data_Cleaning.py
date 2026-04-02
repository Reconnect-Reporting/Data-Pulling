# 211Data_Cleaning.py
from pathlib import Path
import pandas as pd
import re
import ast


def main():
    # =========================
    # Locate dataset (auto-detect user)
    # =========================
    home = Path.home()
    folder = home / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Raw Data" / "211 Find Help Data"
    output_folder = home / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Clean Data" / "211 Find Help Data"
    output_folder.mkdir(parents=True, exist_ok=True)

    files = sorted(folder.glob("findhelp_211_results*.csv"))
    if not files:
        raise FileNotFoundError(f"No findhelp_211_results*.csv found in: {folder}")

    file = files[-1]
    print("Reading file:")
    print(file)

    df = pd.read_csv(file)

    # =========================
    # Drop unwanted columns
    # =========================
    cols_to_remove = [
        "TaxonomyCodes",
        "RecordOwner",
        "OfficialName",
        "Hours2",
        "id",
        "CurrentId",
        "ParentId",
        "Score",
        "Distance",
        "MailingAttentionName",
    ]
    df = df.drop(columns=[c for c in cols_to_remove if c in df.columns])

    # =========================
    # Remove HTML tags (remove anything in < >)
    # =========================
    TAG_RE = re.compile(r"<.*?>")

    def remove_html(x):
        if pd.isna(x):
            return x
        s = str(x)
        s = TAG_RE.sub(" ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    df = df.applymap(remove_html)

    # =========================
    # Clean ServiceArea: remove [] and quotes
    # =========================
    def clean_service_area(x):
        if pd.isna(x):
            return x
        return str(x).replace("[", "").replace("]", "").replace('"', "").strip()

    if "ServiceArea" in df.columns:
        df["ServiceArea"] = df["ServiceArea"].apply(clean_service_area)

    # =========================
    # Replace separators in specific columns
    # =========================
    def replace_separators(x, replace_colon=False):
        if pd.isna(x):
            return x
        s = str(x)
        s = s.replace("*", ",")
        s = s.replace(";", ",")
        if replace_colon:
            s = s.replace(":", ",")
        s = re.sub(r"\s*,\s*", ", ", s)          # normalize comma spacing
        s = re.sub(r"(,\s*){2,}", ", ", s)       # multiple commas -> one
        s = re.sub(r"\s+", " ", s).strip()
        return s

    if "Description" in df.columns:
        df["Description"] = df["Description"].apply(lambda x: replace_separators(x, replace_colon=False))

    for col in ["TaxonomyTerm", "TaxonomyTerms", "Eligibility"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: replace_separators(x, replace_colon=True))

    if "Hours" in df.columns:
        df["Hours"] = df["Hours"].apply(lambda x: replace_separators(x, replace_colon=False))

    for col in ["ApplicationProcess", "FeeStructureSource"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: replace_separators(x, replace_colon=True))

    # =========================
    # Clean PhoneNumbers: JSON/list -> "Name: Phone, Name: Phone"
    # =========================
    def clean_phone_numbers(x):
        if pd.isna(x):
            return x

        s = str(x).strip()
        if s == "":
            return ""

        try:
            phones = ast.literal_eval(s)
            if not isinstance(phones, list):
                return s

            formatted = []
            for p in phones:
                if not isinstance(p, dict):
                    continue
                phone = str(p.get("Phone", "")).strip()
                name = str(p.get("Name", "")).strip()
                desc = str(p.get("Description", "")).strip()
                ptype = str(p.get("Type", "")).strip()

                if not phone:
                    continue

                label_parts = [t for t in [name, ptype, desc] if t]
                label = " - ".join(label_parts)

                formatted.append(f"{label}: {phone}" if label else phone)

            # de-dupe while preserving order
            seen = set()
            out = []
            for item in formatted:
                if item not in seen:
                    seen.add(item)
                    out.append(item)

            return ", ".join(out)

        except Exception:
            return s

    if "PhoneNumbers" in df.columns:
        df["PhoneNumbers"] = df["PhoneNumbers"].apply(clean_phone_numbers)

    # =========================
    # Remove trailing punctuation at end of a cell
    # =========================
    TRAIL_RE = re.compile(r"[,:;\*\s]+$")

    def remove_trailing_punct(x):
        if pd.isna(x):
            return x
        return TRAIL_RE.sub("", str(x))

    df = df.applymap(remove_trailing_punct)

    # =========================
    # Save cleaned dataset
    # =========================
    output = output_folder / "findhelp_211_cleaned.csv"
    df.to_csv(output, index=False, encoding="utf-8-sig")

    print("\nCleaned dataset saved to:")
    print(output)

    print("\nPreview:")
    print(df.head())

    # Optional: return path so your GUI can show it / log it
    return output


# Your GUI runner looks for run() OR main()
def run():
    return main()


if __name__ == "__main__":
    main()