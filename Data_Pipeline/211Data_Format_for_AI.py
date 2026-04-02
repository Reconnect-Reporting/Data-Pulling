# 211Data_Format_To_Target.py

from pathlib import Path
import pandas as pd
import re


def main():
    # =========================
    # Paths (auto-detect user)
    # =========================
    home = Path.home()

    input_folder = home / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Clean Data" / "211 Find Help Data"
    output_folder1 = home / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Clean Data" / "211 Find Help Data"
    output_folder2 = home / "Reconnect Community Health Services" / "Health Services - Documents" / "_library"

    output_folder1.mkdir(parents=True, exist_ok=True)
    output_folder2.mkdir(parents=True, exist_ok=True)

    input_file = input_folder / "findhelp_211_cleaned.csv"
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    print("Reading file:")
    print(input_file)

    df = pd.read_csv(input_file)

    # =========================
    # Helpers
    # =========================
    def null_like(x):
        if pd.isna(x):
            return True
        s = str(x).strip()
        return s == "" or s.lower() in {"nan", "none", "null"}

    def clean_text(x):
        if null_like(x):
            return ""
        s = str(x).strip()
        s = re.sub(r"\s+", " ", s)
        return s

    def fill_none(x):
        s = clean_text(x)
        return s if s else "None"

    def split_public_name(name):
        """
        Split PublicName into Organization + Program.
        Strategy:
        - If there is ' - ', split on the first one
        - Else whole name goes to Organization and Program is None
        """
        s = clean_text(name)
        if not s:
            return "None", "None"

        parts = [p.strip() for p in s.split(" - ", 1)]
        if len(parts) == 2:
            org, prog = parts
            return fill_none(org), fill_none(prog)
        else:
            return fill_none(s), "None"

    def build_criteria(row):
        parts = []

        eligibility = clean_text(row.get("Eligibility", ""))
        if eligibility:
            parts.append(eligibility)

        min_age = row.get("MinAge", "")
        max_age = row.get("MaxAge", "")

        age_parts = []
        if not null_like(min_age):
            try:
                min_age = int(float(min_age))
                age_parts.append(f"Minimum age: {min_age}.")
            except Exception:
                age_parts.append(f"Minimum age: {clean_text(min_age)}.")
        if not null_like(max_age):
            try:
                max_age = int(float(max_age))
                age_parts.append(f"Maximum age: {max_age}.")
            except Exception:
                age_parts.append(f"Maximum age: {clean_text(max_age)}.")

        if age_parts:
            parts.append(" ".join(age_parts))

        result = " ".join(parts).strip()
        return result if result else "None"

    def build_address(row):
        address_parts = [
            clean_text(row.get("PhysicalAddressStreet1", "")),
            clean_text(row.get("PhysicalAddressStreet2", "")),
            clean_text(row.get("PhysicalAddressCity", "")),
            clean_text(row.get("PhysicalAddressProvince", "")),
            clean_text(row.get("PhysicalAddressPostalCode", "")),
            clean_text(row.get("PhysicalAddressCountry", "")),
        ]
        address_parts = [x for x in address_parts if x]
        if not address_parts:
            return "None"
        return ", ".join(address_parts)

    def first_item(text):
        """
        Return first comma-separated item
        """
        s = clean_text(text)
        if not s:
            return ""
        return s.split(",")[0].strip()

    def infer_type(taxonomy_term):
        """
        Use first item from TaxonomyTerm as Type
        """
        s = first_item(taxonomy_term)
        return s if s else "None"

    def infer_primary_issue(taxonomy_terms):
        """
        Use first broad category before the first ' - ' from TaxonomyTerms
        Example:
        'Basic Needs - Food - Meals - Home Delivered Meals'
        -> 'Basic Needs'
        """
        s = first_item(taxonomy_terms)
        if not s:
            return "None"
        return s.split(" - ")[0].strip() if " - " in s else s

    def infer_population(eligibility):
        """
        Keep it simple:
        Use the first chunk of Eligibility before 'Residency Requirements'
        or before first comma-heavy details if possible.
        """
        s = clean_text(eligibility)
        if not s:
            return "None"

        s = re.split(r"Residency Requirements", s, flags=re.IGNORECASE)[0].strip(" ,;.-")

        sentence_split = re.split(r"[.;]", s)
        if sentence_split and sentence_split[0].strip():
            first_sentence = sentence_split[0].strip()
        else:
            first_sentence = s

        if len(first_sentence) > 90:
            first_sentence = first_sentence.split(",")[0].strip()

        return first_sentence if first_sentence else "None"

    def standardize_link(x):
        s = clean_text(x)
        if not s:
            return "None"
        if s.startswith("http://") or s.startswith("https://"):
            return s
        return "https://" + s

    # =========================
    # Build target dataframe
    # =========================
    if "PublicName" in df.columns:
        org_prog = df["PublicName"].apply(split_public_name)
        org_prog_df = pd.DataFrame(org_prog.tolist(), columns=["Organization", "Program"], index=df.index)
    else:
        org_prog_df = pd.DataFrame(
            [("None", "None")] * len(df),
            columns=["Organization", "Program"],
            index=df.index
        )

    out = pd.DataFrame(index=df.index)
    out["Organization"] = org_prog_df["Organization"]
    out["Program"] = org_prog_df["Program"]
    out["Description"] = df["Description"].apply(fill_none) if "Description" in df.columns else "None"
    out["Criteria"] = df.apply(build_criteria, axis=1)
    out["How_to_Apply"] = df["ApplicationProcess"].apply(fill_none) if "ApplicationProcess" in df.columns else "None"
    out["Primary_Contact"] = "None"
    out["Address"] = df.apply(build_address, axis=1)
    out["Phone"] = df["PhoneNumbers"].apply(fill_none) if "PhoneNumbers" in df.columns else "None"
    out["Email"] = df["Email"].apply(fill_none) if "Email" in df.columns else "None"
    out["Fees"] = df["FeeStructureSource"].apply(fill_none) if "FeeStructureSource" in df.columns else "None"
    out["Catchment"] = df["ServiceArea"].apply(fill_none) if "ServiceArea" in df.columns else "None"
    out["Link"] = df["Website"].apply(standardize_link) if "Website" in df.columns else "None"
    out["Population"] = df["Eligibility"].apply(infer_population) if "Eligibility" in df.columns else "None"
    out["Primary_Issue"] = df["TaxonomyTerms"].apply(infer_primary_issue) if "TaxonomyTerms" in df.columns else "None"
    out["Type"] = df["TaxonomyTerm"].apply(infer_type) if "TaxonomyTerm" in df.columns else "None"

    # =========================
    # Final cleanup
    # =========================
    for col in out.columns:
        out[col] = out[col].apply(fill_none)

    out = out.drop_duplicates()

    # =========================
    # Save to TWO locations
    # =========================
    file_name = "findhelp_211_formatted_for_ai.csv"

    output_file1 = output_folder1 / file_name
    output_file2 = output_folder2 / file_name

    out.to_csv(output_file1, index=False, encoding="utf-8-sig")
    out.to_csv(output_file2, index=False, encoding="utf-8-sig")

    print("\nFormatted dataset saved to:")
    print(output_file1)
    print(output_file2)

    print("\nPreview:")
    print(out.head())

    return output_file1, output_file2


def run():
    return main()


if __name__ == "__main__":
    main()