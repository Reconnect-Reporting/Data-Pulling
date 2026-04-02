import os
import pandas as pd
import numpy as np


def parse_coord(coord_str):
    """'lat,lon' -> (lat, lon) as floats, or None if bad/NaN."""
    if pd.isna(coord_str):
        return None
    parts = str(coord_str).split(",")
    if len(parts) != 2:
        return None
    try:
        lat = float(parts[0].strip())
        lon = float(parts[1].strip())
        return lat, lon
    except ValueError:
        return None


def run():
    # Auto-detect user home directory
    user_home = os.path.expanduser("~")

    # Base directory paths
    treat_base = os.path.join(
        user_home,
        r"OneDrive - Reconnect Community Health Services (1)",
        "data",
        "Clean Data",
        "Treat"
    )

    templates_base = os.path.join(
        user_home,
        r"OneDrive - Reconnect Community Health Services (1)",
        "data",
        "Templates"
    )

    # Full file paths
    census_path = os.path.join(treat_base, "rpt_Census.csv")
    prog_path = os.path.join(treat_base, "rpt_ProgStaffFinance.csv")
    mapping_path = os.path.join(templates_base, "Mapping Back Up.csv")

    print("File paths:")
    print("Census:", census_path)
    print("Prog:", prog_path)
    print("Mapping:", mapping_path)

    # =========================
    # Read CSVs with ID as string
    # =========================

    df_census = pd.read_csv(
        census_path,
        dtype={"ID": str},
        encoding="utf-8-sig"
    )

    df_prog = pd.read_csv(
        prog_path,
        dtype={"ID": str},
        encoding="utf-8-sig"
    )

    df_mapping = pd.read_csv(
        mapping_path,
        dtype={"POSTAL_CODE": str},
        encoding="utf-8-sig"
    )

    # Keep only ID + Postal_Zip from census
    df_census = df_census[['ID', 'Postal_Zip']]

    # Filter df_prog by Program + Visit Type
    df_prog = df_prog[
        (df_prog['Program'].isin(['WeTo', 'SoTo', 'ACTT', 'PG ACTT'])) &
        (df_prog['Visit Type'].isin(['FaceHome', 'FaceOffice', 'FaceComm']))
    ].copy()

    # Parse Interaction Start Date
    df_prog["Interaction Start Date"] = pd.to_datetime(
        df_prog["Interaction Start Date"],
        errors="coerce"  # anything weird becomes NaT
    )

    # Last 7 calendar days
    today = pd.Timestamp.today().normalize()
    seven_days_ago = today - pd.Timedelta(days=6)

    mask = (
        (df_prog["Interaction Start Date"] >= seven_days_ago) &
        (df_prog["Interaction Start Date"] <= today + pd.Timedelta(days=1))
    )

    df_prog = df_prog.loc[mask].copy()

    # Merge Postal_Zip from census
    df_prog = pd.merge(df_prog, df_census, how='left', on='ID')

    # If Visit Type == 'FaceOffice', set Postal_Zip to 'M6E 1B8'
    df_prog.loc[df_prog["Visit Type"] == "FaceOffice", "Postal_Zip"] = "M6E 1B8"

    # Clean postal codes
    df_prog["Postal_Zip_clean"] = (
        df_prog["Postal_Zip"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(" ", "", regex=False)
    )

    df_mapping["POSTAL_CODE_clean"] = (
        df_mapping["POSTAL_CODE"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(" ", "", regex=False)
    )

    # 2. Merge to get LAT/LONG into df_prog
    df_merged = df_prog.merge(
        df_mapping[["POSTAL_CODE_clean", "LAT", "LONG"]],
        left_on="Postal_Zip_clean",
        right_on="POSTAL_CODE_clean",
        how="left"
    )

    # 3. Replace Postal_Zip with "LAT,LONG" in one cell
    df_merged["Postal_Zip"] = np.where(
        df_merged["LAT"].notna(),
        df_merged["LAT"].astype(str) + "," + df_merged["LONG"].astype(str),
        df_merged["Postal_Zip"]  # fallback: keep original if no match
    )

    # Drop helper columns
    df_merged = df_merged.drop(columns=["Postal_Zip_clean", "POSTAL_CODE_clean"])

    # This is now your working df
    df = df_merged

    # =========================
    # Build route-wide table
    # =========================

    # 1. Ensure Interaction Start Date is date
    df["Interaction Start Date"] = pd.to_datetime(df["Interaction Start Date"]).dt.date

    # 2. Convert Start Time to sortable time/datetime
    df["Start Time Sort"] = pd.to_datetime(
        df["Start Time"],
        format="%H:%M",
        errors="coerce"
    )

    # 3. Sort by Staff, Interaction Start Date, Start Time
    df_sorted = df.sort_values(
        ["Staff", "Interaction Start Date", "Start Time Sort"]
    ).reset_index(drop=True)

    # 4. Create stop number within each (Staff, Interaction Start Date)
    df_sorted["StopNum"] = df_sorted.groupby(
        ["Staff", "Interaction Start Date"]
    ).cumcount() + 1   # 1,2,3,...

    # 5. Pivot to wide: Stop1, Stop2, ...
    routes_wide = df_sorted.pivot_table(
        index=["Staff", "Interaction Start Date"],
        columns="StopNum",
        values="Postal_Zip",
        aggfunc="first"
    )

    # 6. Rename columns to Stop1, Stop2, ...
    routes_wide.columns = [f"Stop{int(c)}" for c in routes_wide.columns]

    # 7. Bring index back as columns
    routes_wide = routes_wide.reset_index()

    # Add starting and end point (clinic)
    start_end_coord = "43.678495, -79.437015"
    routes_wide["Starting point"] = start_end_coord
    routes_wide["End point"] = start_end_coord

    # =========================
    # Build long-format route table
    # =========================

    routes = routes_wide.copy()

    # 1. Clean / standardize date column
    routes["Interaction Date"] = pd.to_datetime(
        routes["Interaction Start Date"],
        errors="coerce"
    ).dt.date

    # 2. Identify stop columns in order: Stop1, Stop2, ...
    stop_cols = [c for c in routes.columns if c.startswith("Stop")]
    stop_cols = sorted(stop_cols, key=lambda x: int(x[4:]))  # sort by number

    records = []

    for _, row in routes.iterrows():
        staff = row["Staff"]
        date_val = row["Interaction Date"]

        # Route = [Starting point] + Stops (non-null)
        coords = []

        # Starting point as RouteOrder 0
        start_coord = parse_coord(row["Starting point"])
        if start_coord is not None:
            coords.append(start_coord)

        # Stops in order
        for col in stop_cols:
            coord = parse_coord(row[col])
            if coord is not None:
                coords.append(coord)

        # (If you want to ALSO add End point as final stop, uncomment below)
        # end_coord = parse_coord(row["End point"])
        # if end_coord is not None:
        #     coords.append(end_coord)

        # Now write them out with RouteOrder 0,1,2,...
        for order, (lat, lon) in enumerate(coords):
            records.append({
                "Staff": staff,
                "Interaction Date": date_val,
                "RouteOrder": order,
                "Lat": lat,
                "Lon": lon
            })

    # 4. Create long dataframe
    df_long = pd.DataFrame(records)

    # 5. Export to your Treat folder
    out_path = os.path.join(
        user_home,
        r"OneDrive - Reconnect Community Health Services (1)",
        "data",
        "Clean Data",
        "Treat",
        "routes_long_for_powerbi.csv"
    )

    df_long.to_csv(out_path, index=False, encoding="utf-8-sig")
    print("Exported to:", out_path)


def main():
    run()


if __name__ == "__main__":
    main()
