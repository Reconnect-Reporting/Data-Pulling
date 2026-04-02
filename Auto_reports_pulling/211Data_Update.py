import json
import time
import os
from pathlib import Path
from datetime import datetime

import requests
import pandas as pd


# =========================
# CONFIG
# =========================
API_KEY = "kzGNi8Xj9cI1QvwRFOYZoe7VTWrgEy0a5dtJhpAnl3MLmSKqD4Cb6Hu2BxPsUf"
BASE_URL = "https://data.211support.org/api/v2/search"

# Your query (WORKING STYLE)
BASE_PAYLOAD = {
    "Dataset": "on",
    "Lang": "en",
    "Search": "match",
    "MatchMode": "all",
    "SearchType": "proximity",
    "Latitude": 43.6546431,
    "Longitude": -79.4614051,
    "Distance": 30,
    "SortOrder": "distance",
    "PageIndex": 0,
    "PageSize": 100,
    "Fields": "ApplicationProcess,TaxonomyTerm,TaxonomyTerms,TaxonomyCodes,Eligibility,FeeStructureSource,OfficialName,PhysicalCity",
    # Optional: FullRecord true if your dataset supports it (can be huge)
    # "FullRecord": True,
}

# Paging / rate limits
SLEEP_SECONDS = 0.35  # keep under 120 req/min (cloudflare headers show 120 limit)


# =========================
# HELPERS
# =========================
def downloads_dir() -> Path:
    candidates = [
        Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Downloads",
        Path.home() / "Downloads",
        Path.home() / "OneDrive" / "Downloads",
    ]
    for p in candidates:
        if p.exists():
            return p
    return Path.cwd()


def safe_request_json(url: str, payload: dict, timeout: int = 60) -> dict:
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=timeout)
    # If API sometimes returns HTML on errors, guard it:
    try:
        data = r.json()
    except Exception:
        raise RuntimeError(
            f"Non-JSON response. HTTP {r.status_code}. First 500 chars:\n{r.text[:500]}"
        )

    if isinstance(data, dict) and "Error" in data:
        raise RuntimeError(f"API Error: {data['Error']}")
    return data


def flatten_records(records: list[dict]) -> pd.DataFrame:
    """
    Turn Records (list of dicts) into a DataFrame.
    - Keeps nested lists/dicts as JSON strings so CSV export works.
    """
    if not records:
        return pd.DataFrame()

    df = pd.json_normalize(records, sep=".")
    # Convert any remaining un-serializable objects (lists/dicts) into JSON strings
    for c in df.columns:
        if df[c].apply(lambda x: isinstance(x, (list, dict))).any():
            df[c] = df[c].apply(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else x)
    return df


# =========================
# MAIN DOWNLOAD
# =========================
def download_all(payload_base: dict) -> tuple[list[dict], pd.DataFrame]:
    url = f"{BASE_URL}?key={API_KEY}"

    all_records: list[dict] = []
    seen_ids = set()

    page = 0
    page_size = int(payload_base.get("PageSize", 100))

    total_expected = None

    while True:
        payload = dict(payload_base)
        payload["PageIndex"] = page

        print("=" * 90)
        print(f"[PAGE] {page} | PageSize={page_size}")
        data = safe_request_json(url, payload)

        record_count = data.get("RecordCount", None)
        records = data.get("Records", []) or []

        if total_expected is None and isinstance(record_count, int):
            total_expected = record_count
            print(f"[TOTAL RecordCount] {total_expected}")

        print(f"[THIS PAGE] Records returned: {len(records)}")

        if len(records) == 0:
            print("[STOP] 0 records returned on this page.")
            break

        # Deduplicate by id and/or CurrentId
        added = 0
        for rec in records:
            rid = rec.get("id", None)
            cid = rec.get("CurrentId", None)

            # Prefer 'id', fallback to CurrentId, else hash the JSON
            key = rid if rid is not None else (cid if cid is not None else json.dumps(rec, sort_keys=True))
            if key in seen_ids:
                continue
            seen_ids.add(key)
            all_records.append(rec)
            added += 1

        print(f"[DEDUP] Added {added} new records | Running total: {len(all_records)}")

        # Stopping condition: if we got less than a full page, likely last page
        if len(records) < page_size:
            print("[STOP] Last page (returned less than PageSize).")
            break

        page += 1
        time.sleep(SLEEP_SECONDS)

        # Optional safety stop if API reports total_expected and we already got enough
        if total_expected is not None and len(all_records) >= total_expected:
            print("[STOP] Reached RecordCount.")
            break

    df = flatten_records(all_records)
    return all_records, df


def main():
    # Auto-detect Windows username / home directory
    home_dir = Path.home()

    out_dir = home_dir / "OneDrive - Reconnect Community Health Services (1)" / "data" / "Raw Data" / "211 Find Help Data"

    # Ensure the folder exists
    out_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    jsonl_path = out_dir / f"findhelp_211_results.jsonl"
    csv_path = out_dir / f"findhelp_211_results.csv"

    records, df = download_all(BASE_PAYLOAD)

    # Save JSONL
    with open(jsonl_path, "w", encoding="utf-8") as f:
        for rec in records:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    # Save CSV
    if df.empty:
        print("[WARN] No records -> CSV will not be written.")
    else:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print("\n" + "=" * 90)
    print(f"[DONE] Downloaded records: {len(records)}")
    print(f"[SAVED JSONL] {jsonl_path}")
    if not df.empty:
        print(f"[SAVED CSV]  {csv_path}  ({df.shape[0]} rows, {df.shape[1]} cols)")


if __name__ == "__main__":
    main()