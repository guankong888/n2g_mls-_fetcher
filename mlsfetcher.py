#!/usr/bin/env python3
import io
import os
import requests
import pandas as pd
import logging
from msal import ConfidentialClientApplication
from datetime import datetime

# ───────────────────────────────────────────────────────────────
# CONFIG: set these (or export as env vars)
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID",     "430ac0be-52d5-4562-ba2a-4739140e638f")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "koI8Q~h9fRnfGGSC9zy3g.PdJsCTJ13wJwGSsdes")
TENANT_ID     = os.getenv("AZURE_TENANT_ID",     "d72741b9-6bf4-4282-8dfd-0af4f56d4023")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# sheets to pull (case-insensitive, trimmed)
STATE_SHEETS = ["Arizona","California","Nevada","Utah","Florida","Texas"]
# drive + item IDs
DRIVE_ID = "b!BCUflbar8ka0_5exbILvkB5aHEMI7flArYOiUv-56dNWAeHXUqBXS6BBqmv_35m7"
ITEM_ID  = "012R5EVVNAQ23DVVPSV5GYCE7GRIK5D4FL"
# output file
OUTPUT_FILE = "master_location_sheet.csv"
# ───────────────────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")

# Ensure script directory as working directory
def ensure_working_dir():
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()
    os.chdir(base)
    logging.info(f"Working directory set to {base}")

# Authenticate and get access token
def authenticate_graph():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError("Graph authentication failed: " + result.get("error_description","<no error>"))
    return result["access_token"]

# Fetch and combine sheets
def fetch_master_data_graph(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    # Download workbook bytes
    url = f"{GRAPH_API_ENDPOINT}/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()

    # Load Excel file
    excel_file = pd.ExcelFile(io.BytesIO(resp.content), engine="openpyxl")
    available = [s.strip() for s in excel_file.sheet_names]
    logging.info(f"Available sheets: {available}")

    # Parse each desired sheet
    dfs = []
    lower_map = {s.strip().lower(): s for s in excel_file.sheet_names}
    for desired in STATE_SHEETS:
        actual = lower_map.get(desired.strip().lower())
        if not actual:
            logging.warning(f"Sheet matching '{desired}' not found, skipping.")
            continue
        df = excel_file.parse(sheet_name=actual, usecols="D:E")
        logging.info(f"'{actual}': pulled {len(df)} rows")
        dfs.append(df)

    if not dfs:
        logging.error("No sheets loaded. Returning empty DataFrame.")
        return pd.DataFrame(columns=["Club Code","Address"])

    # Log pre-concat totals
    total_pre = sum(len(df) for df in dfs)
    logging.info(f"Total rows loaded before concat: {total_pre}")

    # Concatenate
    combined = pd.concat(dfs, ignore_index=True)
    logging.info(f"Combined rows after concat: {len(combined)}")

    # Trim to two columns and rename
    combined = combined.iloc[:, :2]
    combined.columns = ["Club Code","Address"]
    logging.info(f"Dataset shape after trimming: {combined.shape}")

    # Remove any in-band header rows
    mask_header = (
        combined["Club Code"].astype(str).str.lower().eq("club code") &
        combined["Address"].astype(str).str.lower().eq("address")
    )
    removed = mask_header.sum()
    if removed:
        logging.info(f"Removing {removed} in-band header rows")
    combined = combined.loc[~mask_header]
    logging.info(f"Rows after header removal: {len(combined)}")

    # Drop empty addresses
    before_drop = len(combined)
    combined = combined[combined["Address"].notna()]
    combined = combined[combined["Address"].str.strip().ne("")]
    after_drop = len(combined)
    logging.info(f"Dropped {before_drop - after_drop} empty rows; remaining: {after_drop}")

    # Strip whitespace
    combined["Club Code"] = combined["Club Code"].astype(str).str.strip()
    combined["Address"]   = combined["Address"].astype(str).str.strip()

    return combined

# Write CSV with explicit overwrite and logging
def write_csv(df, path):
    try:
        df.to_csv(path, index=False)
        size = os.path.getsize(path)
        mtime = datetime.fromtimestamp(os.path.getmtime(path)).isoformat()
        logging.info(f"Wrote {path} ({size:,} bytes, modified: {mtime})")
    except Exception as e:
        logging.error(f"Failed to write CSV to {path}: {e}")
        raise

# Main execution
def main():
    ensure_working_dir()

    # Fetch data
    try:
        logging.info("🔐 Authenticating to Graph…")
        token = authenticate_graph()
        logging.info("⬇️ Fetching and parsing workbook…")
        mls = fetch_master_data_graph(token)
    except Exception as e:
        logging.error(f"Error during fetch: {e}")
        mls = pd.DataFrame(columns=["Club Code","Address"])

    # Log and write
    logging.info(f"✅ Final dataset rows: {len(mls)}")
    write_csv(mls, OUTPUT_FILE)

if __name__ == "__main__":
    main()
