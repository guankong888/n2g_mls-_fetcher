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

# Ensure script directory as working directory so file writes are predictable
def ensure_working_dir():
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()
    os.chdir(base)
    logging.info(f"Working directory set to {base}")


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


def fetch_master_data_graph(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_API_ENDPOINT}/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()

    excel_file = pd.ExcelFile(io.BytesIO(resp.content), engine="openpyxl")
    available = [s.strip() for s in excel_file.sheet_names]
    logging.info(f"Available sheets: {available}")

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
        raise RuntimeError("None of the desired sheets were loaded.")

    combined = pd.concat(dfs, ignore_index=True)
    combined = combined.iloc[:, :2]
    combined.columns = ["Club Code","Address"]

    # remove repeated header rows
    mask_header = (
        combined["Club Code"].astype(str).str.lower().eq("club code") &
        combined["Address"].astype(str).str.lower().eq("address")
    )
    combined = combined.loc[~mask_header]

    # drop empty addresses
    combined = combined[combined["Address"].notna()]
    combined = combined[combined["Address"].str.strip().ne("")]

    combined["Club Code"] = combined["Club Code"].str.strip()
    combined["Address"]   = combined["Address"].str.strip()

    return combined


def write_csv(df, path):
    try:
        df.to_csv(path, index=False)
        size = os.path.getsize(path)
        mtime = datetime.fromtimestamp(os.path.getmtime(path)).isoformat()
        logging.info(f"Wrote {path} ({size:,} bytes, modified: {mtime})")
    except Exception as e:
        logging.error(f"Failed to write CSV to {path}: {e}")
        raise


def main():
    ensure_working_dir()
    logging.info("🔐 Authenticating to Graph…")
    token = authenticate_graph()

    logging.info("⬇️ Fetching and parsing workbook…")
    mls = fetch_master_data_graph(token)

    logging.info(f"✅ Combined total rows: {len(mls)} across {len(STATE_SHEETS)} sheets.")
    logging.info("--- Full dataset ---")
    logging.info("\n" + mls.to_string(index=False))

    write_csv(mls, OUTPUT_FILE)

if __name__ == "__main__":
    main()
