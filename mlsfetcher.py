#!/usr/bin/env python3
import io
import os
import requests
import pandas as pd
import logging
from msal import ConfidentialClientApplication

# ───────────────────────────────────────────────────────────────
# CONFIG: set these (or export as env vars)
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID",     "430ac0be-52d5-4562-ba2a-4739140e638f")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "koI8Q~h9fRnfGGSC9zy3g.PdJsCTJ13wJwGSsdes")
TENANT_ID     = os.getenv("AZURE_TENANT_ID",     "d72741b9-6bf4-4282-8dfd-0af4f56d4023")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# these sheets you need
STATE_SHEETS = ["Arizona","California","Nevada","Utah","Florida","Texas"]
# these IDs from earlier Graph calls
DRIVE_ID     = "b!BCUflbar8ka0_5exbILvkB5aHEMI7flArYOiUv-56dNWAeHXUqBXS6BBqmv_35m7"
ITEM_ID      = "012R5EVVNAQ23DVVPSV5GYCE7GRIK5D4FL"
# ───────────────────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format="%(message)s")

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

    # 1) Download workbook bytes
    url = f"{GRAPH_API_ENDPOINT}/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()

    # 2) Load into ExcelFile to inspect sheets
    excel_file = pd.ExcelFile(io.BytesIO(resp.content), engine="openpyxl")
    logging.info(f"Available sheets in workbook: {excel_file.sheet_names}")

    dfs = []
    for sheet in STATE_SHEETS:
        if sheet not in excel_file.sheet_names:
            logging.warning(f"Sheet '{sheet}' not found, skipping.")
            continue
        df = excel_file.parse(sheet_name=sheet, usecols="D:E")
        logging.info(f"'{sheet}': {len(df)} rows")
        dfs.append(df)

    if not dfs:
        raise RuntimeError("No valid state sheets found in workbook.")

    # 3) Combine and clean
    combined = pd.concat(dfs, ignore_index=True)
    combined = combined.iloc[:, :2]
    combined.columns = ["Club Code", "Address"]

    # filter out pulled headers
    mask_header = (
        combined["Club Code"].astype(str).str.strip().str.lower() == "club code"
    ) & (
        combined["Address"].astype(str).str.strip().str.lower() == "address"
    )
    combined = combined.loc[~mask_header]

    # drop empty addresses
    combined = combined[combined["Address"].notna()]
    combined = combined[combined["Address"].astype(str).str.strip() != ""]

    combined["Club Code"] = combined["Club Code"].astype(str).str.strip()
    combined["Address"]   = combined["Address"].astype(str).str.strip()

    return combined


def main():
    logging.info("🔐 Authenticating to Graph…")
    token = authenticate_graph()

    logging.info("⬇️ Downloading and parsing the MLS workbook…")
    mls = fetch_master_data_graph(token)

    logging.info(f"✅ Combined total: {len(mls)} rows across {len(STATE_SHEETS)} sheets")
    logging.info("Here’s a preview:")
    logging.info(mls.head(10).to_string(index=False))

    mls.to_csv("master_location_sheet.csv", index=False)
    logging.info("✅ Wrote master_location_sheet.csv")

if __name__ == "__main__":
    main()
