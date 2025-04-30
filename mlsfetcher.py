#!/usr/bin/env python3
import io
import os
import requests
import pandas as pd
from msal import ConfidentialClientApplication

# ───────────────────────────────────────────────────────────────
# CONFIG: set these (or export as env vars)
CLIENT_ID     = os.getenv("AZURE_CLIENT_ID",     "430ac0be-52d5-4562-ba2a-4739140e638f")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "koI8Q~h9fRnfGGSC9zy3g.PdJsCTJ13wJwGSsdes")
TENANT_ID     = os.getenv("AZURE_TENANT_ID",     "d72741b9-6bf4-4282-8dfd-0af4f56d4023")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# these come from your earlier Graph calls
DRIVE_ID     = "b!BCUflbar8ka0_5exbILvkB5aHEMI7flArYOiUv-56dNWAeHXUqBXS6BBqmv_35m7"
ITEM_ID      = "012R5EVVNAQ23DVVPSV5GYCE7GRIK5D4FL"
STATE_SHEETS = ["Arizona","California","Nevada","Utah","Florida","Texas"]
# ───────────────────────────────────────────────────────────────

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

    # 2) Read only D:E from each state sheet
    xls = pd.read_excel(
        io.BytesIO(resp.content),
        sheet_name=STATE_SHEETS,
        usecols="D:E",
        engine="openpyxl",
    )

    # 3) Concat all sheets
    combined = pd.concat(xls.values(), ignore_index=True)

    # 4) Trim to exactly two columns
    combined = combined.iloc[:, :2]
    combined.columns = ["Club Code", "Address"]

    # 5) Remove any repeated header rows
    header_mask = (
        combined["Club Code"].astype(str).str.strip().str.lower() != "club code"
        & combined["Address"].astype(str).str.strip().str.lower() != "address"
    )
    combined = combined.loc[header_mask]

    # 6) Drop rows with missing or blank Address
    combined = combined[combined["Address"].notna()]
    combined = combined[combined["Address"].astype(str).str.strip() != ""]

    # 7) Strip whitespace from both columns
    combined["Club Code"] = combined["Club Code"].astype(str).str.strip()
    combined["Address"]   = combined["Address"].astype(str).str.strip()

    return combined

def main():
    print("🔐 Authenticating to Graph…")
    token = authenticate_graph()

    print("⬇️ Downloading and parsing the MLS workbook…")
    mls = fetch_master_data_graph(token)

    print(f"✅ Pulled {len(mls)} rows across {len(STATE_SHEETS)} sheets. Here’s a preview:")
    print(mls.head(10).to_string(index=False))

    # write out CSV for GitHub Actions or local use
    mls.to_csv("master_location_sheet.csv", index=False)
    print("✅ Wrote master_location_sheet.csv")

if __name__ == "__main__":
    main()
