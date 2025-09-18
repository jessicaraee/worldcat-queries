#Use the BookOps WorldCat wrapper to pull brief bib data about other editions based on a list of OCLC numbers

import pandas as pd
import requests
from bookops_worldcat import WorldcatAccessToken
import time
import re

#Configure access token
WORLDCAT_KEY = 'mykey' #Insert wskey here
WORLDCAT_SECRET = 'mysecret' #Insert secret key here
SCOPES = 'WorldCatMetadataAPI' #Update scopes as needed

#Configure files
INPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name
OUTPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name

#Generate an access token
def get_token():
    return WorldcatAccessToken(
        key=WORLDCAT_KEY,
        secret=WORLDCAT_SECRET,
        scopes=SCOPES
    )

#Get other editions
def get_other_editions(oclc_number, token, record_id):
    try:
        url = f'https://americas.discovery.api.oclc.org/worldcat/search/v2/brief-bibs/{oclc_number}/other-editions'
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        rows = []

        for rec in data.get("briefRecords", []):
            rec_oclcs = rec.get("oclcNumber", [])
            isbns = rec.get("isbns", [])
            format = rec.get("generalFormat", "None")
            specific_format = rec.get("specificFormat", "None")

            if isinstance(rec_oclcs, list) and rec_oclcs:
                for oclc in rec_oclcs:
                    isbn_str = ", ".join(isbns) if isinstance(isbns, list) else str(isbns)

                    rows.append({
                        "OTHER_OCLCS": str(oclc).strip(),
                        "OTHER_ISBNS": isbn_str.strip(),
                        "FORMAT": str(format),
                        "SPECIFIC_FORMAT": str(specific_format),
                        "RECORD_ID": record_id
                    })
            else:
                rows.append({
                    "OTHER_OCLCS": str(rec_oclcs).strip(),
                    "OTHER_ISBNS": " | ".join(isbns) if isinstance(isbns, list) else str(isbns),
                    "FORMAT": str(format),
                    "SPECIFIC_FORMAT": str(specific_format),
                    "RECORD_ID": record_id
                })
        if not rows: rows.append({"OTHER_OCLCS": "None", "OTHER_ISBNS": "None", "FORMAT": "None", "SPECIFIC_FORMAT": "None", "RECORD_ID": record_id})

        return rows

    except Exception as e:
        print(f"[ERROR] {oclc_number}: {e}")
        return [{"OTHER_OCLCS": "None", "OTHER_ISBNS": "None", "FORMAT": "None", "SPECIFIC_FORMAT": "None", "RECORD_ID": record_id}]

#Run query and export results
def main():
    oclclist_df = pd.read_excel(INPUT_FILE, dtype={'RECORD_ID':str, 'OCLC_NUMBER': str})

    all_results = []
    token = get_token()

    for _, row in oclclist_df.iterrows():
        oclc_number = row['OCLC_NUMBER']
        record_id = row['RECORD_ID']
        if not oclc_number:
            continue

        if token.is_expired():
            print("Token expired, refreshing...")
            token = get_token()

        rows = get_other_editions(oclc_number, token, record_id)
        all_results.extend(rows)

        print(f"{oclc_number}: {len(rows)} rows returned")
        time.sleep(0.2)

    results_df = pd.DataFrame(all_results)
    final_df = oclclist_df.merge(results_df, on="RECORD_ID", how="left")
    final_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Other editions data exported to {OUTPUT_FILE}.")

    #Optional filter to export separate file with only rows matching results
    filtered_df = final_df[final_df['SPECIFIC_FORMAT'] == 'PrintBook']
    if not filtered_df.empty:
        filtered_file = OUTPUT_FILE.replace(".xlsx", "_PrintOnly.xlsx")
        filtered_df.to_excel(filtered_file, index=False)
        print(f"PrintBook data exported to {filtered_file}")
    else:
        print("No rows with SPECIFIC_FORMAT = 'PrintBook' found.")

if __name__ == "__main__":
    main()
