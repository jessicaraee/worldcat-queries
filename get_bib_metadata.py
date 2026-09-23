#Use the BookOps WorldCat wrapper to return bibliographic metadata for a list of OCLC numbers

import pandas as pd
import requests
from bookops_worldcat import WorldcatAccessToken
import time
import json

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
#Flatten nested JSON
def extract_text(value):
    if isinstance(value, str):
        return value
    
    if isinstance(value, dict):
        preferred_keys = [
            "text",
            "statement",
            "publisherName",
            "machineReadableDate",
            "name"
        ]

        for key in preferred_keys:
            if key in value:
                result = extract_text(value[key])
                if result:
                    return result

        for v in value.values():
            result = extract_text(v)
            if result:
                return result

    if isinstance(value, list):
        for item in value:
            result = extract_text(item)
            if result:
                return result

    return None

#Query metadata
def get_bib_metadata(oclc_number, token):
    try:
        url = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs"
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }
        params = {
            'q': f'no:{str(oclc_number).strip()}',
            'limit': 1
        }
        response = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=30
        )

        response.raise_for_status()
        data = response.json()
        records = []

        if 'briefRecords' in data:
            records = data['briefRecords']

        elif 'bibRecords' in data:
            records = data['bibRecords']

        #Add or change fields as needed
        if not records:
            return {
                "OCLC_NUMBER": str(oclc_number).strip(),
                "TITLE": None,
                "PUBLICATION_YEAR": None,
                "PUBLISHER": None,
                "EDITION": None
            }

        record = records[0]

        #Debugging step to determine data structure, uncomment as needed
        #print(json.dumps(record, indent=2))

        title = extract_text(
            record.get("title", {}).get("mainTitles")
        )

        publication_year = extract_text(
            record.get("date")
        )

        publisher = extract_text(
            record.get("publishers")
        )

        edition = extract_text(
            record.get("edition")
        )

        #Add or change fields as needed
        return {
            "OCLC_NUMBER": str(oclc_number).strip(),
            "TITLE": title,
            "PUBLICATION_YEAR": publication_year,
            "PUBLISHER": publisher,
            "EDITION": edition
        }

    except Exception as e:

        print(f"[ERROR] {oclc_number}: {e}")

        #Add or change fields as needed
        return {
            "OCLC_NUMBER": str(oclc_number).strip(),
            "TITLE": None,
            "PUBLICATION_YEAR": None,
            "PUBLISHER": None,
            "EDITION": None
        }

#Main process
def main():
    oclclist_df = pd.read_excel(
        INPUT_FILE,
        dtype={
            'RECORD_ID': str,
            'OCLC_NUMBER': str
        }
    )

    all_results = []

    token = get_token()

    for _, row in oclclist_df.iterrows():

        oclc_number = row['OCLC_NUMBER']

        if pd.isna(oclc_number):
            continue

        if token.is_expired():
            print("Refreshing token!")
            token = get_token()

        metadata = get_bib_metadata(
            oclc_number,
            token
        )

        all_results.append(metadata)
        print(f"Processed OCLC {oclc_number}")

        time.sleep(0.2)

    metadata_df = pd.DataFrame(all_results)

    merged_df = pd.merge(
        oclclist_df,
        metadata_df,
        on="OCLC_NUMBER",
        how="left"
    )

    merged_df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print(f"\nDone! Metadata exported to {OUTPUT_FILE}.")

if __name__ == "__main__":
    main()
