#Use the BookOps WorldCat wrapper to pull LC classification data based on a list of OCLC numbers

import pandas as pd
import requests
from bookops_worldcat import WorldcatAccessToken
import time
import re

#Configure access token
WORLDCAT_KEY = 'mykey' #Insert wskey here
WORLDCAT_SECRET = 'mysecret' #Insert secret key here
SCOPES = 'wcapi WorldCatMetadataAPI configPlatform'

#Configure files
INPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name
OUTPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name

def get_token():
    return WorldcatAccessToken(
        key=WORLDCAT_KEY,
        secret=WORLDCAT_SECRET,
        scopes=SCOPES
    )

def get_classification_bibs(oclc_number, token):
    try:
        url = f'https://metadata.api.oclc.org/worldcat/search/classification-bibs/{oclc_number}'
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }
        response = requests.get(url, headers=headers)

        # Raise error if request failed
        response.raise_for_status()
        return response.json()

    except requests.RequestException as e:
        print(f"[ERROR] Failed to fetch data for OCLC {oclc_number}: {e}")
        return {}

#Clean LC values to remove apostrophes/brackets, normalize spacing, and remove trailing punctuation
def clean_lc_class(lc_value):
    if not lc_value or lc_value == "None":
        return None
    lc_value = lc_value.strip().upper()
    lc_value = re.sub(r"[\'\[\]]", '', lc_value)
    lc_value = re.sub(r'\s+', ' ', lc_value)
    lc_value = re.sub(r'[.,;]+$', '', lc_value)
    if re.match(r'^[A-Z]{1,3}\d+', lc_value):
        return lc_value
    return None

def main():
    #Define INPUT_FILE fields as needed
    oclclist_df = pd.read_excel(INPUT_FILE, dtype={'ISBN': str, 'OCLC_NUMBER': str})

    subjects_data = []

    token = get_token()

    for oclc in oclclist_df['OCLC_NUMBER']:
        oclc = str(oclc).strip()
        if not oclc:
            continue

        if token.is_expired():
            print("Token expired, refreshing...")
            token = get_token()

        result = get_classification_bibs(oclc, token)

        lc_data = result.get('lc', {}).get('mostPopular', 'None')
        subjects_data.append({'OCLC_NUMBER': oclc, 'LC': lc_data})
        print(f"{oclc}, {lc_data}")
        time.sleep(0.2)  # Respectful throttling

    subjects_df = pd.DataFrame(subjects_data)

    final_df = oclclist_df.merge(subjects_df, on='OCLC_NUMBER', how='left')

    final_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Data exported to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
