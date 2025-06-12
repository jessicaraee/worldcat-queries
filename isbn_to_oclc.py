#Use the BookOps WorldCat wrapper to pull OCLC numbers based on a list of ISBNs

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

def get_oclc_numbers(isbn, token):
    try:
        url = f'https://americas.discovery.api.oclc.org/worldcat/search/v2/brief-bibs'
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }
        params = {'q': f'bn:{isbn}'}

        try:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            data = response.json()
            return data["briefRecords"][0]["oclcNumber"]
        
        except (KeyError, IndexError, TypeError):
            return {}
        
    except requests.RequestException as e:
        print(f"[ERROR] Failed to fetch data for ISBN {isbn}: {e}")
        return {}

  def main():
    #Define INPUT_FILE fields as needed
    isbnlist_df = pd.read_excel(INPUT_FILE, dtype={'TITLE': str, 'ISBN': str})

    oclc_numbers = []

    token = get_token()

    for isbn in isbnlist_df['ISBN']:
        isbn = str(isbn).strip()
        if not isbn:
            continue

        if token.is_expired():
            print("Token expired, refreshing...")
            token = get_token()

        result = get_oclc_numbers(isbn, token)

        oclc_number = result if result and result != "None" else "None"

        oclc_numbers.append({"ISBN": isbn, "OCLC_NUMBER": oclc_number})

        print(f"{isbn}, {result}")
        time.sleep(0.2)  # Respectful throttling

    oclc_numbers = pd.DataFrame(oclc_numbers)

    final_df = isbnlist_df.merge(oclc_numbers, on='ISBN', how='left')

    final_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Data exported to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
