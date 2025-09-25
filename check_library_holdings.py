#Use the BookOps WorldCat wrapper to check for holdings based on a designated library (or list of libraries) and a list of OCLC numbers

import pandas as pd
import requests
from bookops_worldcat import WorldcatAccessToken
import time

#Configure access token
WORLDCAT_KEY = 'mykey' #Insert wskey here
WORLDCAT_SECRET = 'mysecret' #Insert secret key here
SCOPES = 'WorldCatMetadataAPI' #Update scopes as needed

#Configure files
INPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name
OUTPUT_FILE = 'FILENAME.xlsx' #Update to filepath and name

#Configure libraries
LIBRARY_SYMBOLS = ['LIBRARY'] #Update to OCLC symbol, if searching multiple libraries separate by comma

#Generate an access token
def get_token():
    return WorldcatAccessToken(
        key=WORLDCAT_KEY,
        secret=WORLDCAT_SECRET,
        scopes=SCOPES
    )

#Get library holdings data
def get_holdings_data(oclc_number, token, library_symbol):
    if isinstance(library_symbol, str):
        library_symbol = [library_symbol]

    try:
        url = f'https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs-holdings'
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }

        params = {
            'oclcNumber': str(oclc_number).strip(),
            'heldBySymbol': ','.join(library_symbol),
            'limit': 50
        }
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        rows = []

        number_of_records = data.get("numberOfRecords", 0)
        if number_of_records > 0:
            first_bib = data['briefRecords'][0]
            if 'institutionHolding' in first_bib and 'briefHoldings' in first_bib['institutionHolding']:
                brief_holdings = first_bib['institutionHolding']['briefHoldings']
                library_symbol = [entry['oclcSymbol'] for entry in brief_holdings if 'oclcSymbol' in entry]
                rows.append({
                    "OCLC_NUMBER": oclc_number,
                    "LIBRARY": library_symbol,
                    "LIBRARY_HOLDINGS_COUNT": len(library_symbol)
                })
            else:
                rows.append({
                    "OCLC_NUMBER": oclc_number,
                    "LIBRARY": [],
                    "LIBRARY_HOLDINGS_COUNT": 0
                })
        else:
            rows.append({
                "OCLC_NUMBER": oclc_number,
                "LIBRARY": [],
                "LIBRARY_HOLDINGS_COUNT": 0
            })
        return rows

    except Exception as e:
        print(f"[ERROR] {oclc_number}: {e}")
        return [{
            "OCLC_NUMBER": oclc_number,
            "LIBRARY": [],
            "LIBRARY_HOLDINGS_COUNT": 0
        }]

#Get summary data of total holdings in OCLC
def get_summary_data(oclc_number, token):
    try:
        url = f'https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs-summary-holdings'
        headers = {
            'Authorization': f'Bearer {token.token_str}',
            'Accept': 'application/json'
        }
        params = {'oclcNumber': str(oclc_number).strip()}
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        rows = []

        number_of_records = data.get("numberOfRecords", 0)
        if number_of_records > 0:
            first_bib_summary = data['briefRecords'][0]
            total_holding_count = first_bib_summary['institutionHolding']['totalHoldingCount']
            rows.append({
                "OCLC_NUMBER": str(oclc_number).strip(),
                "TOTAL_LIBRARIES_HOLDING_ITEM": total_holding_count
            })
        else:
            rows.append({
                "OCLC_NUMBER": str(oclc_number).strip(),
                "TOTAL_LIBRARIES_HOLDING_ITEM": 0
            })
        return rows

    except Exception as e:
        print(f"[ERROR] {oclc_number}: {e}")
        return [{
            "OCLC_NUMBER": "None",
            "TOTAL_LIBRARIES_HOLDING_ITEM": 0
        }]

#Run query and export results
def main():
    oclclist_df = pd.read_excel(INPUT_FILE, dtype={'RECORD_ID': str, 'OCLC_NUMBER': str})
    all_results = []

    token = get_token()

    #Set as True to fetch summary holdings data and harvest TOTAL_LIBRARIES_HOLDING_ITEM, set as False if not needed
    get_summary = True

    for _, row in oclclist_df.iterrows():
        oclc_number = row['OCLC_NUMBER']
        record_id = row['RECORD_ID']
        if not oclc_number:
            continue

        if token.is_expired():
            print("Refreshing token!")
            token = get_token()

        holdings_rows = get_holdings_data(oclc_number, token, LIBRARY_SYMBOLS)

        if get_summary:
            summary_rows = get_summary_data(oclc_number, token)
            for summary_row in summary_rows:
                for holdings_row in holdings_rows:
                    holdings_row.update(summary_row)

        all_results.extend(holdings_rows)
        print(f"Processed OCLC {oclc_number}")
        time.sleep(0.2)

    holdings_df = pd.DataFrame(all_results)
    merged_df = pd.merge(oclclist_df, holdings_df, on="OCLC_NUMBER", how="left")
    merged_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Data exported to {OUTPUT_FILE}.")

    #Optional filter to export separate file with only rows matching parameters, update or comment out as needed
    library_symbol_str = '_'.join(LIBRARY_SYMBOLS)
    filtered_df = merged_df[merged_df['LIBRARY'].apply(lambda x: any(symbol in str(x) for symbol in LIBRARY_SYMBOLS))]
    if not filtered_df.empty:
        filtered_file = OUTPUT_FILE.replace(".xlsx", f"_Filtered.xlsx")
        filtered_df.to_excel(filtered_file, index=False)
        print(f"Library holdings data exported to {filtered_file}")
    else:
        print(f"No rows matching {LIBRARY_SYMBOLS} found.")

if __name__ == "__main__":
    main()
    
