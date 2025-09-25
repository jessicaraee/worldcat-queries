# worldcat-queries
Workflows used to harvest metadata from the WorldCat API. WorldCat WSKey and secret key are required. Input files must be in .xlsx format.

### check_library_holdings.py
Check for holdings by OCLC number and library OCLC symbol or list of symbols. API documentation: https://developer.api.oclc.org/wcv2#/Member%20General%20Holdings/find-bib-holdings.

### get_lc_data.py
Harvest LC classification data by OCLC number. API documentation: https://developer.api.oclc.org/wc-metadata-v2#/Search%20Bibliographic%20Resources/get-classifications.

### get_oclc_numbers.py
Harvest OCLC numbers by ISBN or ISSN. API documentation: https://developer.api.oclc.org/wcv2#/Bibliographic%20Resources/search-brief-bibs.

### get_other_editions.py
Harvest brief bibliographic data of editions related to the input OCLC number. API documentation: https://developer.api.oclc.org/wcv2#/Bibliographic%20Resources/retrieve-other-editions.
