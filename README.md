# Monthly Mailing List Builder

Automates the monthly preparation of email marketing lists from insurance-portfolio exports, turning raw Excel reports into Mailchimp-ready CSVs.

## What it does
- Merges multiple monthly Excel exports into one dataset (`muntiplesfiles.py`) or processes a single file (`singlefile.py`)
- Splits the client's full name (`"LAST NAMES, FIRST NAMES"`) into separate first/last name columns
- Removes duplicate records
- Validates email addresses and exports the invalid ones separately for review
- Outputs a CSV in Mailchimp's import format

Used every month for different campaigns and insurers (health, life and accident products).

## Tech stack
Python · pandas · regex

## Run it
```bash
pip install -r requirements.txt
python singlefile.py      # one file
python muntiplesfiles.py  # merge a folder of files
```
Edit the input/output paths at the top of each script for the month being processed.

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
