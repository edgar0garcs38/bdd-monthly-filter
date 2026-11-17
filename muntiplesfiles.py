import pandas as pd
import glob
import re

# List of your Excel file names (assuming they are in the same directory as the script)
excel_files = glob.glob('Change /*.xlsx')

# Create an empty list to store DataFrames
dfs = []

# Read each Excel file and append its DataFrame to the list
for file in excel_files:
    df = pd.read_excel(file)
    dfs.append(df)

# Concatenate all DataFrames into a single DataFrame
merged_df = pd.concat(dfs, ignore_index=True)

# Save the merged DataFrame to a new Excel file
merged_df.to_excel('CHANGE Unificada.xlsx', index=False)
print("Excel files merged successfully into 'CHANGE Unificada.xlsx'")

# take first and last name from merged file
fNamelName = pd.read_excel('CHANGE Unificada.xlsx', sheet_name='Sheet1')
# Step 1: Split at the first comma
split_names = fNamelName['ASEGURADO'].str.split(',', n=1, expand=True)

# Step 2: Assign and convert to uppercase
fNamelName['Last Name'] = split_names[0].str.upper()
fNamelName['First Name'] = split_names[1].str.upper()

if "E-MAIL" in fNamelName.columns:
    duplicates_removed = fNamelName[fNamelName.duplicated(subset=["E-MAIL"], keep=False)]
    duplicates_removed.to_excel('Duplicates_Removed.xlsx', index=False)
    print("Excel file with duplicates removed successfully created: 'Duplicates_Removed.xlsx'")

def is_valid_email(email: str) -> bool:
    """Validate email addresses, including multi-level domains, and reject empty values."""
    if not email or pd.isna(email):   # empty or NaN → invalid
        return False
    email = str(email).strip()
    if email == "":                   # blank string → invalid
        return False

    pattern = (
        r'^(?!.*\.\.)'                                # no consecutive dots
        r'[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]{0,62}[a-zA-Z0-9])?'  # local part
        r'@'
        r'(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+'  # domain labels
        r'[a-zA-Z]{2,}$'                              # final TLD (letters only, ≥2 chars)
    )
    return re.match(pattern, email) is not None
# Filter invalid emails from the DataFrame
invalid_emails = fNamelName[~fNamelName['E-MAIL'].apply(is_valid_email)]

# Export them
invalid_emails.to_excel('Invalid_Emails.xlsx', index=False)
print("Excel file with invalid emails successfully created: 'Invalid_Emails.xlsx'")

fNamelName.to_excel('CHANGE FNLN.xlsx', index=False)
print("Excel files with names successfully into 'CHANGE FNLN.xlsx'")

mailChimp_format = pd.read_excel('CHANGE FNLN.xlsx', sheet_name='Sheet1')

df = mailChimp_format[['E-MAIL', 'First Name', 'Last Name']].dropna(subset= ['E-MAIL'])

if "E-MAIL" in df.columns:
    df = df.drop_duplicates(subset=["E-MAIL"], keep="first")

df.to_csv('MailChimp_formatCHANGE NOV25.csv', index=False)
print("Excel file for MailChimp format successfully created: 'MailChimp_formatCHANGE NOV25.csv'")