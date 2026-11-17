import re
import pandas as pd

# take first and last name from merged file
fNamelName = pd.read_excel('Mar26/Mejoras en la cobertura de accidentes personales/Mejoras en la cobertura de accidentes personales.xlsx', sheet_name='Worksheet')
# Step 1: Split at the first comma
split_names = fNamelName['ASEGURADO'].str.split(',', n=1, expand=True)

# Step 2: Assign and convert to uppercase
fNamelName['Last Name'] = split_names[0].str.upper()
fNamelName['First Name'] = split_names[1].str.upper()

if "E-MAIL" in fNamelName.columns:
    duplicates_removed = fNamelName[fNamelName.duplicated(subset=["E-MAIL"], keep=False)]
    duplicates_removed.to_excel('BMI Igualas Medica Extra_duplicates_remove.xlsx', index=False)
    print("Excel file with duplicates removed successfully created: 'BMI Igualas Medica Extra_duplicates_remove.xlsx'")

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
# Keep only valid emails
valid_emails = fNamelName[fNamelName['E-MAIL'].apply(is_valid_email)]

# Save cleaned dataset
valid_emails.to_excel('Mejoras en la cobertura de accidentes personales MAR26.xlsx', index=False)
print("Excel file with only valid emails successfully created: 'Mejoras en la cobertura de accidentes personales MAR26.xlsx'")

# Reload cleaned dataset
mailChimp_format = pd.read_excel('Mejoras en la cobertura de accidentes personales MAR26.xlsx', sheet_name='Sheet1')

df = mailChimp_format[['E-MAIL', 'First Name', 'Last Name']].dropna(subset=['E-MAIL'])

if "E-MAIL" in df.columns:
    df = df.drop_duplicates(subset=["E-MAIL"], keep="first")

df.to_csv('Mejoras en la cobertura de accidentes personales MAR26.csv', index=False)
print("CSV file for MailChimp format successfully created: 'Mejoras en la cobertura de accidentes personales MAR26.csv'")