import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Open file explorer to select a file
Tk().withdraw()  # Hide the root window
file_path = askopenfilename(title="Select MDF File", filetypes=[("Excel files", "*.xlsx *.xls")])

if file_path:
    # Read the selected Excel file
    mdfDF = pd.read_excel(file_path)
    print("Original Columns:", mdfDF.columns.tolist())
else:
    print("No file selected.")
    exit()

# Read the beneficiaries Excel file
benDF = pd.read_excel('beneficiaries.xlsx')
benDfCols = benDF.columns.tolist()

#Read tenants
tenDF = pd.read_excel('tenants.xlsx')
tenDFCols = tenDF.columns.tolist()

#Read properties
propDF = pd.read_excel('properties.xlsx')
propDFCols = propDF.columns.tolist()


# Define the column mapping
tenants_mapping = {
    'Tenants(s) Name': 'Business name',
    'Tenants E-Mail Address': 'E-mail address',
    'Additional Landlord E-Mail CC': 'E-mail CC',
    'Landlord Mobile Number': 'Mobile number',
    'Landlord Additional Phone Number': 'Phone number',
    'Landlord Contact Address 1': 'Address 1',
    'Landlord Contact Address 2': 'Address 2',
    'Landlord Contact Address 3': 'Address 3',
    'Landlord Contact City': 'City',
    'Landlord Contact County': 'County',
    'Landlord Contact Postcode': 'Postcode',
    'Landlord Country': 'Country',
    'Landlord Account Name': 'Account name',
    'Landlord Sort Code (6 Digits No Spaces/Dashes)': 'Sort code',
    'Landlord Account Number (8 Digits No Spaces/Dashes)': 'Account number',
    'Landlord Country': 'Bank country',
}
# Define the column mapping
beneficiary_mapping = {
    'Landlord(s) Name': 'Business name',
    'Landlord E-Mail Address': 'E-mail address',
    'Additional Landlord E-Mail CC': 'E-mail CC',
    'Landlord Mobile Number': 'Mobile number',
    'Landlord Additional Phone Number': 'Phone number',
    'Landlord Contact Address 1': 'Address 1',
    'Landlord Contact Address 2': 'Address 2',
    'Landlord Contact Address 3': 'Address 3',
    'Landlord Contact City': 'City',
    'Landlord Contact County': 'County',
    'Landlord Contact Postcode': 'Postcode',
    'Landlord Country': 'Country',
    'Landlord Account Name': 'Account name',
    'Landlord Sort Code (6 Digits No Spaces/Dashes)': 'Sort code',
    'Landlord Account Number (8 Digits No Spaces/Dashes)': 'Account number',
    'Landlord Country': 'Bank country',
}

def process_beneficiaries(mdfDF, newDF, col_mapping, output_file_name):
    # Create a new DataFrame to match benDF's structure
    mapped_data = {ben_col: mdfDF[mdf_col] for mdf_col, ben_col in beneficiary_mapping.items()}
    # Add any missing columns in benDF with default values
    for col in newDF.columns:
        if col not in mapped_data:
            mapped_data[col] = None  # Add missing columns with default values
    # Convert the mapped_data dictionary into a DataFrame
    mdf_to_benDF = pd.DataFrame(mapped_data)
    # Append the rows of mdfDF to benDF
    newDF = pd.concat([newDF, mdf_to_benDF], ignore_index=True)
    #remove duplicates
    finalDF = newDF.drop_duplicates(subset=['Business name', 'Account number', 'Sort code'], keep='first')
    #write to an excel file
    finalDF.to_excel(output_file_name, index=False)

process_beneficiaries(mdfDF, benDF, beneficiary_mapping, 'beneficiariesDf_output.xlsx')


