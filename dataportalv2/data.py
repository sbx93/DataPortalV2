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

# Read the beneficiaries,tenants and properties files for structure
benDF = pd.read_excel('beneficiaries.xlsx')
tenDF = pd.read_excel('tenants.xlsx')
propDF = pd.read_excel('properties.xlsx')



# Define the column mapping
tenants_mapping = {
    'Tenant(s) Name': 'Business name',
    'Tenant E-Mail Address': 'E-mail address',
    'Additional Tenant E-Mail CC': 'E-mail CC',
    'Tenant Mobile Number': 'Mobile number',
    'Tenant Additional Phone Number': 'Phone number',
    'Tenant Address 1': 'Address 1',
    'Tenant Address 2': 'Address 2',
    'Tenant Address 3': 'Address 3',
    'Tenant City': 'City',
    'Tenant County': 'County',
    'Tenant Postcode': 'Postcode',
    'Tenant Country': 'Country',
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

property_mapping = {
    'Property Name': 'Property name',
    'Agent/Branch': 'Agent',
    'Service Level': 'Service level',
    'Rent Amount': 'Monthly payment required',
    'Agreed Float Amount': 'Prop. acc. minimum balance',
    'Notes': 'Comment',
    'Address 1': 'Address 1',
    'Address 2': 'Address 2',
    'Address 3': 'Address 3',
    'City': 'City',
    'County': 'County',
    'Postcode': 'Postcode',
    'Country': 'Country'
}


def process_data(mdfDF, newDF, col_mapping, output_file_name):
    # Create a new DataFrame to match benDF's structure
    mapped_data = {ben_col: mdfDF[mdf_col] for mdf_col, ben_col in col_mapping.items()}
    # Add any missing columns in benDF with default values
    for col in newDF.columns:
        if col not in mapped_data:
            mapped_data[col] = None  # Add missing columns with default values
    # Convert the mapped_data dictionary into a DataFrame
    mdf_to_benDF = pd.DataFrame(mapped_data)
    # Append the rows of mdfDF to benDF
    finalDF = pd.concat([newDF, mdf_to_benDF], ignore_index=True)


    #Filling columns with N and Y, checking for beneficiary output file to add payment advice

    if(output_file_name == 'beneficiariesDf_output.xlsx'):
        finalDF = finalDF.drop_duplicates(subset=['Business name', 'Account number', 'Sort code'], keep='first')
        finalDF[['Notify e-mail','PaymentAdvice']] = 'Y'

    if(output_file_name == 'tenantsDf_output.xlsx'):
        finalDF[['Notify e-mail']] = 'Y'

    if output_file_name == 'tenantsDf_output.xlsx' or output_file_name == 'beneficiariesDf_output.xlsx':
        finalDF['Notify text'] = 'N'

    if(output_file_name == 'propertiesDf_output.xlsx'):
        finalDF = finalDF.drop_duplicates(subset=['Property name'], keep='first')
        finalDF = finalDF.loc[:, ~finalDF.columns.str.contains('^Unnamed')]
        finalDF = finalDF.dropna(how='all')

    #Output final excel file
    finalDF.to_excel(output_file_name, index=False)


process_data(mdfDF, benDF, beneficiary_mapping, 'beneficiariesDf_output.xlsx')
process_data(mdfDF, tenDF, tenants_mapping, 'tenantsDf_output.xlsx')
process_data(mdfDF, propDF, property_mapping, 'propertiesDf_output.xlsx')

