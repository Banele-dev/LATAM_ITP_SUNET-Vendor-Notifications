import pandas as pd
import os
from datetime import date, datetime
import warnings
import glob
import sys
import win32com.client as win32

## Setting variables to check is this version matches with the GSS Automation Team's control
# application = "ITP_SUNET Vendor Notifications"
# version = "v01"
user_name = os.getlogin()
# path = f"C:/Users/{user_name}/Box/Automation Script Versions/versions.xlsx"
# df = pd.read_excel(path)
# filter_criteria = (df['app'] == application) & (df['versão'] == version)
# start_time = None
#
# if not filter_criteria.any():
#     input('Outdated app, talk to the automation team. Press ENTER to close the code \n')
#     quit()

# Disable openpyxl's UserWarning about default style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# Get today's date
today = date.today()
# Format today's date in the desired format
date_format = today.strftime("%Y%m%d")
# Define the path to the subfolder using today's date
subfolder_path = f"C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/{date_format}"
# subfolder_path = f"C:/Users/Public/Documents/LAT_ITP_SUNET Vendor Notifications/{date_format}"

# Get the path of the directory where the script is located
script_directory = os.path.dirname(os.path.abspath(__file__))

# Create the path for the LogControl folder
log_control_path = os.path.join(script_directory, 'LogControl')

# Check if the LogControl folder exists, if not, create it
if not os.path.exists(log_control_path):
    os.makedirs(log_control_path)

# Create the full path to the log file within the Log Control folder
log_file_name = f"ExecutionLog_{datetime.now().strftime('%d%m%Y%H%M')}" + ".txt"
log_file_path = os.path.join(log_control_path, log_file_name)

try:
    # Log the execution start time
    start_time = datetime.now().strftime('%d%m%Y%H%M')
    log_file = open(log_file_path, 'a')
    sys.stdout = log_file
    with open(log_file_path, 'a') as log_file:
        log_file.write(f"Execution started at: {start_time} \n")
except Exception as e:
    print(f"Error during execution log creation: {e}")

# Define execution status
SUCCESS = "Successful"
FAILED = "Failed"
SKIPPED = "Skipped"

# Get the list of files in the subfolder
try:
    files = os.listdir(subfolder_path)
    print(f"Found {len(files) - 1} XLSX files in the subfolder")
    print("Listing file names:")
    for file in files:
        if file.endswith(".xlsx"):
            print(file)
except Exception as e:
    print(f"Error while listing files in the subfolder: {e}")
    print(f"Listing files status: {FAILED}")
# Create an empty DataFrame to store the merged data
merged_data = pd.DataFrame()

# Process each file in the subfolder
for step, file in enumerate(files, start=1):
    print(f"\n Processing file {step} of {len(files)}...")
    try:
        if file.endswith(".xlsx"):
            # Construct the full path to the file
            file_path = os.path.join(subfolder_path, file)
            # Read the Excel file using pandas
            data = pd.read_excel(file_path)
            # Append the data to the merged DataFrame
            merged_data = pd.concat([merged_data, data])
    except Exception as e:
        print(f"Error while processing file {file}: {e}")
        print(f"File {step} execution status: {FAILED}")
    else:
        print(f"File {step} execution status: {SUCCESS}")

# Reset the index of the merged DataFrame
merged_data = merged_data.reset_index(drop=True)

# Write the merged data to a new Excel file and save the merged SUNAT reports
try:
    merged_data.to_excel(os.path.join(subfolder_path, 'Merged_data.xlsx'), index=False)
except Exception as e:
    print(f"Error while saving merged data: {e}")

# Directory where the SAP reports are stored.
directory_path = f"C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/{date_format}/SAP Reports"
# directory_path = f"C:/Users/Public/Documents/LAT_ITP_SUNET Vendor Notifications/{date_format}/SAP Reports"
# Get a list of all Excel files in the directory
excel_files = glob.glob(os.path.join(directory_path, '*.xlsx'))
# Sort the list of files by modification time (most recent first)
excel_files.sort(key=os.path.getmtime, reverse=True)


# Function to translate column names
def translate_column_names(df, translation_dict):
    # Rename the columns using the translation dictionary
    df.rename(columns=translation_dict, inplace=True)


# Define a translation dictionary for column names
column_translation_dict = {
    # Map the original column names to Spanish, Portuguese names
    'Document Id': 'ID de documento',
    'Company Code': 'Sociedad',
    'Document Number': 'Nº documento',
    'Fiscal Year': 'Ejercicio',
    'Document Status': 'Status del documento',
    'Exception Reason': 'Motivo exception',
    'Document Type': 'Clase de documento',
    'DP Document Type': 'Clase doc. PD',
    'Credit Memo': 'Abono',
    'Reference': 'Referencia',
    'Supplier': 'Proveedor',
    # 'Vendor': 'Proveedor',
    'Total Amt in Doc Curr': 'Im. total en mon. doc',
    'Doc Currency': 'Moneda de documento',
    'Amt in Report Currency': 'Importe moneda informe',
    'Re. Currency': 'Moneda local',
    'Posting date': 'Fecha contabilizac.',
    'Due Date': 'Fecha de vencimiento',
    'Related Object Key': 'Clave obj. relac.',
    'Days to Due': 'Días hasta vencim.',
    'Overdue': 'Atrasado',
    'Requisitioner Name': 'Nombre solicitante',
    'Vendor Name': 'Nombre del proveedor',
    'Purchasing Document': 'Documento compras',
    'Plant': 'Centro',
    'Purchasing Group': 'Grupo de compras',
    'Cycle Time': 'Duración del ciclo',
    'Document Date': 'Fecha de documento',
    'Exception Date': 'Fec. excepción',
    'Enter on': 'Introducir el',
    'Enter at': 'Introducir a las',
    'Start on': 'Iniciar el',
    'Start at': 'Iniciar a las',
    'End on': 'Finaliza el',
    'End at': 'Finalizar a las',
    'Update Date': 'Fecha de actualización',
    'Update Time': 'Hora actualización',
    'Reversal Doc#': 'Nº doc anulac',
    'Revsed F_Year': 'Ejercicio anulado',
    'Old Doc Num': 'Nº documen. anterior',
    'Old Company Code': 'Sociedad anterior',
    'Old Fiscal Year': 'Ejercicio anterior',
    'Creation Date': 'Fecha de creación',
    'Creation Time': 'Creado a las',
    'End on': 'Finaliza el',
    'End at': 'Finalizar a las',
    'Doc Currency': 'Moneda de documento',
    'Process Type': 'Tipo de proceso',
    'Rescan Reason': 'Motivo de reescaneo',
    'Delete Reason': 'Motivo del borrado',
    'Target System': 'Sist. destino',
    'Rescan Reason Code': 'Cód. mot. reescaneo',
    'Obsoleter Reason Code': 'Cód. motivo obsoleto',
    'IDoc number': 'Número IDOC',
    'OpenText User Id': 'ID obj. asig. usuario',
    'DP ID before restart': 'ID TratDoc ants rein'

}

# Check if there are any Excel files in the directory
if excel_files:
    # Read the most recently modified Excel file, if there are, we read the most recently modified one
    most_recent_file = excel_files[0]
    Sap_vim_report = pd.read_excel(most_recent_file, engine='openpyxl')
    Sap_vim_report.to_excel(f"C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/{date_format}/SAP Reports/SAP_VIM_Report.xlsx", index=False)
    # Sap_vim_report.to_excel(f"C:/Users/Public/Documents/LAT_ITP_SUNET Vendor Notifications/{date_format}/SAP Reports/SAP_VIM_Report.xlsx", index=False)
    print(f"Processing {most_recent_file}\n")
else:
    print("No Excel files found in the directory.")

# Load the SUNAT report, SAP Vim report and Vendor Codes spreadsheet into pandas DataFrames
sunat_report = pd.read_excel(os.path.join(subfolder_path, 'merged_data.xlsx'))
vendor_codes_sap = pd.read_excel(fr'C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/Dependencies/Base Peru - RUC Codigo Proveedores.xlsx')
# vendor_codes_sap = pd.read_excel(r'C:\Users\Public\Documents\LAT_ITP_SUNET Vendor Notifications\Dependencies\Base Peru - RUC Codigo Proveedores.xlsx')
sap_vim_report = pd.read_excel(f"C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/{date_format}/SAP Reports/SAP_VIM_Report.xlsx")
# sap_vim_report = pd.read_excel(f"C:/Users/Public/Documents/LAT_ITP_SUNET Vendor Notifications/{date_format}/SAP Reports/SAP_VIM_Report.xlsx")

# Group vendor codes by RUC number
grouped_vendor_codes = vendor_codes_sap.groupby('RUC')['Codigo'].agg(['unique'])

# Identify RUC numbers with multiple vendor codes
ruc_with_multiple_codes = grouped_vendor_codes[grouped_vendor_codes['unique'].apply(len) > 1].reset_index()

# Save the report of RUC numbers with multiple vendor codes to an Excel file
if not ruc_with_multiple_codes.empty:
    # Define the file path for the report
    report_file_path = os.path.join(subfolder_path, 'Ruc_with_multiple_codes_report.xlsx')
    # Save the report to Excel
    ruc_with_multiple_codes.to_excel(report_file_path, index=False)
    print(f"Report of RUC numbers with multiple vendor codes successfully saved to {report_file_path}")
else:
    print("No RUC numbers with multiple vendor codes found.")

# Filter out entire rows from sunat_report and vendor_codes_sap where the RUC number is in the list of RUC numbers with multiple codes.
ruc_with_multiple_codes = ruc_with_multiple_codes['RUC'].unique()
# Filter rows where the 'RUC' column is not in the list of RUC numbers with multiple codes
sunat_report = sunat_report[~sunat_report['Número  documento de identidad del emisor'].isin(ruc_with_multiple_codes)]
vendor_codes_sap_filtered = vendor_codes_sap[~vendor_codes_sap['RUC'].isin(ruc_with_multiple_codes)]

# Check for and remove duplicate values in the RUC of vendor_codes_sap_filtered
if vendor_codes_sap_filtered['RUC'].duplicated().any():
    vendor_codes_sap_filtered = vendor_codes_sap_filtered[~vendor_codes_sap_filtered['RUC'].duplicated()]


# Add the SAP vendor code to the SUNAT report
sunat_report['SAP Vendor Code'] = sunat_report['Número  documento de identidad del emisor'].map(
    vendor_codes_sap_filtered.set_index('RUC')['Codigo']
)

# Update the column name from 'Reference' to 'Referencia', 'Vendor' to 'Proveedor'
sap_vim_report.rename(columns={'Reference': 'Referencia'}, inplace=True)
sap_vim_report.rename(columns={'Supplier': 'Proveedor'}, inplace=True)


# This line of code processes each element in the 'Referencia' column of the sap_vim_report DataFrame, extracting and converting numeric parts of the string values, and replacing non-numeric or invalid values with None.
def extract_number(value):
    if pd.isna(value) or pd.isnull(value) or value == '':  # Check if value is NaN, null, or empty string
        return None
    parts = str(value).split('-')  # Convert value to string before splitting
    if len(parts) > 1 and parts[-1].isdigit():
        return int(parts[-1])
    return None


sap_vim_report['Referencia'] = sap_vim_report['Referencia'].apply(extract_number)
sunat_report['Concatenated Field'] = sunat_report['SAP Vendor Code'].astype(str) + sunat_report['Número  Correlativo de CP'].astype(str)
# this line of code effectively creates a new column 'concatenated field' in the sap_vim_report DataFrame by concatenating the 'Proveedor' column with the 'Referencia' column, handling missing values appropriately.
sap_vim_report['concatenated field'] = sap_vim_report['Proveedor'].fillna('').astype(str) + sap_vim_report['Referencia'].apply(lambda x: str(int(x)) if pd.notnull(x) else '').astype(str)

# Load Vendors email list spreadsheet
vendors_email_list = pd.read_excel(fr'C:/Users/{user_name}/Anglo American/GSS Automation Team - Automation - Bot Dependencies/LAT_AP_PERU ITP/Dependencies\Vendors_emails_list.xlsx')
# vendors_email_list = pd.read_excel(r'C:\Users\Public\Documents\LAT_ITP_SUNET Vendor Notifications\Dependencies\Vendors_emails_list.xlsx')
# Merge the 'vendors_email_list'  with 'sunat_report' to add the 'Vendor Email' column
sunat_report = sunat_report.merge(vendors_email_list, how='left', left_on='SAP Vendor Code', right_on='Proveedor')

# removing rows where the values in the 'Concatenated Field' column are duplicates, keeping only the first occurrence of each unique value in that column.
sunat_report = sunat_report.drop_duplicates(subset='Concatenated Field', keep='first')

# Define the file path for the updated SUNAT report
updated_sunat_report_path = os.path.join(subfolder_path, 'Updated_sunat_report.xlsx')
# Save the updated SUNAT report with SAP vendor codes
try:
    sunat_report.to_excel(updated_sunat_report_path, index=False)
    print(f"Report of updated SUNAT report successfully saved to: {updated_sunat_report_path}")
except Exception as e:
    print(f"Error occurred while saving the updated SUNAT report: {e}")


# Define the file path for the updated SAP report
updated_sap_vim_report_path = os.path.join(subfolder_path, 'Updated_sap_vim_report.xlsx')
# Save the updated SAP VIM report
try:
    sap_vim_report.to_excel(updated_sap_vim_report_path, index=False)
    print(f"Report of updated SAP VIM report successfully saved to: {updated_sap_vim_report_path}")
except Exception as e:
    print(f"Error occurred while saving the updated SAP VIM report: {e}")

# Compare SUNAT report with SAP VIM report and send notifications. Perform VLOOKUP by merging the DataFrames on the specified columns
merged_df = sunat_report.merge(sap_vim_report, how='left', left_on='Concatenated Field', right_on='concatenated field')

# Identify invoices in SUNAT report but not in SAP VIM report
missing_invoices = merged_df[merged_df['concatenated field'].isnull()]

# Identify invoices in SUNAT report and also in SAP VIM report
common_invoices = merged_df[merged_df['concatenated field'].notnull()]

# Exclude common invoices from the missing invoices list
missing_invoices = missing_invoices[~missing_invoices['concatenated field'].isin(common_invoices['concatenated field'])]

# Iterate through missing invoices
for index, row in missing_invoices.iterrows():
    Vendor_email = row['Dirección internet del responsable']
    invoice_number = row['Número  Correlativo de CP']
    vendor_code = row['SAP Vendor Code']
    social_reason_emisor = row['Razón social emisor']
    concatenated_value = row['Concatenated Field']
    invoice_date = row['Fecha de puesta a disposición']

# Create a new DataFrame to store the identified missing invoices
missing_invoices_data = pd.DataFrame({
    'Número  Correlativo de CP': missing_invoices['Número  Correlativo de CP'],
    'Dirección internet del responsable': missing_invoices['Dirección internet del responsable'],
    'Vendor Code': missing_invoices['SAP Vendor Code'],
    'Razón social emisor': missing_invoices['Razón social emisor'],
    'Vendor code + Invoice number': missing_invoices['Concatenated Field'],
    'Invoice_date': missing_invoices['Fecha de puesta a disposición']

})

# Define the file path for the missing invoices report
missing_invoices_report_path = os.path.join(subfolder_path, 'Missing_invoices_data.xlsx')
try:
    # Write the missing invoices data to a new Excel file and save it
    missing_invoices_data.to_excel(missing_invoices_report_path, index=False)
    print(f"Report of missing invoices successfully saved to: {missing_invoices_report_path}")
except Exception as e:
    print(f"Error during missing invoices report saving: {e}")

try:
    # Filter the missing invoices data for invoices without an email address
    missing_invoices_without_email = missing_invoices_data[missing_invoices_data['Dirección internet del responsable'].isnull()]
    # Define the file path for the missing invoices without email report
    missing_invoices_without_email_path = os.path.join(subfolder_path, 'Missing_invoices_without_email.xlsx')
    # Write the missing invoices without email data to a new Excel file and save it
    missing_invoices_without_email.to_excel(missing_invoices_without_email_path, index=False)
    print(f"Report of missing invoices without email successfully saved to: {missing_invoices_without_email_path}")
except Exception as e:
    print(f"Error during missing invoices without email report saving: {e}")

try:
    # Filter the missing invoices data for invoices with an email address
    missing_invoices_with_email = missing_invoices_data[missing_invoices_data['Dirección internet del responsable'].notnull()]
    # Define the file path for the missing invoices with email report
    missing_invoices_with_email_path = os.path.join(subfolder_path, 'Missing_invoices_with_email.xlsx')
    # Write the missing invoices with email data to a new Excel file and save it
    missing_invoices_with_email.to_excel(missing_invoices_with_email_path, index=False)
    print(f"Report of missing invoices with email successfully saved to: {missing_invoices_with_email_path}")
except Exception as e:
    print(f"Error during missing invoices with email report saving: {e}")


# Redirect sys.stdout back to the standard output stream
sys.stdout = sys.__stdout__

# Create an Outlook application object and Create a new email
outlook = win32.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)
email.Subject = 'Automation Team - Execution Log File'
email_body = f"""
<html>
<body>
<p> Dear Automation Team,</p>
<p> 'Please find attached the execution log file for the LAT_ITP_SUNET Vendor Notifications automation executed on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}' </p>
<p> Best regards, <br></p>
</html>
</body>
"""
email.HTMLBody = email_body
email.To = 'banele.madikane@angloamerican.com'

# Attach the log file
attachment = os.path.abspath(log_file_path)
email.Attachments.Add(attachment)
email.Send()



