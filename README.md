# Project Title:
LATAM_ITP_SUNET-Vendor-Notifications

# Description:
In order for SAP to record the invoice, the vendor must first create the invoice in the government portal known as SUNAT and then submit it to a designated shared mailbox. This is how the vendor issues an invoice against Anglo American.

However, at the moment, the vendor occasionally submits the invoice on the government site but fails to forward it to the shared mailbox. The Peru team will reject the invoice and require the vendor to provide a new one if they did not receive it within seven days or if it is inaccurate for any other reason.

Presently, when the team performs their checks and discovers that the invoice was captured on the portal but not on SAP, they reject that invoice, and the vendor will have to wait 37 days to be paid. This is because the vendors capture the invoice on the SUNAT portal but do not send it to the mailbox to be captured on SAP.

The request was to develop an automation that would use business rules to perform daily control checks between the SUNAT and SAP. In the event that an invoice is missing from both SUNAT and SAP, the automation would need to notify the vendor to resend the invoice so they can be paid on time.

This process was automated, which enhanced accuracy, controls, and saved time.

# Prerequisites and Dependencies:
To proceed with the task, we will need the following reports:
1. Three SUNAT Reports.
2. SAP Report.
3. Vendors Emails List Report.
4. Vendor Codes Report.

Please note that due to data sensitivity, I cannot upload these reports. However, if you wish to see the automation running and the output files, you can contact me so that we can proceed with the task promptly.

# Code Explanation:
This script is designed to automate the process of comparing two reports, SUNAT and SAP VIM, and identifying any discrepancies between them, specifically missing invoices in the SAP VIM report. The script performs the following steps:

1. Checks the version of the script and alerts the user if it is outdated.
2. Sets up logging and execution status variables.
3. Lists and merges XLSX files from a specified subfolder.
4. Loads the SAP VIM report, SUNAT report, and vendor codes spreadsheet into pandas DataFrames.
5. Groups vendor codes by RUC number and identifies RUC numbers with multiple vendor codes.
6. Filters out rows with RUC numbers that have multiple vendor codes from the sunat_report and vendor_codes_sap DataFrames.
7. Adds the SAP vendor code to the SUNAT report and updates column names.
8. Processes the 'Referencia' column of the sap_vim_report DataFrame by extracting and converting numeric parts of the string values, and replacing non-numeric or invalid values with None.
9. Concatenates the 'Proveedor' and 'Referencia' columns of the sap_vim_report DataFrame, handling missing values appropriately.
10. Merges the vendors_email_list DataFrame with the sunat_report DataFrame to add the 'Vendor Email' column.
11. Removes duplicate rows based on the 'Concatenated Field' column in the sunat_report DataFrame.
12. Saves the updated SUNAT report and SAP VIM report as Excel files.
13. Compares the SUNAT report with the SAP VIM report by merging the DataFrames on the specified columns.
14. Identifies invoices in the SUNAT report but not in the SAP VIM report, and invoices in both reports.
15. Saves the identified missing invoices as an Excel file and creates email notifications for missing invoices with and without email addresses.

The script uses several Python libraries, including pandas, os, glob, datetime, and win32com.client for handling files, data manipulation, and email notifications. It also includes error handling and logging functionalities to ensure smooth execution and traceability.

# Outcome Achieved:
1. Significant reduction in rejection rates and payment delays.
2. Vendors were promptly notified of missing invoices, allowing them to resubmit documents in a timely manner.
3. This automation led to improved invoice processing efficiency and increased vendor compliance, ultimately resulting in smoother operations and fewer payment postponements.
4. Approximately 554 notification reminders have been sent to vendors so far.
