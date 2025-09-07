Payslip Extraction
This project automates the extraction of key payroll data from PDF payslips received via email and appends the results to a Google Spreadsheet.

Features
Connects to an IMAP email account and scans for new payslip emails.
Downloads and decrypts PDF attachments.
Extracts payroll values (gross pay, net pay, tax, PRSI, USC, payment date) using Perplexity AI.
Appends extracted data to a specified Google Spreadsheet.
Tracks processed emails to avoid duplicates.
Setup
Clone the repository.
Create a .env file in the app directory.
Use .env.template as a reference and fill in your credentials and API keys.
Install dependencies:
Run pip install -r requirements.txt in the app directory.
Add your Google service account credentials:
Place your credentials.json file in the app directory.
Usage
Run the main script:

The script will process new payslip emails and update your Google Spreadsheet automatically.

Security
Do not commit your .env or credentials.json files.
Use .env.template for sharing configuration requirements.
License
MIT License
