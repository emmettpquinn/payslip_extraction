# Payslip Extraction

This project automates the extraction of key payroll data from PDF payslips received via email and appends the results to a Google Spreadsheet.

## Features

- Connects to an IMAP email account and scans for new payslip emails
- Downloads and decrypts PDF attachments
- Extracts payroll values (gross pay, net pay, tax, PRSI, USC, payment date) using Perplexity AI
- Appends extracted data to a specified Google Spreadsheet
- Tracks processed emails to avoid duplicates

## Setup

1. **Clone the repository.**
2. **Create a `.env` file in the `app/` directory.**  
   Use `.env.template` as a reference and fill in your credentials and API keys.
3. **Install dependencies:**  
   Run `pip install -r requirements.txt` in the `app/` directory.
4. **Add your Google service account credentials:**  
   Place your `credentials.json` file in the `app/` directory.

## Environment Variables

Copy `.env.template` to `.env` and fill in the following values:

```
EMAIL_ACCOUNT=
EMAIL_PASSWORD=
IMAP_SERVER=
PDF_PASSWORD=
PERPLEXITY_API_KEY=
SPREADSHEET_URL=
```

## Usage

Run the main script:

```powershell
python app/main.py
```

The script will process new payslip emails and update your Google Spreadsheet automatically.

## Security

- **Do not commit your `.env` or `credentials.json` files.**  
  Use `.env.template` for sharing configuration requirements.

## Troubleshooting

- Ensure all required Python packages are installed. If you see import errors, run:
  ```powershell
  pip install -r requirements.txt
  ```
- Make sure your Google service account has access to the target spreadsheet.
- Check that your IMAP credentials and PDF password are correct.

## License

MIT License
