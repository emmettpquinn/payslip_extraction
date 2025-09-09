import os
import json
import logging
from dotenv import load_dotenv
from datetime import datetime, timedelta
from imapclient import IMAPClient
import pyzmail
from PyPDF2 import PdfReader
import gspread
from google.oauth2.service_account import Credentials
import requests
import re
import time

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler("payslip.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- Env Loading ---
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
IMAP_SERVER = os.getenv('IMAP_SERVER')
EMAIL_ACCOUNT = os.getenv('EMAIL_ACCOUNT')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
PDF_PASSWORD = os.getenv('PDF_PASSWORD')
GOOGLE_CREDENTIALS_JSON = os.path.join(os.path.dirname(__file__), 'credentials.json')
SPREADSHEET_URL = os.getenv('SPREADSHEET_URL')
PROCESSED_UID_FILE = os.path.join(os.path.dirname(__file__), "processed_emails.json")
PERPLEXITY_API_KEY = os.getenv('PERPLEXITY_API_KEY')


# --- UID Tracker ---
def load_processed_uids():
    if not os.path.exists(PROCESSED_UID_FILE):
        return set()
    try:
        with open(PROCESSED_UID_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                return set(str(uid) for uid in data)
            else:
                logger.error("Malformed processed_emails.json, resetting...")
                with open(PROCESSED_UID_FILE, 'w', encoding='utf-8') as wf:
                    json.dump([], wf)
                return set()
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        logger.exception("Corrupted processed_emails.json file, resetting...")
        with open(PROCESSED_UID_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f)
        return set()


def save_processed_uid(uid):
    uids = []
    if os.path.exists(PROCESSED_UID_FILE):
        try:
            with open(PROCESSED_UID_FILE, 'r', encoding='utf-8') as f:
                uids = json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            logger.exception("Error reading processed_emails.json file.")
            uids = []
    if str(uid) not in map(str, uids):
        uids.append(str(uid))
        try:
            with open(PROCESSED_UID_FILE, 'w', encoding='utf-8') as f:
                json.dump(uids, f)
        except Exception as e:
            logger.exception("Failed to save processed UID.")


# --- PDF Extraction ---
def extract_text_from_pdf(pdf_path, password):
    logger.info(f"Attempting to decrypt PDF: {pdf_path}")
    try:
        with open(pdf_path, "rb") as infile:
            reader = PdfReader(infile)
            if reader.is_encrypted:
                logger.info("PDF is encrypted, attempting to decrypt...")
                # Only use provided password
                if not reader.decrypt(password):
                    logger.error("Failed to decrypt PDF with provided password.")
                    return None
                logger.info("PDF decrypted successfully.")
            else:
                logger.info("PDF is not encrypted.")
            text = []
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)
            return "\n".join(text)
    except Exception as e:
        logger.exception("Error processing PDF")
        return None


# --- Perplexity Extraction ---
def extract_values_with_perplexity(text):
    api_key = PERPLEXITY_API_KEY
    if not api_key:
        logger.error("Perplexity API key not set.")
        return {}
    prompt = (
        "Extract these values from the following Irish payslip text and return as a JSON object with these keys: "
        "gross_pay, net_pay, tax, prsi, usc, payment_date, payer. "
        "Here is the text:\n" + text
    )
    url = "https://api.perplexity.ai/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    data = {
        "model": "sonar",
        "messages": [
            {"role": "system", "content": "You are a data extraction assistant."},
            {"role": "user", "content": prompt}
        ]
    }
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            res = response.json()
            answer = res.get("choices", [{}])[0].get("message", {}).get("content", "")
            match = re.search(r'({.*})', answer, re.DOTALL)
            if match:
                try:
                    result = json.loads(match.group(1))
                    logger.info("Perplexity extraction successful.")
                    return result
                except Exception as e:
                    logger.exception("Perplexity JSON parsing error.")
            else:
                logger.error("No JSON found in Perplexity response.")
        else:
            logger.error(f"Perplexity API error: {response.status_code}")
    except Exception as e:
        logger.exception("Perplexity API call failed.")
    return {}


# --- Google Sheets Append ---
def append_to_google_sheet(data_dict, email_id, email_date):
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_JSON, scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(SPREADSHEET_URL)
        sheet = spreadsheet.worksheet('Payslip Email Finder')
        row = [
            email_id,
            email_date,
            data_dict.get('payment_date', ''),
            data_dict.get('payer', ''),
            data_dict.get('gross_pay', ''),
            data_dict.get('tax', ''),
            data_dict.get('prsi', ''),
            data_dict.get('usc', ''),
            data_dict.get('net_pay', ''),
            ''  # Processed Indicator
        ]
        sheet.append_row(row)
        logger.info(f"GoogleSheet: Row appended for email UID {email_id}.")
    except Exception as e:
        logger.exception(f"GoogleSheet: Error writing row for email UID {email_id}.")


def process_email(mail, uid):
    raw_message = mail.fetch([uid], ['BODY[]', 'ENVELOPE'])
    message = pyzmail.PyzMessage.factory(raw_message[uid][b'BODY[]'])
    envelope = raw_message[uid][b'ENVELOPE']
    email_id = str(uid)
    email_date = envelope.date.strftime('%d/%m/%Y %H:%M:%S') if envelope.date else ''
    perplexity_status = None
    google_sheet_status = None
    for part in message.mailparts:
        if part.filename and part.filename.lower().endswith('.pdf'):
            work_dir = "/app" if os.path.exists("/app") else "."
            file_path = f"{work_dir}/{part.filename}"
            try:
                with open(file_path, 'wb') as f:
                    f.write(part.get_payload())
                text = extract_text_from_pdf(file_path, PDF_PASSWORD)
                if text:
                    data = extract_values_with_perplexity(text)
                    perplexity_status = "success" if data else "error"
                    append_to_google_sheet(data if data else {}, email_id, email_date)
                    google_sheet_status = "success"
                else:
                    perplexity_status = "error"
                    google_sheet_status = "error"
            except Exception as e:
                logger.exception(f"Error processing email UID {email_id}")
            finally:
                try:
                    os.remove(file_path)
                except Exception:
                    logger.warning(f"Could not delete file {file_path}")
    return perplexity_status, google_sheet_status


def main_loop():
    logger.info("Starting PDF decryption app...")
    while True:
        now = datetime.now()
        today = now.weekday()
        if today > 4:
            days_until_monday = 7 - today
            next_run = (now + timedelta(days=days_until_monday)).replace(hour=0, minute=0, second=0, microsecond=0)
            sleep_seconds = (next_run - now).total_seconds()
            logger.info(f"Weekend. Sleeping until next Monday ({next_run}) for {int(sleep_seconds)} seconds.")
            time.sleep(sleep_seconds)
            continue
        EMAIL_ACCOUNT_ENV = EMAIL_ACCOUNT
        EMAIL_PASSWORD_ENV = EMAIL_PASSWORD
        PDF_PASSWORD_ENV = PDF_PASSWORD
        IMAP_SERVER_ENV = IMAP_SERVER
        FOLDERS = ['INBOX', '1. Payslips']
        if not all([EMAIL_ACCOUNT_ENV, EMAIL_PASSWORD_ENV, PDF_PASSWORD_ENV]):
            logger.error("ERROR: Missing required configuration.")
            return
        logger.info(f"Email: {EMAIL_ACCOUNT_ENV}")
        processed_uids = load_processed_uids()
        logger.info(f"Loaded {len(processed_uids)} processed UIDs.")
        since_date = (datetime.now() - timedelta(days=30)).strftime('%d-%b-%Y')
        logger.info(f"Searching for emails since: {since_date}")
        try:
            with IMAPClient(IMAP_SERVER_ENV) as server:
                server.login(EMAIL_ACCOUNT_ENV, EMAIL_PASSWORD_ENV)
                for folder in FOLDERS:
                    try:
                        server.select_folder(folder)
                        search_criteria = [u'FROM', 'payslips@brightpay.ie', u'SINCE', since_date]
                        uids = server.search(search_criteria)
                        new_uids = [uid for uid in uids if str(uid) not in processed_uids]
                        logger.info(f"EmailSearch: Folder: {folder} | Total: {len(uids)} | New: {len(new_uids)}")
                        total_new_emails = 0
                        perplexity_success = 0
                        perplexity_error = 0
                        gs_success = 0
                        gs_error = 0
                        for uid in new_uids:
                            try:
                                perplexity_status, google_sheet_status = process_email(server, uid)
                                save_processed_uid(uid)
                                processed_uids.add(str(uid))
                                total_new_emails += 1
                                if perplexity_status == "success":
                                    perplexity_success += 1
                                else:
                                    perplexity_error += 1
                                if google_sheet_status == "success":
                                    gs_success += 1
                                else:
                                    gs_error += 1
                            except Exception as e:
                                logger.exception(f"Error processing UID {uid}")
                        logger.info(f"Summary: {folder}: Processed {total_new_emails} | Perplexity success: {perplexity_success}, error: {perplexity_error} | GoogleSheet success: {gs_success}, error: {gs_error}")
                    except Exception as e:
                        logger.exception(f"EmailSearch: Error accessing folder {folder}.")
        except Exception as e:
            logger.exception("Error in main loop.")
        logger.info("Finished processing for today.")
        next_run = now + timedelta(days=1)
        while next_run.weekday() > 4:
            next_run += timedelta(days=1)
        next_run = next_run.replace(hour=0, minute=0, second=0, microsecond=0)
        sleep_seconds = (next_run - now).total_seconds()
        logger.info(f"Sleeping until next weekday ({next_run}) for {int(sleep_seconds)} seconds.")
        time.sleep(sleep_seconds)


if __name__ == '__main__':
    main_loop()
