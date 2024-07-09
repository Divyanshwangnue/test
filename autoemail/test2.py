import imaplib
import email
from email.policy import default
from bs4 import BeautifulSoup
import pandas as pd
from tqdm import tqdm

def fetch_emails(username, password, imap_server):
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(username, password)
    mail.select('inbox')

    status, data = mail.search(None, 'ALL')
    mail_ids = data[0].split()

    emails = []
    for i in tqdm(mail_ids, desc="Fetching emails", unit="email"):
        status, data = mail.fetch(i, '(RFC822)')
        emails.append(data[0][1])

    mail.logout()
    return emails

def parse_ftp_details_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    account_info_table = soup.find('table', {'class': 'MsoNormalTable'})
    if not account_info_table:
        return None
    
    rows = account_info_table.find_all('tr')
    ftp_details = {}
    
    for row in rows:
        cells = row.find_all('td')
        if len(cells) == 3:
            label = cells[0].get_text(strip=True)
            key = cells[1].get_text(strip=True)
            value = cells[2].get_text(strip=True)
            if key and value:
                ftp_details[key] = value
    
    return ftp_details if ftp_details else None

def extract_ftp_details_from_emails(emails):
    all_details = []
    for raw_email in tqdm(emails, desc="Extracting FTP details", unit="email"):
        msg = email.message_from_bytes(raw_email, policy=default)
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                html_content = part.get_payload(decode=True)
                try:
                    html_content = html_content.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        html_content = html_content.decode('latin-1')
                    except UnicodeDecodeError:
                        continue

                details = parse_ftp_details_from_html(html_content)
                if details:
                    all_details.append(details)
    return all_details

def save_to_excel(details_list, output_file):
    df = pd.DataFrame(details_list)
    df.to_excel(output_file, index=False)
    print(f"Details saved to {output_file}")

def automate_ftp_extraction(username, password, imap_server, output_file):
    emails = fetch_emails(username, password, imap_server)
    details = extract_ftp_details_from_emails(emails)
    save_to_excel(details, output_file)

# Parameters
username = "project020304@outlook.com"
password = "divyansh2004"
imap_server = "imap.outlook.com"
output_file = "ftp_details.xlsx"

# Execute the function
automate_ftp_extraction(username, password, imap_server, output_file)
