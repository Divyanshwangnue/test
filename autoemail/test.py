import email
import imaplib
import os
import re
import pandas as pd
from email.policy import default
from bs4 import BeautifulSoup

# Function to extract email content
def extract_email_content(mail, email_ids):
    contents = []
    for num in email_ids[0].split():
        status, data = mail.fetch(num, '(RFC822)')
        if status != 'OK':
            print(f"Failed to fetch email with id {num}")
            continue
        msg = email.message_from_bytes(data[0][1], policy=default)
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                contents.append({'type': 'plain', 'content': content})
                print(f"Extracted plain text content: {content}")
            elif part.get_content_type() == 'text/html':
                content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                contents.append({'type': 'html', 'content': content})
                print(f"Extracted HTML content: {content}")
    return contents

# Function to parse FTP details from the HTML content
def parse_ftp_details_from_html(html_content):
    ftp_details = []
    soup = BeautifulSoup(html_content, 'html.parser')
    rows = soup.find_all('tr')
    login_id = None
    password = None

    for row in rows:
        cells = row.find_all('td')
        for cell in cells:
            if 'Login ID' in cell.text:
                login_id = cell.find_next_sibling('td').text.strip()
            if 'Password' in cell.text:
                password = cell.find_next_sibling('td').text.strip()
            if login_id and password:
                ftp_details.append({
                    'Login ID': login_id,
                    'Password': password,
                    'Date': pd.Timestamp.now()
                })
                login_id = None
                password = None
    return ftp_details

# Function to parse FTP details from plain text content
def parse_ftp_details_from_plain(plain_content):
    ftp_details = []
    regex = r"Login ID\s*:\s*(\S+)\s*Password\s*:\s*(\S+)"
    matches = re.finditer(regex, plain_content, re.DOTALL)
    for match in matches:
        ftp_details.append({
            'Login ID': match.group(1),
            'Password': match.group(2),
            'Date': pd.Timestamp.now()
        })
    return ftp_details

# Function to save details to Excel
def save_to_excel(details, filename):
    df = pd.DataFrame(details)
    if not df.empty:
        df['Date'] = pd.to_datetime(df['Date'])
        df.sort_values(by='Date', inplace=True)
        df.to_excel(filename, index=False)
    else:
        print("No details to save")

# Main function to automate the process
def automate_ftp_extraction(username, password, imap_server, output_file):
    # Connect to the server and login
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(username, password)

    # Select the mailbox you want to check
    mail.select("inbox")

    # Search for specific mails
    status, email_ids = mail.search(None, '(SUBJECT "ftp names and passwords")')
    if status != 'OK' or not email_ids[0]:
        print("No emails found")
        return

    # Print out subjects of emails found
    for num in email_ids[0].split():
        status, data = mail.fetch(num, '(BODY[HEADER.FIELDS (SUBJECT)])')
        if status == 'OK':
            print(f"Email ID: {num} Subject: {data[0][1].decode()}")

    # Extract email content
    contents = extract_email_content(mail, email_ids)
    
    # Close the connection and logout
    mail.close()
    mail.logout()

    # Parse FTP details and save to Excel
    ftp_details = []
    for content in contents:
        if content['type'] == 'html':
            details = parse_ftp_details_from_html(content['content'])
        else:
            details = parse_ftp_details_from_plain(content['content'])
        print(f"Parsed FTP details: {details}")
        ftp_details.extend(details)
    save_to_excel(ftp_details, output_file)
    print(f"Saved FTP details to {output_file}")

# Configuration
username = "project020304@outlook.com"
password = "divyansh2004"
imap_server = "imap.outlook.com"
output_file = "ftp_details.xlsx"

# Run the automation
automate_ftp_extraction(username, password, imap_server, output_file)
