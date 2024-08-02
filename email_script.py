import os
import sys
import json

from dotenv import load_dotenv
load_dotenv('.env', override=True)

import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import gspread
from oauth2client.service_account import ServiceAccountCredentials

def authenticate_google_sheets(json_keyfile_name, scopes):
    # Authenticate using the service account JSON file and the specified scopes
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scopes)
    client = gspread.authorize(credentials)
    return client

def get_sheet_data(client, spreadsheet_name, sheet_name):
    # Open the spreadsheet and select the sheet by name
    sheet = client.open(spreadsheet_name).worksheet(sheet_name)
    # Get all the data from the sheet as a list of dictionaries
    data = sheet.get_all_records(head=1)
    return data

def extract_sheet_data():
    # Define the scope and the path to the service account key file
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    json_keyfile_name = "service_account.json"

    # Authenticate and get the client
    client = authenticate_google_sheets(json_keyfile_name, scopes)

    # Define the spreadsheet and sheet names
    spreadsheet_name = "RTPA Member Roster"
    sheet_name = "Form Responses 1"

    # Get the data from the sheet
    data = get_sheet_data(client, spreadsheet_name, sheet_name)

    # Print the data
    for row in range(len(data)):
        print(row)

def load_html_file(file_path: str) -> str:
    with open(file_path, 'r') as file:
        return file.read()

def send_acceptance(applicant_email: str, html_body: str, server: smtplib.SMTP):
    
    msg = MIMEMultipart()
    msg['From'] = os.environ.get('OUTLOOK_EMAIL')
    msg['To'] = applicant_email
    msg['Subject'] = "Membership Application Update - RTPA"
    
    msg.attach(MIMEText(html_body, 'html'))
    text = msg.as_string()
    
    # Send email
    server.sendmail(msg)

def sendEmails(email_addresses):
    # Create email server object
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(os.environ.get('OUTLOOK_EMAIL'), os.environ.get('OUTLOOK_PASSWORD'))
    
    # Load our email HTML template
    html_template = load_html_file('acceptance_email.html')
    
    # Send email to each applicant
    for applicant in email_addresses:
        send_acceptance(applicant['email'], html_template, server)
        print(f"Email sent to {applicant['name']}")
    
    # Close the server connection
    server.quit()

def main():
    # Define the scope and the path to the service account key file
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    json_keyfile_name = "service_account.json"

    # Authenticate and get the Google Sheets client
    client = authenticate_google_sheets(json_keyfile_name, scopes)
    data = get_sheet_data(client, "RTPA Member Roster", "Form Responses 1")
    
    records = []
    # Print the data
    with open('data.json', 'w') as f:
        for rownum in range(len(data)):
            if data[rownum]["Current GSU GPA"] > 2.0 and rownum["Approved?"] == "":
                records.append((data[rownum], rownum))
                json.dump(data, f, indent=4)
    # for row in sheet:
    #     print(row)
    # data = extract_sheet_data()
    
    # Send emails to each applicant
    # sendEmails(email_addresses)

if __name__ == '__main__':
    # Load email addresses from file
    sys.exit(main())