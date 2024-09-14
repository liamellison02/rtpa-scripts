import os
import sys
import json
from typing import List, Tuple

from dotenv import load_dotenv
load_dotenv('.env', override=True)

import win32com.client as win32client

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


def load_html_file(file_path: str) -> str:
    with open(file_path, 'r') as file:
        return file.read()


def send_acceptance(applicant_email: str, html_body: str, outlook):
    
    # Create new email message object
    msg = outlook.CreateItem(0)
    msg.To = applicant_email
    msg.Subject = "Membership Application Update - RTPA"
    
    # Add acceptance HTML template to body
    msg.HTMLBody = html_body
    
    # Send email
    msg.Save()
    msg.Send()


def send_error_report(invalidApplicants: List[Tuple], outlook):

    # Create new email message object
    msg = outlook.CreateItem(0)
    msg.To = os.environ.get('RTPA_EMAIL')
    msg.Subject = "WARNING: Invalid Applicants found"

    # Add applicants to message body
    msg.Body = "The following applicant records were flagged as invalid for determining qualification for acceptance. Please review: \n"
    msg.Body += str(invalidApplicants)

    # Send email
    msg.Save()
    msg.Send()



def send_emails(qualifiedApplicants: List[Tuple], invalidApplicants: List[Tuple], unqualifiedApplicants: List[Tuple]):
    outlook = win32client.Dispatch("Outlook.Application")
    
    # Load our email HTML template
    html_template = load_html_file('acceptance_email.html')
    
    # Send email to each applicant
    for applicant in qualifiedApplicants:
        send_acceptance(applicant[0]["GSU Email"], html_template, outlook)
        print(f'Sent email to: {applicant[0]["GSU Email"]}')
    
    # Send error report if applicable.
    # if len(invalidApplicants) > 0:
    #     send_error_report(invalidApplicants, outlook)


def main():
    # Define the scope and the path to the service account key file
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    json_keyfile_name = "service_account.json"

    # Authenticate and get the Google Sheets client
    client = authenticate_google_sheets(json_keyfile_name, scopes)
    data = get_sheet_data(client, "RTPA Member Roster", "Form Responses 1")
    
    qualApps, unqualApps, invalidApps = [], [], [] # arrays for qualified, unqualified, and invalid applications

    # Sort all applications by qualification
    for rownum in range(len(data)):
        if data[rownum]["Reviewed?"] == "":
            # Try-catch block provides input validation for non-numeric inputs for GPA field.
            try:
                gpa = float(data[rownum]["Current GSU GPA"])
                if gpa >= 2.0:  # If GPA qualifies, add as qualified applicant
                    qualApps.append([data[rownum], rownum+2])
                else:
                 unqualApps.append([data[rownum], rownum+2])  # Otherwise, add to unqualified array
            
            except ValueError as e:
                # Invalid input for GPA will be reported via email to EC.
                invalidApps.append([e.args[0], data[rownum], rownum+2])

    # Export records to JSON file.
    with open('data.json', 'w') as f:
        json.dump({"qualApps": qualApps, "unqualApps": unqualApps, "invalidApps": invalidApps}, f, indent=4)
    
    # Send emails to each applicant
    send_emails(qualApps, invalidApps, unqualApps)


if __name__ == '__main__':
    # Load email addresses from file
    sys.exit(main())
