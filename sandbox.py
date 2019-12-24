import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from drive_handler import DriveHandler
from spreadsheet_handler import SpreadsheetHandler


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

CREDS = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        CREDS = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not CREDS or not CREDS.valid:
    if CREDS and CREDS.expired and CREDS.refresh_token:
        CREDS.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        CREDS = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(CREDS, token)

SERVICE = build('sheets', 'v4', credentials=CREDS)

"""
    Creating a Spreadsheet
"""
# spreadsheet = {
#     'properties': {
#         'title': 'expenses_tracker'
#     }
# }
# spreadsheet = SERVICE.spreadsheets().create(body=spreadsheet,
#                                             fields='spreadsheetId').execute()

# print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))

spreadsheet_hdl = SpreadsheetHandler(CREDS)

print(spreadsheet_hdl.create_spreadsheet())

a, b, c, d, e, f = input("Insira 6 valore: ").split()

data = [[a, b], [c, d], [e, f]]
spreadsheet_hdl.write_data(data, 'A2')

spreadsheet_hdl.append_data(data)
