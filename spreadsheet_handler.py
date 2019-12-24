import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from drive_handler import DriveHandler

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class SpreadsheetHandler:
    def __init__(self, credentials, spreadsheet_id=None):
        self.credentials = credentials
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.spreadsheet_id = spreadsheet_id

    def create_spreadsheet(self):

        drv_hdl = DriveHandler(self.credentials)
        # Check for existing spreadsheet:
        spreadsheet_id = drv_hdl.get_spreadsheet_id()

        if spreadsheet_id:
            self.spreadsheet_id = spreadsheet_id
            return spreadsheet_id
        else:
            spreadsheet = {
                'properties': {
                    'title': 'expenses_tracker'
                }
            }
            spreadsheet = self.service.spreadsheets().create(body=spreadsheet,
                                                             fields='spreadsheetId').execute()
            self.spreadsheet_id = spreadsheet_id
            return spreadsheet.get('spreadsheetId')

    def write_data(self, data, range, value_input_option='USER_ENTERED'):
        body = {
            'values': data
        }
        result = self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id, range=range,
            valueInputOption=value_input_option, body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))

    def append_data(self, data, range='Sheet1', value_input_option='USER_ENTERED'):
        body = {
            'values': data
        }
        result = self.service.spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id, range=range,
            valueInputOption=value_input_option, body=body).execute()
        print('{0} cells appended.'.format(result
                                           .get('updates')
                                           .get('updatedCells')))
