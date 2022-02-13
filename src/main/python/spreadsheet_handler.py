import os
import sys
import pickle
import json
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from drive_handler import DriveHandler
from fbs_runtime.application_context.PyQt5 import ApplicationContext


# TODO: Understand what does this "type" argument means
class Singleton(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args,
                                                                 **kwargs)
        return cls._instances[cls]


# TODO: Undestand metaclass implementation
class SpreadsheetHandler(metaclass=Singleton):
    def __init__(self, credentials=None, file_name=None, spreadsheet_id=None):
        self.credentials = credentials
        try:
            self.service = build('sheets', 'v4', credentials=self.credentials)
        except:
            DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
            self.service = build('sheets', 'v4', credentials=self.credentials, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
        
        self.spreadsheet_id = spreadsheet_id
        self.file_name = file_name

    def create_spreadsheet(self):

        drv_hdl = DriveHandler(self.credentials)

        # Check for existing spreadsheet:
        try:
            spreadsheet_id = drv_hdl.get_spreadsheet_id(self.file_name)
            self.spreadsheet_id = spreadsheet_id
        except FileNotFoundError:
            spreadsheet = {
                "properties": {
                    "title": self.file_name
                },
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1,
                            "title": "Expenses"
                        }
                    },
                    {
                        "properties": {
                            "sheetId": 2,
                            "title": "Incomes"
                        }
                    },
                    {
                        "properties": {
                            "sheetId": 3,
                            "title": "Summary"
                        }
                    },
                    {
                        "properties": {
                            "sheetId": 4,
                            "title": "Resources",
                            "hidden": True
                        }
                    }
                ]
            }
            spreadsheet = self.service.spreadsheets().create(body=spreadsheet,
                                                             fields="spreadsheetId").execute()

            self.spreadsheet_id = spreadsheet.get('spreadsheetId')

            self.format_spreadsheet()

        return self.spreadsheet_id

    def write_data(self, data, range, value_input_option='USER_ENTERED'):
        body = {
            'values': data
        }
        result = self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id, range=range,
            valueInputOption=value_input_option, body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))

    def append_data(self, data, range='Sheet1',
                    value_input_option='USER_ENTERED'):
        body = {
            'values': data
        }
        result = self.service.spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id, range=range,
            valueInputOption=value_input_option, body=body).execute()
        print('{0} cells appended.'.format(result
                                           .get('updates')
                                           .get('updatedCells')))

    def update_spreadsheet(self, json_file):
        # Check if json_file is the name of the file
        if isinstance(json_file, str):
            with open(json_file, encoding='utf-8') as json_file:
                body = json.load(json_file)

        # Check if json_file is the json already parsed as a dict
        elif isinstance(json_file, dict):
            body = json_file

        self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body=body).execute()

    def format_spreadsheet(self):
        appctx = ApplicationContext()
        self.update_spreadsheet(appctx.get_resource("spreadsheet_formatting.json"))

    def expenses_sort_by_date(self):
        appctx = ApplicationContext()
        self.update_spreadsheet(appctx.get_resource("expenses_sort_by_date.json"))

    def income_sort_by_date(self):
        appctx = ApplicationContext()
        self.update_spreadsheet(appctx.get_resource("income_sort_by_date.json"))

    def delete_spreadsheet(self):
        drv_hdl = DriveHandler(self.credentials)
        drv_hdl.service.files().delete(fileId=self.spreadsheet_id).execute()

    def rename_spreadsheet(self, new_name):
        drv_hdl = DriveHandler(self.credentials)

        file = {'name': new_name}

        drv_hdl.service.files().update(fileId=self.spreadsheet_id,
                                       body=file,
                                       fields='name').execute()

    def add_category(self, new_category):

        body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": 3,
                            "dimension": "COLUMNS",
                            "startIndex": 3,
                            "endIndex": 4
                        },
                        "inheritFromBefore": False
                    }
                },
                {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "stringValue": new_category
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue.stringValue",
                        "start": {
                            "sheetId": 3,
                            "rowIndex": 2,
                            "columnIndex": 3
                        }
                    }
                },
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": 3,
                            "startRowIndex": 3,
                            "endRowIndex": 4,
                            "startColumnIndex": 2,
                            "endColumnIndex": 3
                        },
                        "destination": {
                            "sheetId": 3,
                            "startRowIndex": 3,
                            "endRowIndex": 15,
                            "startColumnIndex": 3,
                            "endColumnIndex": 4
                        },
                        "pasteType": "PASTE_FORMULA"
                    }
                }
            ]
        }

        self.update_spreadsheet(body)

    def delete_category(self, category):
        # Find column index of "category"
        current_column_ascII = 67

        while True:
            current_category = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                                        range="Summary!" + chr(current_column_ascII) + "3").execute()

            if current_category['values'][0][0] == category:
                index = current_column_ascII - 65
                break
            else:
                current_column_ascII += 1

        body = {
            "requests": [
                {
                    "deleteDimension": {
                        "range": {
                            "sheetId": 3,
                            "dimension": "COLUMNS",
                            "startIndex": index,
                            "endIndex": index + 1
                        }
                    }
                }
            ]
        }

        self.update_spreadsheet(body)

    def read_categories(self):
        categories = list()
        current_column_ascII = 67

        while True:
            category = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                                range="Summary!" + chr(current_column_ascII) + "3").execute()

            if category['values'][0][0] == "Total":
                break
            else:
                categories.append(category['values'][0][0])
                current_column_ascII += 1

        return sorted(categories)

    def get_latest_upload(self, sheet_specification):

        latest_uploads = list()

        if sheet_specification == "expenses":
            range = "Expenses!C4:F"
        elif sheet_specification == "incomes":
            range = "Incomes!C4:E"
        else:
            raise ValueError("get_latest_upload: 'sheet_specification'" +
                             "must be 'expenses' or 'incomes'")

        latest_uploads_request = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheet_id,
                                                                               ranges=range).execute()
        # Check if the there is any data stored in the spreadsheet
        if "values" in latest_uploads_request['valueRanges'][0]:
            values = latest_uploads_request['valueRanges'][0]['values']
        else:
            # There is no data in the requested range, return a empty list
            return list()

        # Look for the items from the last 2 days displayed in the sheet
        flag = 0
        date = values[-1][0]
        latest_uploads.insert(0, values[-1][:])
        for row in values[-2::-1]:
            if row[0] == date:
                latest_uploads.insert(0, row)
            elif flag == 0 or flag == 1:
                # Strike 1 and Strike 2
                latest_uploads.insert(0, row)
                date = row[0]
                flag += 1
            else:
                # Strike 3
                break

        return latest_uploads
