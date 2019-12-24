import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from exceptions import *


SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']


class DriveHandler:
    def __init__(self, credentials):
        self.credentials = credentials
        self.service = build('drive', 'v3', credentials=self.credentials)

    def get_spreadsheet_id(self):

        results = self.service.files().list(
            q="mimeType='application/vnd.google-apps.spreadsheet' and name='expenses_tracker' and trashed=false",
            spaces="drive", fields="nextPageToken, files(id, name)").execute()

        item = results.get('files', [])

        if not item:
            raise FileNotFoundError
        if len(item) > 1:
            raise MultipleFilesFoundError
        else:
            return item[0]['id']
