import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from exceptions import MultipleFilesFoundError


# TODO: Understand what does this "type" argument means
class Singleton(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args,
                                                                 **kwargs)
        return cls._instances[cls]


# TODO: Undestand metaclass implementation
class DriveHandler(metaclass=Singleton):
    def __init__(self, credentials=None):
        self.credentials = credentials
        self.service = build('drive', 'v3', credentials=self.credentials)

    def get_spreadsheet_id(self, file_name):

        results = self.service.files().list(  # pylint: disable=no-member
            q="mimeType='application/vnd.google-apps.spreadsheet' and \
               name='{}' and trashed=false".format(file_name),
            spaces="drive", fields="nextPageToken, files(id, name)").execute()

        item = results.get('files', [])

        if not item:
            raise FileNotFoundError
        if len(item) > 1:
            raise MultipleFilesFoundError
        else:
            return item[0]['id']
