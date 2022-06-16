import os
import sys
import pickle
import webbrowser
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from spreadsheet_handler import SpreadsheetHandler
from widget_gallery import WidgetGallery
import google.auth.exceptions as GoogleExceptions
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import Qt

SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.metadata.readonly',
          'https://www.googleapis.com/auth/drive.file']

if __name__ == '__main__':

    if not os.path.exists('credentials.json'):

        error_message = "The application did not find the 'credentials.json' file." + \
                        " In order to use this application, " + \
                        "you must first enable the Google " + \
                        "Sheets API on your google account." + \
                        "\n\nPlease acccess " + \
                        "https://developers.google.com/sheets/api/quickstart/python" + \
                        ", go to Step 1, dowload the configuration" + \
                        " file 'credentials.json' " + \
                        "and save it on the application folder. " + \
                        "\n\nThen, access " + \
                        "https://developers.google.com/drive/api/v3/enable-drive-api " + \
                        "to enable the Drive API on your google account. " + \
                        "After following the steps above, run the application again."
        print(error_message)

        app = QApplication([])
        error_window = QMessageBox()
        error_window.setWindowTitle("Expenses Tracker")
        error_window.setWindowIcon(QIcon(ApplicationContext().get_resource("submit.ico")))
        error_window.setIcon(QMessageBox.Critical)
        error_window.setText(error_message)
        error_window.setStandardButtons(QMessageBox.Ok)
        error_window.setDefaultButton(QMessageBox.Ok)
        error_window.setTextInteractionFlags(Qt.TextSelectableByMouse)

        webbrowser.open("https://developers.google.com/sheets/api/quickstart/python")
        webbrowser.open("https://developers.google.com/drive/api/v3/enable-drive-api")

        sys.exit(error_window.exec_())

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
            try:
                CREDS.refresh(Request())
            except GoogleExceptions.RefreshError:
                error_message = "Your access token has expired. Please, complete the authentication flow once again, " + \
                                "accessing your google account on the website" + \
                                " that will be opened when you close this window."
                print(error_message)
                app = QApplication([])
                error_window = QMessageBox()
                error_window.setWindowTitle("Expenses Tracker")
                error_window.setWindowIcon(QIcon(ApplicationContext().get_resource("submit.ico")))
                error_window.setIcon(QMessageBox.Critical)
                error_window.setText(error_message)
                error_window.setStandardButtons(QMessageBox.Ok)
                error_window.setDefaultButton(QMessageBox.Ok)
                error_window.exec_()

                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                CREDS = flow.run_local_server(port=0)
        else:
            error_message = "Please, complete the authentication flow " + \
                            "accessing your google account on the website" + \
                            " that will be opened when you close this window."
            print(error_message)
            app = QApplication([])
            error_window = QMessageBox()
            error_window.setWindowTitle("Expenses Tracker")
            error_window.setWindowIcon(QIcon(ApplicationContext().get_resource("submit.ico")))
            error_window.setIcon(QMessageBox.Critical)
            error_window.setText(error_message)
            error_window.setStandardButtons(QMessageBox.Ok)
            error_window.setDefaultButton(QMessageBox.Ok)
            error_window.exec_()

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            CREDS = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(CREDS, token)

    spreadsheet_hdl = SpreadsheetHandler(CREDS, file_name="Expenses Tracker")
    spreadsheet_hdl.create_spreadsheet()

    application_context = ApplicationContext()
    application_context.app.setFont(QFont("Fixed", 8))
    gallery = WidgetGallery()
    gallery.show()
    sys.exit(application_context.app.exec_())
