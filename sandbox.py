import os
import sys
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from drive_handler import DriveHandler
from spreadsheet_handler import SpreadsheetHandler
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox,\
                            QDoubleSpinBox, QWidget, QVBoxLayout,\
                            QHBoxLayout, QDateEdit
from PyQt5.QtCore import QDate


SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive.metadata.readonly']

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


spreadsheet_hdl = SpreadsheetHandler(CREDS)
spreadsheet_hdl.create_spreadsheet()

# data = [["3", "4"], ["=2+3", "=85"], ["34", "15"]]
# spreadsheet_hdl.write_data(data, 'A2')


def submit_button_clicked():
    alert = QMessageBox()
    data = [[str(doublespinBox.value())]]
    spreadsheet_hdl.append_data(data)
    alert.setText("The expense was submitted!")
    alert.exec_()


app = QApplication([])
window = QWidget()
verticallayout = QVBoxLayout()
horizontallayout = QHBoxLayout()

today = QDate.currentDate()

date = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy", date=today)
submitbutton = QPushButton('Click')
doublespinBox = QDoubleSpinBox(maximum=1000, decimals=2, minimum=0)

submitbutton.clicked.connect(submit_button_clicked)

horizontallayout.addWidget(doublespinBox)
horizontallayout.addWidget(date)

verticallayout.addLayout(horizontallayout)
verticallayout.addWidget(submitbutton)

window.setLayout(verticallayout)

window.show()

sys.exit(app.exec_())

# a, b, c, d, e, f = input("Insira 6 valore: ").split()
