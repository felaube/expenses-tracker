This project intends to develop an integrated platform in Android (TODO), Ubuntu and Windows to track and categorize expenses and incomes.

A Google Sheets spreadsheet is used to store the incomes and expenses. The communication is provided by Google Sheet's Python API.

For now, the application is implemented only for Windows and Ubuntu.

Do not install the application in "Programs File" folder.

In order to use the software first you must install the released version, or download the source code. Then, enable the Google Sheets API following the instructions in the topic "Step 1: Turn on the Google Sheets API" in https://developers.google.com/sheets/api/quickstart/python. After enabling it, click on "DOWNLOAD CLIENT CONFIGURATION" to download a "credential.json" file, copy it and paste in the application installation folder or in src/main/python.

After doing so, you must enable the Google Drive API following the instructions in https://developers.google.com/drive/api/v3/enable-drive-api.
