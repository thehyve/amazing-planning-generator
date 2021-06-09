# amazing-planning-generator
Generate the weekly planning from and to a Google spreadsheet.

## Setup
1. Create a folder at `~/.config/gspread/` and copy config.yml from this repo into it.
2. Install the dependencies listed in `requirements.txt`, preferably into a virtual environment (Python 3.8 is recommended).
3. Enable the [Google sheets API](https://developers.google.com/sheets/api/quickstart/python) and create a new project.

When using a service account (recommended)
4. Create a new service account in the Google Developer Console for your project created in step 3.
5. Create a new JSON key, download and store at `~/.config/gspread/service_account.json`.
6. Give the service account access to the relevant Google spreadsheets by granting view/edit rights to the service account email address.

Alternatively you can use your own Google account directly
4. Create an OAuth 2.0 Client ID, then download and store the client secret at `~/.config/gspread/credentials.json`.

## Run
- If needed, update config.yml such that it points to the source and target spreadsheets.
- run ```python main.py``` (On the first run you will need to authenticate for both spreadsheets being accessed).
