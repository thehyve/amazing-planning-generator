# amazing-planning-generator
Generate the weekly planning from and to a Google spreadsheet.

## Setup
1. Enable the [Google sheets API](https://developers.google.com/sheets/api/quickstart/python) 
and download the credentials.json file.
2. Create a folder at `~/.config/gspread/`.
3. Put credentials.json (downloaded in step 1) and config.yml (from this repo) into this folder.
4. Install the dependencies listed in `requirements.txt`, preferably into a virtual environment (Python 3.8 is recommended).

## Run
- If needed, update config.yml such that it points to the source and target spreadsheets.
- run ```python main.py``` (On the first run you will need to authenticate for both spreadsheets being accessed).
