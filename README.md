# amazing-planning-generator
Generate the weekly planning from a Google spreadsheet.

## Setup
- Enable the [Google sheets API](https://developers.google.com/sheets/api/quickstart/python) 
and download the credentials.json file.
- Create a folder at `~/.config/gspread/`.
- Put credentials.json and config.yml from this repo into this folder.

## Run
- Update config.yml such that it points to the source and target 
  spreadsheets and will select the correct week column.
- run ```python main.py```