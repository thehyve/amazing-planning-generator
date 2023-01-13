# amazing-planning-generator
Generate the weekly planning from and to a Google spreadsheet.

## Setup
1. Enable the [Google sheets API](https://developers.google.com/sheets/api/quickstart/python) and create a new project.
2. Create a new service account in the Google Developer Console for the project created in the previous step.
3. Create a JSON key, and download the `service_account.json` file.
4. If needed, update `config.yml` to point to the correct spreadsheets.
5. Open the source and target Google spreadsheets and give the service account access by granting view/edit rights to the service account email address.

## Run

### Docker
1. Put `service_account.json` into the project root.
2. Build the image: `docker build -t rwd-apg <path/to/project/root>`
3. Run with: `docker run rwd-apg /app/main.py -c .`

### Without Docker
1. Make sure `service_account.json` and `config.yml` are together somewhere in the same folder, e.g. `~/.config/gspread/`.
2. Install the dependencies listed in `requirements.txt` (Python 3.8+ is required).
3. Run `./main.py -c <path/to/folder/from/step1>` (See all options with `./main.py --help`).
