#!/usr/bin/env python3

import logging
import pickle
from datetime import datetime
from pathlib import Path
from string import ascii_uppercase
from typing import Dict, List

import gspread
import numpy as np
import pandas as pd
import yaml
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from gspread import WorksheetNotFound
from gspread.models import Worksheet
from gspread.utils import rowcol_to_a1
from gspread_formatting import (
    Color, ColorStyle, ConditionalFormatRule, GradientRule, GridRange,
    InterpolationPoint, get_conditional_format_rules,
)
from itertools import cycle

CURR_WEEK_NR: int = datetime.now().isocalendar()[1]
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

CONFIG_DIR = Path.home() / '.config' / 'gspread'
CONFIG_FILE = CONFIG_DIR / 'config.yml'
CREDENTIALS_FILE = CONFIG_DIR / 'credentials.json'
SERVICE_ACCOUNT_FILE = CONFIG_DIR / 'service_account.json'
TOKEN_FILE = CONFIG_DIR / 'token.pickle'


logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)


def read_yaml_file(path: Path) -> Dict:
    with path.open('rt') as f:
        contents = yaml.safe_load(f.read())
    return contents


def gsheet_api_check(scopes: List[str]):
    try:
        return Credentials.from_service_account_file(str(SERVICE_ACCOUNT_FILE), scopes=scopes)
    except FileNotFoundError:
        logger.warning(f"No service account credentials available at {SERVICE_ACCOUNT_FILE} "
                       f"(Switching to regular user authentication)")

    creds = None
    if TOKEN_FILE.exists():
        with TOKEN_FILE.open('rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, scopes)
            creds = flow.run_console()
        with TOKEN_FILE.open('wb') as token:
            pickle.dump(creds, token)
    return creds


def pull_sheet_data(scopes: List[str], spreadsheet_id: str, range_name: str) -> List:
    creds = gsheet_api_check(scopes)
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name).execute()
    values = result.get('values', [])

    if not values:
        raise ValueError(f'No data found for workbook {spreadsheet_id} '
                         f'in worksheet {range_name}.')
    else:
        rows = sheet.values().get(spreadsheetId=spreadsheet_id,
                                  range=range_name).execute()
        data = rows.get('values')
        logger.info("Data collected")
        return data


def excel_col_to_int(col: str) -> int:
    col = col.upper()
    num = 0
    for c in col:
        if c not in ascii_uppercase:
            raise ValueError(f'Invalid column name {col}')
        num = num * 26 + (ord(c) - ord('A')) + 1
    return num - 1


def add_planning_worksheet_formatting(worksheet: Worksheet,
                                      project_type_header: List[str]
                                      ) -> None:
    # They are really quite beautiful
    rgb_colors = [
        (224, 187, 228),
        (149, 125, 173),
        (210, 145, 188),
        (254, 200, 216),
        (255, 223, 211),
        (193, 231, 227),
        (249, 240, 194),
        (177, 212, 236),
        (143, 193, 169),
    ]

    # Determine the first and last cell of consecutive project types
    merge_ranges: List[List[int]] = []
    previous_val = None
    for i, project_type in enumerate(project_type_header[1:], start=2):
        if project_type == 'Total':
            break
        if project_type != previous_val:
            merge_ranges.append([i, i])
        else:
            merge_ranges[-1][1] = i
        previous_val = project_type

    # Merge consecutive project type cells and color them
    for merge_range, rgb in zip(merge_ranges, cycle(rgb_colors)):
        start_col_idx = merge_range[0]
        end_col_idx = merge_range[1]
        worksheet.merge_cells(1, start_col_idx, 1, end_col_idx)

        color_range = ':'.join([rowcol_to_a1(1, start_col_idx), rowcol_to_a1(1, end_col_idx)])
        worksheet.format(color_range, {
            "backgroundColor": {"red": rgb[0]/256, "green": rgb[1]/256, "blue": rgb[2]/256}
        })

    # Center-align project names and make them bold
    worksheet.format("1:2", {
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "bold": True
        }
    })

    # Make the first column (with persons) bold
    worksheet.format("A:A", {
        "textFormat": {
            "bold": True
        }
    })

    # Freeze the first (person) column and the top 2 (project) rows
    worksheet.freeze(rows=2, cols=1)

    # Wrap text in project titles
    worksheet.format("2:2", {
        "wrapStrategy": "WRAP"
    })

    # Add conditional formatting, to create gradient for hour numbers
    grad_color = Color(102/256, 205/256, 170/256)
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range('B3:ZZ999', worksheet)],
        gradientRule=GradientRule(
            minpoint=InterpolationPoint(colorStyle=ColorStyle(themeColor='BACKGROUND'),
                                        type='NUMBER', value='-7'),
            maxpoint=InterpolationPoint(colorStyle=ColorStyle(rgbColor=grad_color),
                                        type='NUMBER', value='40'),
        )
    )

    rules = get_conditional_format_rules(worksheet)
    rules.clear()
    rules.append(rule)
    rules.save()


def get_week_planning(spreadsheet_id: str, range_name: str) -> pd.DataFrame:
    data = pull_sheet_data(SCOPES, spreadsheet_id, range_name)

    df = pd.DataFrame(data, columns=None)

    # Get the col idx of the current week based on week number
    week_row = df.iloc[1, :]
    col_idx_curr_week = week_row[week_row == str(CURR_WEEK_NR)].index.values[0]

    # Only keep rows after the first empty row
    idx_first_empty_row = df.index[df.isnull().all(axis=1)][0]
    idx_first_project_row = idx_first_empty_row + 1
    df = df.loc[idx_first_project_row:, ]

    # Forward fill project type/name in the first two columns
    df.loc[:, 0:1] = df.loc[:, 0:1].replace('', np.nan).ffill(axis=0)

    # Only keep columns project_type, project_name, person and hours
    df = df.loc[:, [0, 1, 2, col_idx_curr_week]]

    # Only keep full rows (project name, person and number of hours)
    df.replace('', np.nan, inplace=True)
    df.dropna(inplace=True, how='any')

    df.columns = ['project_type', 'project', 'person', 'hours']

    # Drop rows where hours == 0
    df = df[df.hours != '0']
    # Drop rows with persons that start with "?"
    df = df[~df.person.str.startswith('?')]

    df.set_index(keys=['project_type', 'project'], inplace=True)

    project_list = df.index.unique()
    person_list = sorted(df['person'].unique())

    # Create the overview_df (week planning) from df
    overview_df = pd.DataFrame(columns=project_list, index=person_list)

    for row in df.itertuples(index=True):
        try:
            hours = int(row.hours)
        except ValueError:
            logger.info(f'Ignoring invalid hours value "{row.hours}" for project {row.Index}')
            continue
        overview_df.loc[row.person, row.Index] = hours

    # Add column with hours week total per person
    overview_df.loc[:, 'Total'] = overview_df.sum(axis=1).astype(int)

    # Replace NaN with empty strings and reset index
    overview_df.fillna('', inplace=True)
    overview_df.index.name = ''
    overview_df.reset_index(inplace=True)
    return overview_df


def write_week_planning_to_gsheet(df: pd.DataFrame, spreadsheet_id: str) -> None:
    gc = gspread.oauth(flow=gspread.auth.console_flow)
    sht1 = gc.open_by_key(spreadsheet_id)

    # Create or replace worksheet
    new_worksheet_name = f"Week {CURR_WEEK_NR}"
    try:
        new_worksheet = sht1.worksheet(new_worksheet_name)
        sht1.del_worksheet(new_worksheet)
        logger.info(f'Worksheet {new_worksheet_name} already exists. Replacing.')
    except WorksheetNotFound:
        logger.info(f'Worksheet {new_worksheet_name} created')
    finally:
        new_worksheet = sht1.add_worksheet(title=new_worksheet_name, rows="100", cols="100")

    header_row1 = [val[0] for val in df.columns]
    header_row2 = [val[1] for val in df.columns]
    new_worksheet.update([header_row1, header_row2] + df.values.tolist())

    logger.info('Applying formatting to worksheet')
    add_planning_worksheet_formatting(worksheet=new_worksheet, project_type_header=header_row1)


if __name__ == '__main__':
    config = read_yaml_file(CONFIG_FILE)
    logger.info(f'Config params: {config}')

    week_planning_df = get_week_planning(spreadsheet_id=config['SOURCE_SPREADSHEET_ID'],
                                         range_name=config['SOURCE_WORKSHEET'])
    logger.info('Writing to target sheet')
    write_week_planning_to_gsheet(df=week_planning_df,
                                  spreadsheet_id=config['TARGET_SPREADSHEET_ID'])
    logger.info('Completed')
