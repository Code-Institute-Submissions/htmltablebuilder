"""
Set up the API connection
"""
import gspread
from google.oauth2.service_account import Credentials

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
TSHEET = GSPREAD_CLIENT.open('html_table_builder')


def list_sheets():
    """
    Get List of WorkSheets in Sheet
    """
    worksheet_list = TSHEET.worksheets()

    return worksheet_list


def loop_through_worksheets(ws_list):
    """
    Loop through the worksheets
    """
    for wsh in ws_list:
        get_wsheet_data(wsh)


def get_wsheet_data(worksheet):
    """
    Set up the File to Print to
    Get all the data from the worksheet
    """
    lines = ['<figure id="swappera" class="wp-block-table">']
    with open(f"assets/htmlfiles/{worksheet}.txt", 'w', encoding='utf-8') as f:
        for line in lines:
            f.write(line)
            f.write('\n')

    print(worksheet)
    all_data = worksheet.get_all_values()
    get_the_table_header(all_data)


def get_the_table_header(all_data):
    """
    Get the table header from the worksheet
    """
    headings = all_data[0]
    # for head in headings:
        # print(head)


my_list = list_sheets()
loop_through_worksheets(my_list)
