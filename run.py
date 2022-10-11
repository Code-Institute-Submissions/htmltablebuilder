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
        manage_wsheet_data(wsh)


def manage_wsheet_data(worksheet):
    """
    Set up the Files to Write to using worksheet as title
    Open the file for writing to
    Get all the data from the worksheet
    """
    file_name = f"assets/htmlfiles/{worksheet}.txt"
    f_name = open(f'{file_name}', 'w', encoding='utf-8')
    write_table_definition(f_name)
    table_data = get_all_data(worksheet)


def write_table_definition(txt_file):
    """
    Write the definition for the table
    """
    lines = ['<figure id="swappera" class="wp-block-table">\n']
    lines1 = ['<div style="overflow-x: auto;">\n']
    lines2 = ['<table id="tabletimeA" class="tabletime">\n']
    lines3 = ['<tbody>\n']
    lines4 = ['<tr>\n']

    with txt_file:
        txt_file.writelines(lines)
        txt_file.writelines(lines1)
        txt_file.writelines(lines2)
        txt_file.writelines(lines3)
        txt_file.writelines(lines4)


def get_all_data(data):
    """
    Get all the data from the worksheet in a list of lists
    """
    list_of_lists = data.get_all_values()

    return list_of_lists


def get_the_table_header(all_data):
    """
    Get the table header from the worksheet
    """
    headings = all_data[0]
    # for head in headings:
        # print(head)


my_list = list_sheets()
loop_through_worksheets(my_list)
