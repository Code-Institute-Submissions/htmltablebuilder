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
    # f_name = open(f'{file_name}', 'w', encoding='utf-8')
    write_table_definition(file_name)
    table_data = get_all_data(worksheet)
    th_heading_data = get_the_table_header(table_data)
    th_heading_loop = loop_th_headings(th_heading_data)
    f_name = open(f'{file_name}', 'w', encoding='utf-8')
    write_table_th(f_name, th_heading_loop)


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


def append_new_line(file_name, text_to_append):
    """Append given text as a new line at the end of file"""
    # Open the file in append & read mode ('a+')
    with open(file_name, "a+", encoding='utf-8') as file_object:
        # Move read cursor to the start of file.
        file_object.seek(0)
        # If file is not empty then append '\n'
        data = file_object.read(100)
        if len(data) > 0:
            file_object.write("\n")
        # Append text at the end of file
        file_object.write(text_to_append)


def append_multiple_lines(file_name, lines_to_append):
    """Append given text lines as a new line at the end of file"""
    # Open the file in append & read mode ('a+')
    with open(file_name, "a+", encoding='utf-8') as file_object:
        append_eol = False
        # Move read cursor to the start of file.
        file_object.seek(0)
        # Check if file is not empty
        data = file_object.read(100)
        if len(data) > 0:
            append_eol = True
        # Iterate over each string in the list
        for line in lines_to_append:
            # If file is not empty then append '\n' before first line for
            # other lines always append '\n' before appending line
            if append_eol is True:
                file_object.write("\n")
            else:
                append_eol = True
            # Append element at the end of file
            file_object.write(line)


def loop_th_headings(header_data):
    """
    Loop through the th headings
    We watch out for colspan requirements in the head data
    """
    k = 0

    # Subsequent lines may have a colspan in them
    # This is picked up by counting the blanks trailing a header
    # This means we cannot actually write the head to the file until
    # we know how many blanks are trailing it.
    # So need to work in reverse

    thisth = []

    for head in reversed(header_data):
        # Check the head
        if head == "HEADEND":
            k = 0
            continue
        if head == "":
            k += 1
            continue
        if k > 0:
            linex = '<th style="background-color: #1d4d71;"'
            linex = linex + f' colspan="{k + 1}">{head}</th>'
            thisth.append(linex)
        else:
            linex = f'<th style="background-color: #1d4d71;">{head}</th>'
            thisth.append(linex)

        k = 0

    return thisth


def write_table_th(txt_file, th_data):
    """
    Write the th's for the table
    """
    open_tr = "<tr>"
    txt_file.write(f'{open_tr}\n')
    for each_th in reversed(th_data):
        txt_file.write(f'{each_th}\n')

    close_tr = "<\\tr>"
    txt_file.write(close_tr)


def get_all_data(data):
    """
    Get all the data from the worksheet in a list of lists
    """
    list_of_lists = data.get_all_values()

    return list_of_lists


def get_the_table_header(all_data):
    """
    Get the table header from the worksheet data
    """
    th_heading = all_data[0]

    return th_heading


my_list = list_sheets()
loop_through_worksheets(my_list)
