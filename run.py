"""
Set up the API connection
"""
import os
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
        wsname = wsh.title
        x_find = wsname.find('HTML')
        if x_find >= 0:
            pass
        else:
            manage_wsheet_data(wsh)


def manage_wsheet_data(worksheet):
    """
    Set up the Files to Write to using worksheet as title
    Open the file for writing to
    Get all the data from the worksheet
    """
    file_name = f"assets/htmlfiles/{worksheet}.txt"
    txt = worksheet.title
    write_table_definition(file_name, txt)
    table_data = get_all_data(worksheet)
    th_heading_data = get_the_table_header(table_data)
    style_rule = write_table_th(file_name, th_heading_data)
    td_rows_data = get_the_table_rows(table_data)
    write_table_td(file_name, td_rows_data, style_rule)
    write_table_foot(file_name)


def write_table_definition(txt_file, txt):
    """
    Put the table definition lines into a list
    Write the definition for the table to a txt file
    """

    table_defs = []

    swap = "swappera"
    tid = "tabletimeA"

    x_find = txt.find("Return")

    if x_find == 0:
        swap = "swapperb"
        tid = "tabletimeB"

    lines = [f'<figure id={swap} class="wp-block-table">']
    table_defs.append(lines)
    lines1 = ['<div style="overflow-x: auto;">']
    table_defs.append(lines1)
    lines2 = [f'<table id={tid} class="tabletime">']
    table_defs.append(lines2)
    lines3 = ['<tbody>']
    table_defs.append(lines3)
    lines4 = ['<tr>']
    table_defs.append(lines4)

    string_type = "complexList"

    append_multiple_lines(txt_file, table_defs, string_type)


def write_table_th(txt_file, header_data):
    """
    Loop through the th headings
    We watch out for colspan requirements in the head data
    Write to txt file
    """
    k = 0

    # Subsequent lines may have a colspan in them
    # This is picked up by counting the blanks trailing a header
    # This means we cannot actually write the head to the file until
    # we know how many blanks are trailing it.
    # So need to work in reverse

    thisth = []
    thisrule = []

    linex = '</tr>'
    thisth.append(linex)

    for head in reversed(header_data):
        # Check the head
        if head.find(":") > 0:
            thisrule.append(head)
            k = 0
            continue
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

    thisrvth = thisth.copy()
    thisrvth.reverse()
    string_type = "simpleList"

    append_multiple_lines(txt_file, thisrvth, string_type)

    return thisrule


def write_table_td(txt_file, table_data, row_rule):
    """
    Loop through the td rows
    Write to txt file
    """
    thistd = []
    icount = 1
    for table_row in table_data:
        linex = '<tr>'
        thistd.append(linex)
        for table_time in table_row:
            # is there a cell border rule
            if len(row_rule) == 0:
                linex = f'<td>{table_time}</td>'
            else:
                # There is a rule. Does it apply to this cell
                str_count = str(icount)
                str_rule = row_rule[0]
                xfind = str_rule.find(str_count)
                if xfind >= 0:  # we have a result
                    # get the rule
                    xslice = slice(xfind + 1, xfind + 2)  # the cut off
                    fslice = str_rule[xslice]  # Now we have the rule
                    # Use the rule
                    if fslice == "B":  # border left and right
                        bclass = 'class="thickborder"'
                    elif fslice == "L":  # border on left
                        bclass = 'class="thickleft"'
                    else:  # border on right
                        bclass = 'class="thickright"'
                    linex = f'<td {bclass}>{table_time}</td>'
                else:
                    linex = f'<td>{table_time}</td>'
            if table_time == "":
                continue
            thistd.append(linex)
            icount += 1
        linex = '</tr>'
        thistd.append(linex)
        icount = 1

    string_type = "simpleList"

    append_multiple_lines(txt_file, thistd, string_type)


def write_table_foot(txt_file):
    """
    Put the table footer lines into a list
    Write the footer for the table to a txt file
    """

    table_foot = []

    lines = ['</tbody>']
    table_foot.append(lines)
    lines1 = ['</table>']
    table_foot.append(lines1)
    lines2 = ['</div>']
    table_foot.append(lines2)
    lines3 = ['</figure>']
    table_foot.append(lines3)

    string_type = "complexList"

    append_multiple_lines(txt_file, table_foot, string_type)


def clear_txt_files():
    """
    Clear the txt files out before pulling from sheets begins
    """
    path_of_the_directory = 'assets/htmlfiles/'
    for file_name in os.listdir(path_of_the_directory):
        # construct full file path
        file = path_of_the_directory + file_name
        if os.path.isfile(file):
            print('Deleting file:', file)
            os.remove(file)


def write_file_back():
    """
    Write the file back to Worksheet
    """
    path_of_the_directory = 'assets/htmlfiles/'

    for filename in os.listdir(path_of_the_directory):
        fname = os.path.join(path_of_the_directory, filename)
        xfind = filename.find(":")  # look for this in filename
        xslice = slice(12, xfind - 4)  # now we know the cut off
        fslice = filename[xslice]  # Now we have the piece we want
        rcheck = f"HTML {fslice}"
        shfound = False

        text_file = open(fname, encoding='utf-8')
        data = text_file.read()
        htmlsh = ""

        # Now we have the file name with HTML prefix
        # Check that we do not have this sheet already
        # Loop through Google Sheets to see if it exists already
        for htmlsh in TSHEET.worksheets():
            httitle = htmlsh.title
            if httitle == rcheck:
                # If this is true then we already have this.
                # So flag that
                shfound = True

        if shfound:
            # we are editing not adding the sheet
            htmlsh.update('A1', data)
        else:
            # Go ahead and add the sheet
            try:
                wsh = TSHEET.add_worksheet(title=rcheck, rows=20, cols=1)

                wsh.update('A1', data)
            except ValueError as errnum:
                print(errnum)
                return False

    return True


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


def append_multiple_lines(file_name, lines_to_append, ls_type):
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
            if ls_type == "simpleList":
                line_txt = line
            else:
                line_txt = line[0]

            file_object.write(line_txt)


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
    th_heading = all_data.pop(0)

    return th_heading


def get_the_table_rows(all_data):
    """
    Get the table rows from the worksheet data
    """
    td_rows = all_data.copy()

    return td_rows


clear_txt_files()
my_list = list_sheets()
loop_through_worksheets(my_list)
# write_file_back()
