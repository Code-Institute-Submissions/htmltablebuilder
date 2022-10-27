"""
Pull in imports where needed
"""
import csv
import sys
import os
from time import sleep
import gspread
from google.oauth2.service_account import Credentials

# Set up the API connection
# Credit to Anna and Love Sandwiches Project

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
TSHEET = GSPREAD_CLIENT.open('html_table_builder')

LOGGEDIN = ""


def list_sheets():
    """
    Get List of WorkSheets in Sheet
    """

    try:
        worksheet_list = TSHEET.worksheets()
    except ValueError:
        print(f"Google Sheets Issue {ValueError()}")
    finally:
        return worksheet_list


def loop_through_worksheets(ws_list, u_name):
    """
    Loop through the worksheets
    """
    print("Preparing data from worksheets...")
    for wsh in ws_list:
        wsname = wsh.title  # Get the title of the sheet
        x_find = wsname.find('HTML')
        y_find = wsname.find('Rules')
        if x_find >= 0:  # We wish to ignore any tabs with a HTML prefix
            pass
        elif y_find >= 0:  # We wish to ignore the Rules Tab
            pass
        else:
            is_good = check_wsh_validity(wsh)  # Check Worksheet is Valid
            if is_good:
                print(f"Writing HTML Table code for {wsh} to file.")
                manage_wsheet_data(wsh)
            else:
                print(f"{wsh} is not correct. Please Fix it!\n")
                clear_console(10)
                main_menu(u_name)

    print("HTML Code for worksheets complete!\n")


def check_wsh_validity(w_sheet):
    """
    Error Trap that this sheet is valid to work on
    """
    wh_title = w_sheet.title
    # First thing is to check: has the user forgot to rename worksheet
    if wh_title[0:5].lower() == "sheet":
        # All worksheets need a Route Number
        print(f"Rename {wh_title} to a Route Number\n")
        data_good = False
    else:
        data_good = True

    # All worksheets labels should begin with a Route Number
    # Route Numbers have a minimum of 3 digits
    if wh_title[0:3].isnumeric():
        data_good = True
    else:
        # All worksheets need a Route Number
        print(f"Rename {wh_title} to a Route Number\n")
        data_good = False

    # Check if First Row is empty
    values_list = w_sheet.row_values(1)
    if len(values_list) == 0:
        # All worksheets need a Header
        print(f"{wh_title} has no Header\n")
        data_good = False
    else:
        data_good = True

    return data_good


def manage_wsheet_data(worksheet):
    """
    Set up the Files to Write to using worksheet name as title
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

    # Initialise these variables with defaults
    swap = "swappera"
    tid = "tabletimeA"

    x_find = txt.find("Return")  # Look for Return Prefix

    if x_find == 0:  # This is a follow on table to first table
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
    k = 0  # This counts blanks

    # Subsequent lines may have a colspan in them indicated by a blank
    # This is picked up by counting the blanks trailing a header
    # This means we cannot actually write the head to the file until
    # we know how many blanks are trailing it.
    # So need to work in reverse

    thisth = []
    thisrule = []

    linex = '</tr>'  # because we are working reverse the row close comes first
    thisth.append(linex)

    for head in reversed(header_data):
        # Check the head
        if head.find(":") > 0:  # We have found a rule
            thisrule.append(head)  # Append it to our rule list
            k = 0
            continue
        if head == "HEADEND":  # This is the end of the row
            k = 0
            continue
        if head == "":  # We have hit a blank so start counting
            k += 1
            continue
        if k > 0:  # There are blanks so we need a colspan
            ck_rt = check_creturn(head)  # Is there a new line in header
            linex = '<th style="background-color: #1d4d71;"'
            linex = linex + f' colspan="{k + 1}">{ck_rt}</th>'
            thisth.append(linex)
        else:  # There is no colspan
            ck_rt = check_creturn(head)  # Is there a new line in header
            linex = f'<th style="background-color: #1d4d71;">{ck_rt}</th>'
            thisth.append(linex)

        k = 0

    thisrvth = thisth.copy()  # Make a copy of list
    thisrvth.reverse()  # Reverse the list to now read in correct order

    string_type = "simpleList"

    append_multiple_lines(txt_file, thisrvth, string_type)

    return thisrule  # This will be needed to apply to table td rows


def write_table_td(txt_file, table_data, row_rule):
    """
    Loop through the td rows
    Write to txt file
    """
    thistd = []
    icount = 1  # This works instead of index
    # Index is not reliable in this instance

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


def check_creturn(head_txt):
    """
    Check for a carraige return in txt
    If so amend the text to include html <br />
    """
    decor_head = head_txt
    xfind = decor_head.find('\n')
    if xfind >= 0:
        xslice = slice(0, xfind)  # now we know the cut off
        fslice = decor_head[xslice]  # Now we have the piece we want
        yslice = slice(xfind + 1, len(decor_head))  # now we know the cut off
        lslice = decor_head[yslice]  # Now we have the piece we want
        decor_head = f'{fslice}<br />{lslice}'

    return decor_head


def write_file_back():
    """
    Write the file back to Worksheet
    See Credits pythonpool.com
    """
    print("Writing HTML code from files to worksheets...")

    path_of_the_directory = 'assets/htmlfiles/'

    for filename in os.listdir(path_of_the_directory):
        fname = os.path.join(path_of_the_directory, filename)
        xfind = filename.find(":")  # look for this in filename
        xslice = slice(12, xfind - 4)  # now we know the cut off
        fslice = filename[xslice]  # Now we have the piece we want
        rcheck = f"HTML {fslice}"

        print(f"Writing HTML code to {rcheck}!")

        text_file = open(fname, encoding='utf-8')
        data = text_file.read()  # read the open file

        # Go ahead and add the sheet
        try:
            wsh = TSHEET.add_worksheet(title=rcheck, rows=20, cols=1)
            wsh.update('A1', data)
        except ValueError as errnum:
            print(errnum)
            return False

    print("HTML Table code is now written from files to worksheets!")

    return True


def append_multiple_lines(file_name, lines_to_append, ls_type):
    """
    Append given text lines as a new line at the end of file
    See Credits thispointer.com
    """
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


def main():
    """
    Run all program functions
    """
    login_user()


def clear_all_sheets(u_name):
    """
    Clear all of the worksheets
    """
    print("This will remove all of the worksheets other than Rules\n")
    user_input = input(f'{u_name}: Do you wish to proceed (yes/no):\n')
    if user_input.lower() == 'yes':
        print("Ok, you typed yes, so the program is now executing!\n")
        print("Removing worksheets off of Google sheets...")
        # Loop through Google Sheets
        for wsheet in TSHEET.worksheets():
            if wsheet.title == "Rules":
                pass
            else:
                wsheet = TSHEET.del_worksheet(wsheet)
                print(f"{wsheet} is deleted")

        print("The worksheets are now deleted!\n")
        clear_console(1)

    if user_input.lower() != 'yes':
        print("Ok, you did not type yes, so returning to Main menu!")
        clear_console(1)

    main_menu(u_name)


def delete_txt_files(u_name):
    """
    Delete the txt files out before pulling from sheets begins
    See Credits pynative.com
    """
    print("This will remove all of the HTML Text Files\n")
    if u_name == "autom":
        deltxt_loop()
    else:
        user_input = input(f'{u_name}: Do you wish to proceed (yes/no):\n')
        if user_input.lower() == 'yes':
            print("Ok, you typed yes, so the program is now executing!\n")
            deltxt_loop()
        else:
            print("Ok, you did not type yes, so returning to Main menu!")

        clear_console(1)

        main_menu(u_name)


def deltxt_loop():
    """
    Loop through the txt files deleting them
    """
    print("Removing any existing txt files...")
    path_of_the_directory = 'assets/htmlfiles/'
    for file_name in os.listdir(path_of_the_directory):
        # construct full file path
        file = path_of_the_directory + file_name
        if os.path.isfile(file):
            print('Deleting file:', file)
            os.remove(file)

    print("The txt files are now deleted!\n")


def clear_html_sheets(u_name):
    """
    Clear all of the html worksheets
    """
    print("This will remove all of the HTML worksheets\n")
    if u_name == "autom":
        delhtml_loop()
    else:
        user_input = input(f'{u_name}: Do you wish to proceed (yes/no):\n')
        if user_input.lower() == 'yes':
            print("Ok, you typed yes, so the program is now executing!\n")
            delhtml_loop()
        else:
            print("Ok, you did not type yes, so returning to Main menu!")

        clear_console(1)

        main_menu(u_name)


def delhtml_loop():
    """
    Loop through all the HTML worksheets deleting them
    """
    print("Removing HTML worksheets off of Google sheets...")
    # Loop through Google Sheets
    # Check the sheet name for HTML prefix
    for htmlsh in TSHEET.worksheets():
        shtitle = htmlsh.title
        xfind = shtitle.find("HTML")  # look for this in filename
        if xfind >= 0:
            wsh = TSHEET.del_worksheet(htmlsh)
            print(f"{wsh} is deleted")

    print("The HTML worksheets are now deleted!\n")


def run_automation(u_name):
    """
    This function runs the automated html generation
    """
    print("This will overwrite, your previous HTML results!")
    print("This will not overwrite, your Worksheets!")
    user_input = input(f'{u_name}: Do you wish to proceed (yes/no):\n')
    if user_input.lower() == 'yes':
        print("Ok, you typed yes, so the program is now executing!\n")
        delete_txt_files("autom")
        clear_html_sheets("autom")
        my_list = list_sheets()
        loop_through_worksheets(my_list, u_name)
        write_file_back()
        print("\n")
        print("The program is finished executing!")
        print("Take a look at Google Sheets to view your HTML Table code.")
        print("Just copy the contents of Cell A1 in the HTML worksheet.")
        print("Then Paste into matching Wordpress Schedule Post.")
        print("Always Check the Wordpress Post Result in the Browser.\n")
        clear_console(10)

    if user_input.lower() != 'yes':
        print("Ok, you did not type yes, so returning to Main menu!")
        clear_console(1)

    main_menu(u_name)


def main_menu(u_name):
    """
    This is the main menu presented ot logged in users
    """
    print(f"Welcome {u_name} to HTML Table Builder Automation.")
    print("Building HTML Table Code from Google Sheets.\n")

    user_option = input("""
        C: Clear All Worksheets
        D: Delete All HTML Text Files
        H: Clear HTML Worksheets
        R: Run HTML Automation
        Q: Quit

        Please enter your choice here: """)

    if user_option == "C" or user_option == "c":
        clear_all_sheets(u_name)
    elif user_option == "D" or user_option == "d":
        delete_txt_files(u_name)
    elif user_option == "H" or user_option == "h":
        clear_html_sheets(u_name)
    elif user_option == "R" or user_option == "r":
        run_automation(u_name)
    elif user_option == "Q" or user_option == "q":
        print("Quitting")
        clear_console(1)
        sys.exit()


def login_user():
    """
    This function logs in users with the correct password
    """
    print("Welcome to HTML Table Builder Automation.")
    print(
        "Please Login!")
    # Ask the user for the password
    u_name = input(
        "Please enter your username:\n")
    p_word = input(
        "Please enter your password:\n")

    LOGGEDIN = check_password(u_name.lower(), p_word.lower())

    if LOGGEDIN is True:
        # Then we are good to go show menu
        main_menu(u_name.capitalize())


def check_password(u_name, p_word):
    """
    Check the password is valid
    """
    path_of_the_directory = 'assets/registered/'
    filename = "users.csv"
    fname = os.path.join(path_of_the_directory, filename)

    with open(fname) as f_lines:
        lines_list = csv.reader(f_lines, delimiter=',')
        for line in lines_list:
            if line[0] == u_name and line[1] == p_word:
                # This is a valid login
                LOGGEDIN = True
                break
            else:
                # This is not a valid login
                LOGGEDIN = False

    if LOGGEDIN is True:
        print("Login Successful\n")
        clear_console(1)

    else:
        print("Sorry username or password is incorrect\n")
        clear_console(1)

    return LOGGEDIN


def clear_console(how_long):
    # Waiting for 1 second to clear the screen
    sleep(how_long)

    # Clearing the Screen
    os.system('clear')


main()
