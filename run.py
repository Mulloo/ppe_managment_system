import re
import time
from datetime import datetime
import gspread
import tabulate
from google.oauth2.service_account import Credentials
from simple_term_menu import TerminalMenu
from colorama import Fore, Back, Style
from welcome import *

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

CREDS = Credentials.from_service_account_file("creds.json")
SCOPE_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPE_CREDS)
SHEET = GSPREAD_CLIENT.open("ppe_management_sheet")


def main_menu():
    """"""
    terminal_menu = TerminalMenu(options, title="Choices")
    choice_index = terminal_menu.show()
    if choice_index == 0:
        import_new()
    elif choice_index == 1:
        quarantine_equipment()
    elif choice_index == 2:
        repair_equipment()
    elif choice_index == 3:
        retire_equipment()
    elif choice_index == 4:
        update_equipment()
    elif choice_index == 5:
        view_sheet()


def import_new():
    """Asks the user for in puts to then adds them to the In_use worksheet"""
    print(
        f"""
Please give the required information
Name:
Type:
Code:
Serial
Date of first use:
Date of Manufacture:
    """
    )
    while True:
        name = input("Name:\n").strip()
        if all(x.isalpha() or x.isspace() for x in name) and name:
            break
        print("Name entered is invalid, letters and spaces only.")
    while True:
        type = input("Type:\n").strip()
        if all(x.isalpha() or x.isspace() for x in type) and type:
            break
        print("Type entered is invalid, letters and spaces only.")
    while True:
        code = input("Code:\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        try:
            re.match(code_pattern, code)
            print("Valid code Format.")
            break
        except ValueError:
            print("Code enter is invalid. Please use the format xxx/111")
    while True:
        serial = input("Serial:\n ").strip()
        serial_pattern = r"^\d{5}[A-Z]{2}\d{4}$"
        try:
            re.match(serial_pattern, serial)
            print("Valid serial format.")
            break
        except ValueError:
            print(
                "Serial number entered is invalid. Please use petzl's serial format ie:22041OI0001"
            )
    parsed_date_first_use = None
    while parsed_date_first_use is None:
        date_first_use = input("Date of first use dd/mm/yyyy:\n")
        try:
            parsed_date_first_use = datetime.strptime(date_first_use, "%d/%m/%Y")
            print("Date of first use valid", parsed_date_first_use)
        except ValueError:
            print("Invalid date format. Please user the format dd\mm\yyyy.")
    parsed_date_manufacture = None
    while parsed_date_manufacture is None:
        date_of_manufacture = input("Date of manufacture dd/mm/yyyy:\n")
        try:
            parsed_date_manufacture = datetime.strptime(date_first_use, "%d/%m/%Y")
            print("Date of manufacture valid", parsed_date_manufacture)
        except ValueError:
            print("Invalid date format. Please user the format dd\mm\yyyy.")
    print("Saving Data...")
    row_new_input = [
        name,
        type,
        code,
        serial,
        date_first_use,
        date_of_manufacture,
    ]
    print("Updating In-use Sheet...\n")
    SHEET.worksheet("in_use").append_row(row_new_input)
    print("In-use sheet successfully updated.")


def quarantine_equipment():
    """Gathers the code of quarantined equipment finds the cell and then
    finds and retrieves the row of information"""
    while True:
        quarantine_item_code = input("Code:\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, quarantine_item_code):
            print("Valid code Format.")
            break
    else:
        print("Code enter is invalid. Please use the format xxx/111")
    print("Finding Equipment...")
    cell_find = SHEET.worksheet("in_use").find(quarantine_item_code)
    cell_row = row_num_finder(cell_find, quarantine_item_code)
    quarantine_item_row = SHEET.worksheet("in_use").row_values(int(cell_row))
    print(f"Confirm Data => {quarantine_item_row}")
    date_of_quarantine = input("Quarantined Date:  ")
    issue = input("Please specify the issue with this equipment:\n")
    quarantine_item_row.append(issue)
    quarantine_item_row.append(date_of_quarantine)
    print(f"adding date and issues please confirm {quarantine_item_row}")
    SHEET.worksheet("quarantine").append_row(quarantine_item_row)
    SHEET.worksheet("in_use").delete_rows(int(cell_row))
    print(
        f"""Moving data to quarantine sheet... 
            Data move to sheet successful.   """
    )


def repair_equipment():
    print("Caution Item must already be in quarantine for a repair to be made.")
    while True:
        repair_item_code = input("Code:\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, repair_item_code):
            print("Valid code Format.")
            break
    else:
        print("Code enter is invalid. Please use the format xxx/111")
    cell_find = SHEET.worksheet("quarantine").find(repair_item_code)


def view_sheet():
    headers = [
        "name",
        "type",
        "code",
        "serial",
        "date_first_use",
        "date_of_manufacturing",
        "issue",
    ]
    table_data = []
    worksheet_titles = [sheet.title for sheet in SHEET.worksheets()]
    worksheet_titles.append("Back")
    terminal_menu = TerminalMenu(
        worksheet_titles, title="Select the sheet you whish to view:"
    )
    menu_entry_index = terminal_menu.show()
    if menu_entry_index == len(worksheet_titles) - 1:
        print("Returning to the previous menu...")
        time.sleep(3)
        main()
    sheet_selected = SHEET.get_worksheet(menu_entry_index)
    all_rows = sheet_selected.get_all_values()
    print(f"\nContents of '{sheet_selected}':\n")
    for row in all_rows:
        print(" | ".join(row))
        table_data.append(row)
        print(tabulate(table_data, headers=headers))


def update_equipment():
    print()


def retire_equipment():
    """Gathers the equipment code and moves it to the retired sheet"""
    print("Equipment can only be retire if it is already placed into quarantine: \n")
    while True:
        retired_equipment_code = input("Code:\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, retired_equipment_code):
            print("Valid code Format.")
            break
    else:
        print("Code enter is invalid. Please use the format xxx/111 (x = [a-z])")
    retired_equipment_code = input(
        "Please input the code of the equipment you wish to retire: \n"
    )
    date_of_destruction = input("Please input the date the equipment was destroyed: \n")
    cell_find = SHEET.worksheet("quarantine").find(retired_equipment_code)
    cell_row = row_num_finder(cell_find, retired_equipment_code)
    print("Getting equipment data")
    retired_equipment_row = SHEET.worksheet("quarantine").row_values(int(cell_row))
    print(f"Please confirm the data {retired_equipment_row}")
    print("Moving equipment to retired")
    retired_equipment_row.append(date_of_destruction)
    SHEET.worksheet("retired").append_row(retired_equipment_row)
    SHEET.worksheet("quarantine").delete_rows(int(cell_row))


def row_num_finder(cell_find, equipment_code):
    """Removes all but the row number form cell_find"""
    cell_row = (
        str(cell_find)
        .replace("Cell", "")
        .replace("C3", "")
        .replace("Cell", "")
        .replace(equipment_code, "")
        .replace("''", "")
        .replace("R", "")
        .replace("<", "")
        .replace(">", "")
    )
    return cell_row


def main():
    welcome_initial_input()
    main_menu()


if __name__ == "__main__":
    main()
