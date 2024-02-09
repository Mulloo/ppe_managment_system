import gspread
import re
from google.oauth2.service_account import Credentials
from pprint import pprint
from simple_term_menu import TerminalMenu
from datetime import datetime


SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

CREDS = Credentials.from_service_account_file("creds.json")
SCOPE_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPE_CREDS)
SHEET = GSPREAD_CLIENT.open("ppe_management_sheet")


def welcome_initial_input():
    """Welcomes the user and displays the terminal menu options"""
    print(
        """\
  _____  _____  ______   __  __                                                   _      _____           _                 
 |  __ \|  __ \|  ____| |  \/  |                                                 | |    / ____|         | |                
 | |__) | |__) | |__    | \  / | __ _ _ __   __ _  __ _  ___ _ __ ___   ___ _ __ | |_  | (___  _   _ ___| |_ ___ _ __ ___  
 |  ___/|  ___/|  __|   | |\/| |/ _` | '_ \ / _` |/ _` |/ _ \ '_ ` _ \ / _ \ '_ \| __|  \___ \| | | / __| __/ _ \ '_ ` _ \ 
 | |    | |    | |____  | |  | | (_| | | | | (_| | (_| |  __/ | | | | |  __/ | | | |_   ____) | |_| \__ \ ||  __/ | | | | |
 |_|    |_|    |______| |_|  |_|\__,_|_| |_|\__,_|\__, |\___|_| |_| |_|\___|_| |_|\__| |_____/ \__, |___/\__\___|_| |_| |_|
                                                   __/ |                                        __/ |                      
                                                  |___/                                        |___/                       
          """
    )
    options = [
        "1: New Equipment Input",
        "2: Quarantine An Item",
        "3: Repair Equipment Log",
        "4: Retire Equipment",
    ]
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
        if re.match(code_pattern, code):
            print("Valid code Format.")
            break
    else:
        print("Code enter is invalid. Please use the format xxx/111")
    while True:
        serial = input("Serial:\n ").strip()
        serial_pattern = r"^\d{5}[A-Z]{2}\d{4}$"
        if re.match(serial_pattern, serial):
            print("Valid serial format.")
            break
        else:
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
    quarantine_item_code = input("Equipment unique code: ")
    print("Finding Equipment...")
    cell_find = SHEET.worksheet("in_use").find(quarantine_item_code)
    cell_row = (
        str(cell_find)
        .replace("Cell", "")
        .replace("C3", "")
        .replace("Cell", "")
        .replace(quarantine_item_code, "")
        .replace("''", "")
        .replace("R", "")
        .replace("<", "")
        .replace(">", "")
    )
    quarantine_item_row = SHEET.worksheet("in_use").row_values(int(cell_row))
    print(f"Confirm Data => {quarantine_item_row}")
    date_of_quarantine = input("Quarantined Date:  ")
    issue = input("Please specify the issue with this equipment:\n")
    quarantine_item_row.append(issue)
    quarantine_item_row.append(date_of_quarantine)
    print(f"adding date and issues please confirm {quarantine_item_row}")
    SHEET.worksheet("quarantine").append_row(quarantine_item_row)
    SHEET.worksheet("in_use").delete_rows(int(cell_row))
    print("Moving data to quarantine sheet...")
    print("Data move to sheet successful")


def repair_equipment():
    print()


def retire_equipment():
    """Gathers the equipment code and moves it to the retired sheet"""
    print("Equipment can only be retire if it is already placed into quarantine: \n")
    retired_equipment_code = input(
        "Please input the code of the equipment you wish to retire: \n"
    )
    date_of_destruction = input("Please input the date the equipment was destroyed: \n")
    cell_find = SHEET.worksheet("quarantine").find(retired_equipment_code)
    cell_row = (
        str(cell_find)
        .replace("Cell", "")
        .replace("C3", "")
        .replace("Cell", "")
        .replace(retired_equipment_code, "")
        .replace("''", "")
        .replace("R", "")
        .replace("<", "")
        .replace(">", "")
    )
    print("Getting equipment data")
    retired_equipment_row = SHEET.worksheet("quarantine").row_values(int(cell_row))
    print(f"Please confirm the data {retired_equipment_row}")
    print("Moving equipment to retired")
    retired_equipment_row.append(date_of_destruction)
    SHEET.worksheet("retired").append_row(retired_equipment_row)
    SHEET.worksheet("quarantine").delete_rows(int(cell_row))


def main():
    welcome_initial_input()


if __name__ == "__main__":
    main()
