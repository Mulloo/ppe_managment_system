import gspread
from google.oauth2.service_account import Credentials
from pprint import pprint
from simple_term_menu import TerminalMenu


SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

CREDS = Credentials.from_service_account_file("creds.json")
SCOPE_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPE_CREDS)
SHEET = GSPREAD_CLIENT.open("ppe_management_sheet")

# ppe = SHEET.worksheet("in_use").get_all_values()
# print(ppe)


def welcome_initial_input():
    print(
        "Hello welcome to PPE Management System \n Which of the following choices are you looking for\n 1: New Equipment Input\n 2: Quarantine An Item\n 3: Repair Equipment Log\n 4: Retire Equipment\n"
    )
    options = [
        "1: New Equipment Input",
        "2: Quarantine An Item",
        "3: Repair Equipment Log",
        "4: Retire Equipment",
    ]
    terminal_menu = TerminalMenu(options, title="Choices")
    choice_index = terminal_menu.show()
    print(choice_index)
    if choice_index == 0:
        import_new()
    elif choice_index == 1:
        quarantine_equipment()
    elif choice_index == 2:
        repair_equipment()
    elif choice_index == 3:
        retire_equipment()


def import_new():
    print(
        "Please give the required information\n Name: \n Type: \n Code: \n Serial: \n Date of first use: \n Date of Manufacture:\n"
    )
    name = input("Name: ")
    type = input("Type: ")
    code = input("Code: ")
    serial = input("Serial: ")
    date_first_use = input("Date of first use dd/mm/yyyy: ")
    date_of_manufacture = input("Date of Manufacture dd/mm/yyy: ")

    row_new_input = [
        name,
        type,
        code,
        serial,
        date_first_use,
        date_first_use,
        date_of_manufacture,
    ]
    print("Updating In-use sheet\n")
    SHEET.worksheet("in_use").append_row(row_new_input)
    print("In-use sheet successfully updated")


def quarantine_equipment():
    print()


def repair_equipment():
    print()


def retire_equipment():
    print()


def main():
    welcome_initial_input()


main()
