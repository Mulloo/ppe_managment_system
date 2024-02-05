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
    name = input("Name:\n")
    type = input("Type:\n ")
    code = input("Code:\n ")
    serial = input("Serial:\n ")
    date_first_use = input("Date of first use dd/mm/yyyy:\n")
    date_of_manufacture = input("Date of Manufacture dd/mm/yyy:\n")
    print("Saving Data...")
    row_new_input = [
        name,
        type,
        code,
        serial,
        date_first_use,
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
    print()


def main():
    welcome_initial_input()


if __name__ == "__main__":
    main()
