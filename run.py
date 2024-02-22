import re
import time
import textwrap
from datetime import datetime
import gspread
from tabulate import tabulate
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


def go_to_main_menu():
    input("Hit enter to go back")
    welcome_initial_input()
    main_menu()


def main_menu():
    """
    Sets the choices for the terminal menu
    and executes the required function
    """
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
    """
    Asks the user for all necessary inputs then adds them to
    the In_use worksheet
    """
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

    # Gather name of equipment and validate it only has letters and spaces
    while True:
        name = input("Name: (Enter 'b' to exit import new)\n").strip()
        if name == "b":
            go_to_main_menu()
        elif all(x.isalpha() or x.isspace() for x in name) and name:
            print(Fore.GREEN + "Valid Name Format" + Fore.RESET)
            break
        else:
            print(
                Fore.RED
                + "Name entered is invalid, letters and spaces only."
                + Fore.RESET
            )

    # Gather equipment type and insure it only has letters and spaces
    while True:
        type = input("Type: (Enter 'b' to exit import new)\n").strip()
        if type == "b":
            go_to_main_menu()
        elif all(x.isalpha() or x.isspace() for x in type) and type:
            print(Fore.GREEN + "Valid Type" + Fore.RESET)
            break
        else:
            print(
                Fore.RED
                + "Type entered is invalid, letters and spaces only."
                + Fore.RESET
            )

    # Gather code from user validate it using re.match with the pattern given
    while True:
        code = input("Code: (enter 'b' to exit import new)\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, code):
            print(Fore.GREEN + "Valid code Format." + Fore.RESET)
            break
        elif code == "b":
            go_to_main_menu()
        else:
            print(
                Fore.RED
                + "Code enter is invalid. Please use the format xxx/111"
                + Fore.RESET
            )

    # Gather serial from user validate it using re.match with the pattern given
    while True:
        serial = input("Serial: (enter 'b' to exit import new)\n ").strip()
        serial_pattern = r"^\d{5}[A-Z]{2}\d{4}$"
        if re.match(serial_pattern, serial):
            print(Fore.GREEN + "Valid serial format." + Fore.RESET)
            break
        elif serial == "b":
            go_to_main_menu()
        else:
            print(
                Fore.RED
                + "Serial number entered is invalid. Please use petzl's serial \n format ie:22041OI0000"
                + Fore.RESET
            )

    # set date of first use to None
    parsed_date_first_use = None

    # Gather first use from user validate it using datetime
    while parsed_date_first_use is None:
        date_first_use = input("Date of first use dd/mm/yyyy:\n")
        if date_first_use == "b":
            go_to_main_menu()
        else:
            try:
                parsed_date_first_use = datetime.strptime(date_first_use, "%d/%m/%Y")
                print(
                    Fore.GREEN
                    + f"Date of first use valid {parsed_date_first_use}"
                    + Fore.RESET
                )
            except ValueError:
                print(
                    Fore.RED
                    + "Invalid date format. Please user the format dd\mm\yyyy."
                    + Fore.RESET
                )

    # set date of manufacture use to None
    parsed_date_manufacture = None

    # Gather manufactured date from user validate it using datetime
    while parsed_date_manufacture is None:
        date_of_manufacture = input("Date of manufacture dd/mm/yyyy:\n")
        if date_of_manufacture == "b":
            go_to_main_menu()
        else:
            try:
                parsed_date_manufacture = datetime.strptime(date_first_use, "%d/%m/%Y")
                print(
                    Fore.GREEN
                    + f"Date of manufacture valid {parsed_date_manufacture}"
                    + Fore.RESET
                )
            except ValueError:
                print(
                    Fore.RED
                    + "Invalid date format. Please user the format dd\mm\yyyy."
                    + Fore.RESET
                )
    print(Fore.YELLOW + "Saving Data..." + Fore.RESET)

    # set row data using inputs taken from user
    row_new_input = [
        name,
        type,
        code,
        serial,
        date_first_use,
        date_of_manufacture,
    ]
    print(Fore.YELLOW + "Updating In-use Sheet...\n" + Fore.RESET)

    # add new equipment to in_use sheet
    SHEET.worksheet("in_use").append_row(row_new_input)
    print(Fore.YELLOW + "In-use sheet successfully updated." + Fore.RESET)
    go_to_main_menu()


def quarantine_equipment():
    """
    Gathers the code of quarantined equipment finds the cell and then
    finds and retrieves the row of information the user inputs the equipments
    issue and appends it to the row it is then added to the quarantine sheet
    and removed from the in_use sheet
    """
    while True:
        quarantine_item_code = input(
            "Code, Please use the format xxx/111:\n(enter 'b' to exit)"
        ).strip()

        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, quarantine_item_code):
            print(Fore.GREEN + "Valid code Format." + Fore.RESET)
            break
        elif quarantine_item_code == "b":
            go_to_main_menu()
        else:
            print(
                Fore.RED
                + "Code enter is invalid. Please use the format xxx/111"
                + Fore.RESET
            )

    print(Fore.YELLOW + "Finding Equipment..." + Fore.RESET)

    # Find the cell and send it to row number finder
    cell_find = SHEET.worksheet("in_use").find(quarantine_item_code)
    cell_row = row_num_finder(cell_find, quarantine_item_code)

    # Find row data
    quarantine_item_row = SHEET.worksheet("in_use").row_values(int(cell_row))

    # send data found to user to insure correct equipment is returned
    print(Fore.YELLOW + f"Confirm Data => {quarantine_item_row}" + Fore.RESET)

    # gather issue from user to add to row later
    issue = input("Please specify the issue with this equipment:\n")

    # set date to None
    parsed_date_quarantine = None

    # gather date form user validate format using datetime
    while parsed_date_quarantine is None:
        date_quarantine = input("Date of Quarantine dd/mm/yyyy:\n")
        try:
            parsed_date_quarantine = datetime.strptime(date_quarantine, "%d/%m/%Y")
            print(
                Fore.GREEN
                + f"Date of quarantine valid {parsed_date_quarantine}"
                + Fore.RESET
            )
        except ValueError:
            print(
                Fore.RED
                + "Invalid date format. Please user the format dd\mm\yyyy."
                + Fore.RESET
            )

    # append issue
    quarantine_item_row.append(issue)

    # append date
    quarantine_item_row.append(date_quarantine)
    print(
        Fore.YELLOW
        + f"adding date and issues please confirm {quarantine_item_row}"
        + Fore.RESET
    )

    # quarantine row added to quarantine sheet
    SHEET.worksheet("quarantine").append_row(quarantine_item_row)

    # quarantine row removed from in_use sheet
    SHEET.worksheet("in_use").delete_rows(int(cell_row))
    print(
        f""" {Fore.YELLOW} Moving data to quarantine sheet...
Data move to sheet successful.  {Fore.RESET} """
    )
    go_to_main_menu()


def repair_equipment():
    """
    Gathers the unique code form the user gets the cell then row data
    takes that data adds the job completed and logs it in the repair sheet
    """
    print("Caution Item must already be in quarantine \n for a repair to be made.")
    while True:
        repair_item_code = input(
            "Code: xxx/000 (enter 'b' to return the main menu)\n"
        ).strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, repair_item_code):
            print(Fore.GREEN + "Valid code Format." + Fore.RESET)
            break
        elif repair_item_code == "b":
            go_to_main_menu()
            break
        else:
            print(
                Fore.RED
                + "Code enter is invalid. Please use the format xxx/111"
                + Fore.RESET
            )

    # Find the cell and send it to row number finder
    try:
        cell_find = SHEET.worksheet("quarantine").find(repair_item_code)
    except TypeError:
        print(Fore.RED + "Item must be quarantined first" + Fore.RESET)

    # Find the cell row using row_num_finder function
    cell_row = row_num_finder(cell_find, repair_item_code)
    print(Fore.YELLOW + "Getting equipment data..." + Fore.RESET)

    # Find row data
    repair_equipment_row = SHEET.worksheet("quarantine").row_values(int(cell_row))

    # Gather Job done by the user.
    job_completed = input("Please detail the job completed: ")

    # add equipment back to in_use sheet without job to keep uniform
    SHEET.worksheet("in_use").append_row(repair_equipment_row[:-2])

    # add job done to the row
    repair_equipment_row.append(job_completed)
    print(repair_equipment_row)
    print(Fore.YELLOW + "Logging repair" + Fore.RESET)

    # add row to repair log sheet
    SHEET.worksheet("repair").append_row(repair_equipment_row)
    print(Fore.YELLOW + "Repair Logged" + Fore.RESET)

    # remove equipment form quarantine sheet
    SHEET.worksheet("quarantine").delete_rows(int(cell_row))
    go_to_main_menu()


def view_sheet():
    headers = [
        "name",
        "type",
        "code",
        "serial",
        "first_use",
        "manufacture_date",
        "issue",
        "quarantine_date",
        "destruction_date",
        "job",
    ]

    # set table data to empty array
    table_data = []

    # set worksheet titles form sheet dependent on sheet choice
    worksheet_titles = [sheet.title for sheet in SHEET.worksheets()]

    # add the back option to the titles
    worksheet_titles.append("Back")
    terminal_menu = TerminalMenu(
        worksheet_titles, title="Select the sheet you whish to view:"
    )

    menu_entry_index = terminal_menu.show()
    if menu_entry_index == len(worksheet_titles) - 1:
        print(Fore.YELLOW + "Returning to the previous menu..." + Fore.RESET)
        time.sleep(3)
        main()
    # get sheet from user choice
    sheet_selected = SHEET.get_worksheet(menu_entry_index)

    # get all values form sheet chosen
    all_rows = sheet_selected.get_all_values()
    print(f"\nContents of '{sheet_selected}':\n")

    max_column_width = {
        "name": 5,
        "type": 6,
        "code": 8,
        "serial": 10,
        "first_use": 7,
        "manufacture_date": 7,
        "issue": 5,
        "quarantine_date": 5,
        "destruction_date": 5,
        "job": 5,
    }

    # add each row to all data
    for row in all_rows[1:]:
        wrapped_row = [
            "\n".join(textwrap.wrap(cell, width=max_column_width[header]))
            for cell, header in zip(row, headers)
        ]
        table_data.append(wrapped_row)

    # print table to terminal
    print(tabulate(table_data, headers=headers, tablefmt="simple"))

    # back to start
    go_to_main_menu()


def update_equipment():
    headers = [
        "name",
        "type",
        "code",
        "serial",
        "first_use",
        "manufacture_date",
        "issue",
        "quarantine_date",
        "destruction_date",
        "job",
    ]

    # set worksheet titles form sheet dependent on sheet choice
    worksheet_titles = [sheet.title for sheet in SHEET.worksheets()]

    worksheet_titles.append("Back")
    terminal_menu = TerminalMenu(
        worksheet_titles, title="Select the sheet you wish to update: "
    )

    menu_entry_index = terminal_menu.show()
    if menu_entry_index == len(worksheet_titles) - 1:
        print(Fore.YELLOW + "Returning to the previous menu..." + Fore.RESET)
        time.sleep(3)
        main()
    # get sheet from user choice
    sheet_selected = SHEET.get_worksheet(menu_entry_index)

    # Gather code form user to find equipment for update
    while True:
        update_item_code = input(
            "Code: xxx/000 (enter 'b' to return the main menu)\n"
        ).strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, update_item_code):
            print(Fore.GREEN + "Valid code Format." + Fore.RESET)
            break
        elif update_item_code == "b":
            return
        else:
            print(
                Fore.RED
                + "Code enter is invalid. Please use the format xxx/111"
                + Fore.RESET
            )

    # Find the cell of the code given by the user
    try:
        cell_find = sheet_selected.find(update_item_code)
    except TypeError:
        print(Fore.RED + "Item code not found" + Fore.RESET)
        return

    # Get all values from the row found
    update_item_row = sheet_selected.row_values(cell_find.row)
    print(Fore.YELLOW + "Current Data:" + Fore.RESET)

    # Print data to user for to confirm
    for header, value in zip(headers, update_item_row):
        print(f"{Fore.CYAN}{header}: {value}{Fore.RESET}")

    attribute_menu = TerminalMenu(
        headers, title="Select the attribute you wish to update: "
    )
    attribute_index = attribute_menu.show()

    update_attribute = headers[attribute_index]
    new_value = input(
        f"Enter new value for {update_attribute}:\n Please make sure to use the correct format\n (Enter 'cancel' to return to main menu) "
    ).strip()

    if new_value == "cancel":
        go_to_main_menu()

    # Send data to sheet using row, column+1 and new value
    sheet_selected.update_cell(cell_find.row, attribute_index + 1, new_value)
    print(Fore.GREEN + f"{update_attribute} updated successfully." + Fore.RESET)
    go_to_main_menu()


def retire_equipment():
    """
    Gathers the equipment code gathers the row data adds the date of
    destruction and then moves it to the retired sheet the row is then removed
    from the in_use sheet
    """
    print(
        Fore.RED
        + "Equipment can only be retire if it is already placed into quarantine"
        + Fore.RESET
    )

    # gather code form user and validate t user re.match
    while True:
        retired_equipment_code = input("Code: (enter 'b' to exit)\n").strip()
        code_pattern = r"^[a-z]+/\d+$"
        if re.match(code_pattern, retired_equipment_code):
            print(Fore.GREEN + "Valid code Format." + Fore.RESET)
            break
        elif retired_equipment_code == "b":
            go_to_main_menu()
        else:
            print(
                Fore.RED
                + "Code enter is invalid. Please use the format xxx/111 (x = [a-z])"
                + Fore.RESET
            )

    # gather date of destruction from user and validate it with datetime
    parsed_date_destruction = None
    while parsed_date_destruction is None:
        date_destruction = input("Date of destruction dd/mm/yyyy:\n")
        try:
            parsed_date_destruction = datetime.strptime(date_destruction, "%d/%m/%Y")
            print(
                Fore.GREEN
                + f"Date of destruction valid {parsed_date_destruction}"
                + Fore.RESET
            )
        except ValueError:
            print(
                Fore.RED
                + "Invalid date format. Please user the format dd\mm\yyyy."
                + Fore.RESET
            )

    # find the code in the quarantine sheet
    cell_find = SHEET.worksheet("quarantine").find(retired_equipment_code)

    # find the row number by sending it to the row num finder function
    cell_row = row_num_finder(cell_find, retired_equipment_code)
    print(Fore.YELLOW + "Getting equipment data..." + Fore.RESET)

    # get all the data form the the row returned from the row number finder
    retired_equipment_row = SHEET.worksheet("quarantine").row_values(int(cell_row))
    user_confirm = input(
        f"Please confirm the data {retired_equipment_row}\nenter 'y' to Confirm 'n' to cancel"
    ).strip()
    if user_confirm.lower() == "y":

        print(Fore.YELLOW + "Moving equipment to retired" + Fore.RESET)

        # adds the date of destruction to the row
        retired_equipment_row.append(date_destruction)

        # adds the new retire row data to the retired sheet
        SHEET.worksheet("retired").append_row(retired_equipment_row)

        # removes the row from quarantine sheet
        SHEET.worksheet("quarantine").delete_rows(int(cell_row))
        go_to_main_menu()
    else:
        go_to_main_menu()


def row_num_finder(cell_find, equipment_code):
    """
    Removes all unwanted string characters form cell_find the cell
    row number is then returned
    """
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
    """Starts the welcome function the loads the main menu"""
    welcome_initial_input()
    main_menu()


if __name__ == "__main__":
    main()
