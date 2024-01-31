import gspread
from google.oauth2.service_account import Credentials
from pprint import pprint


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


def main():
    input


import_new()
