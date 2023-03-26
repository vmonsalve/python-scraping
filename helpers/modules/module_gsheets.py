import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv()

NAME_GOOGLE_ACCOUNT = os.getenv('NAME_GOOGLE_ACCOUNT')
JSON_KEY            = os.getenv('JSON_KEY')
NAME_CREATE_SHEETS  = os.getenv('NAME_CREATE_SHEETS')
NAME_WORKSHEET      = os.getenv('NAME_WORKSHEET')

def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)

def get_scope():
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY, scope)
    customer = gspread.authorize(credentials)
    return customer

def create_google_sheets():
    customer = get_scope()
    sheet = customer.create(NAME_CREATE_SHEETS)
    sheet.share(NAME_GOOGLE_ACCOUNT, perm_type='user', role='writer')

def get_actual_sheet():
    customer = get_scope()
    sheet = customer.open(NAME_CREATE_SHEETS)
    rpaActual = sheet.worksheet(NAME_WORKSHEET)
    return rpaActual

def insert_row(sheet, insertRow, next_row):
    sheet.insert_row(insertRow, int(next_row))