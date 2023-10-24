#!venv/bin/python3
# read contacts from xsl files
# write to contact manager software

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

wb = load_workbook(filename="./files/alumni_list_gbg_20231020.xlsx")

print("Sheets")
for sheet in wb:
  print("'" + sheet.title + "'")

