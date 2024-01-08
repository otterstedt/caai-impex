#!venv/bin/python3
# read contacts from xsl files
# write to contact manager software

import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import json

wb = load_workbook(filename="./files/alumni_list_gbg_20231020.xlsx")

print("Sheets")
for sheet in wb:
  print("'" + sheet.title + "'")
  full_list=[]

  for row in sheet.iter_rows(min_row=2):
    drow={}
    for cell in row:
      if sheet.cell(row=1, column=cell.column).value and cell.value:
        drow[sheet.cell(row=1, column=cell.column).value]=str(cell.value)
    full_list.append(drow)
    #print(json.dumps(drow, indent=2))

print(json.dumps(full_list, indent=2))

