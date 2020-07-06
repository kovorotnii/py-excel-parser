from openpyxl import load_workbook, workbook, worksheet
from datetime import datetime
import json
import functions
wb = load_workbook('sh1.xlsx')

# target dict
target_dict = {}
# each sheet parsing result will be here
values = []

# get a list of sheets
sheets = functions.get_excel_sheets(wb)

for current_sheet in sheets:
  row_count = wb[current_sheet].max_row
  col_count = wb[current_sheet].max_column

  print(row_count)
  print(col_count)
  print(current_sheet)

  values = functions.parse_excel(wb, current_sheet, col_count, row_count)
  
  if (values):
    target_dict.update({ current_sheet: values })
    values = []
    headers = []
    inter_dict = {}
    counter = 0

with open('output.json', 'w') as output:
    json.dump(target_dict, output, ensure_ascii=False)