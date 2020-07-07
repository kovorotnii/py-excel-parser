from openpyxl import load_workbook, workbook, worksheet
from datetime import datetime
import json
import functions
wb = load_workbook('data.xlsm', data_only=True)

# target dict
target_dict = {}
# each sheet parsing result will be here
values = []

# get a list of sheets
sheets = functions.get_excel_sheets(wb)

for current_sheet in sheets:
  row_count = wb[current_sheet].max_row
  col_count = wb[current_sheet].max_column

  print('Parsing sheet: ', current_sheet)

  values = functions.parse_excel(wb, current_sheet, col_count, row_count)
  
  if not values:
    print("Nothing to assign to target object!")
    quit()
  # Get data from parser, assign to dict
  target_dict.update({ current_sheet: values })
  values = []
    
try:
  with open('output.json', 'w') as output:
    data = json.dumps(target_dict, ensure_ascii=False)
    output.write(data)
    output.close()
  
  print('Parsing completed!')
except IOError as e:
  print('Writing to file error! ', e)
  quit()