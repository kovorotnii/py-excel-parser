from openpyxl import load_workbook, workbook, worksheet
from datetime import datetime
import json

wb = load_workbook('sh1.xlsx')

inter_dict = {}
target_dict = {}
headers = []
values = []
row_count = wb['Sheet1'].max_row
col_count = wb['Sheet1'].max_column
counter = 1

for row in wb['Sheet1'].iter_rows(min_row=0, max_col=col_count, max_row=row_count):
  for i, cell in enumerate(row):
    if (len(headers) == col_count):
      # try to parse date as string to format
      if ('DATE' in headers[i] or 'TIME' in headers[i]):
        if (not cell.value):
          date = ''
          inter_dict.update({ headers[i]:  date })
        # cell value not none
        if (cell.value):
          date = datetime.strftime(cell.value, '%Y-%m-%dT%H:%M:%SZ')
          inter_dict.update({ headers[i]:  date })
      else:
        inter_dict.update({ headers[i]: cell.value })

    if (len(headers) != col_count):
      headers.append(cell.value)
    
    counter += 1
    if (counter == col_count):
      counter = 1
      if (inter_dict):
        values.append(inter_dict)
        inter_dict = {}

if (values):
  target_dict.update({ 'Sheet1': values })
  print(target_dict)
  # with open('output.json', 'w') as output:
  #  json.dump(target_dict, output)
