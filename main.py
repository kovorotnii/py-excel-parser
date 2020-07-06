from openpyxl import load_workbook, workbook, worksheet

wb = load_workbook('sh1.xlsx')

# sheetNames = ['ASOUP', 'loco_1', 'loco_26', 'acts_31L', 
# 'LocoSeries']

# for sheetName in wb.sheetnames[0]:
#   print(wb[sheetName])
# # print(wb.sheetnames)

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
      inter_dict.update({ headers[i]: cell.value })

    if (len(headers) != col_count):
      headers.append(cell.value)
    
    counter += 1

    if (counter == col_count):
      counter = 1
      values.append(inter_dict)
      inter_dict = {}


print(values) 
    