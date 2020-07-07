from datetime import datetime
from dateparser import parse

def isAccessible(path, mode='r'):
    """ check if file is accessbile """
    try:
        f = open(path, mode)
        f.close()
    except IOError:
        return False
    return True

def get_excel_sheets(wb):
  return wb.sheetnames

def parse_excel(wb, current_sheet, col_count, row_count):
  # an array which stores column headers
  headers = []
  # an array contains each row as object 
  values = []
  # processed columns counter
  counter = 0
  # sores row as an object with columns as keys
  inter_dict = {}
  for row in wb[current_sheet].iter_rows(min_row=0, max_col=col_count, max_row=row_count):
    for i, cell in enumerate(row):
      counter += 1
      if (len(headers) == col_count):
        # try to parse date as string to format
        if ('DATE' in headers[i] or 'TIME' in headers[i]):
          if (cell.value is None):
            inter_dict.update({ headers[i]: None })
          # cell value not none
          if (cell.value):
            try:
              date = datetime.strftime(cell.value, '%Y-%m-%dT%H:%M:%SZ')
              inter_dict.update({ headers[i]:  date })
            except TypeError:
              new_date = parse(cell.value, date_formats=['%Y-%m-%d'])
              if ('DATE' in headers[i]):
                inter_dict.update({ headers[i]: datetime.strftime(new_date, '%Y-%m-%d') })
              if ('TIME' in headers[i]):
                inter_dict.update({ headers[i]: cell.value })
        else:
          # if cell.value is None:
          #   inter_dict.update({ headers[i]: cell.value })
          # if cell.value:
          inter_dict.update({ headers[i]: cell.value })
      if (len(headers) != col_count):
        headers.append(cell.value)
      if (counter == col_count):
        counter = 0
        if (inter_dict):
          values.append(inter_dict)
          print(values)
          inter_dict = {}
  
  return values