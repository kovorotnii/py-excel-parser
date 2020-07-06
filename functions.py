
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
