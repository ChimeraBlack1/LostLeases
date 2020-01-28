import math
import xlrd
import xlwt

def OpenSheet(wb, tab=0):
  """
  opens a workbook and returns the first sheet
  """
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(tab)
  return sheet

def GetValue(sheet, row=0, col=0):
  """
  Get the value from the cell in a specified sheet
  """
  serial = sheet.cell_value(row,col)
  return serial

def FindLastRow(sheet, row=0, col=0):
  """
  Find the number of populated rows in an excel workbook
  """  
  content = sheet.cell_value(row, col)
  rowCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row + rowCount, col)
    except:
      break
    rowCount = rowCount + 1
    
  return rowCount


def FindLastRowZeroIndex(sheet, row=0, col=0):
  """
  Find last populated row from starting point, returns 0 index value
  """
  content = sheet.cell_value(row, col)
  rowCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row + rowCount, col)
    except:
      break
    rowCount = rowCount + 1
  
  #zero indexing
  rowCount = rowCount - 1
  if rowCount < 0:
    rowCount = 0

  return rowCount

def FindLastCol(sheet, row=0, col=0):
  """
  Find last populated column from starting point
  """
  content = sheet.cell_value(row, col)
  colCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row, col + colCount)
    except:
      break
    colCount = colCount + 1
  return colCount

def FindLastColZeroIndex(sheet, row=0, col=0):
  """
  Find last populated column from starting point, returns 0 index value
  """
  content = sheet.cell_value(row, col)
  colCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row, col + colCount)
    except:
      break
    colCount = colCount + 1
  
  #zero indexing
  colCount = colCount - 1
  if colCount < 0:
    colCount = 0

  return colCount

def GetStatusDetails(wb, row=0, col=0):
  """
  Collect Status Details from Manager Workbook
  """
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(0)
  status = sheet.cell_value(row, col)
  notes = sheet.cell_value(row, col + 1)
  renewalDate = sheet.cell_value(row, col + 2)
  acctStatus = {
    "status": status,
    "notes": notes,
    "renewalDate": renewalDate,
  }
  return acctStatus

def Newb():
  """
  Create and return a new workbook object
  """
  workbook = xlwt.Workbook()
  return workbook

def News(workbook, name="New Sheet"):
  worksheet = workbook.add_sheet(name)
  return worksheet


def Save(wb, NewWorkbookName="New Workbook.xlsx"):
  """
  Save the output in a workbook
  """
  NewWorkbookName = str(NewWorkbookName + ".xls")
  wb.save(NewWorkbookName)
  print("Saved: " + str(NewWorkbookName))


def FindMissing(sheet1, sheet2):
  """
  Find missing data from sheet 1 in sheet 2
  """
  return



