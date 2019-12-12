import math
import xlrd
import xlwt

def Device_Speed(speed):
  device_speed = {
    1: "Low",
    2: "Medium",
    3: "High",
    4: "Unknown"
  }
  return device_speed.get(speed)

# Take a filename as Input.  If it doesn't work, try again until it does work.
goodFile = False

while goodFile == False:
  fileToRead = input("Please enter the name of the PREVIOUS month's report)> ")
  if fileToRead == "":
    prevReport = ("ProdMAPP.xlsx")
    goodFile = True
    wb = xlrd.open_workbook(prevReport)
  elif fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    prevReport = fileToRead + ".xlsm"
    try:
      prevwb = xlrd.open_workbook(prevReport)
      prevSheet = prevwb.sheet_by_index(0)
      goodFile = True
    except:
      print("I can't find that file, try again...")

#open PREVIOUS month Sherpa Report
# prevReport = "SherpaReport(Oct).xlsm"
# prevwb = xlrd.open_workbook(prevReport)
# prevSheet = prevwb.sheet_by_index(0)

# Take a filename as Input.  If it doesn't work, try again until it does work.
goodReport = False

while goodReport == False:
  fileToRead2 = input("Please enter the name of the THIS month's report)> ")
  if fileToRead2 == "":
    currentReport = ("ProdMAPP.xlsx")
    goodReport = True
    currentwb = xlrd.open_workbook(currentReport)
  elif fileToRead2 == "exit" or fileToRead2 == "quit":
    print("ok, bye!")
    exit()
  else:
    currentReport = fileToRead2 + ".xlsm"
    try:
      currentwb = xlrd.open_workbook(currentReport)
      currentSheet = currentwb.sheet_by_index(0)
      goodReport = True
    except:
      print("I can't find that file, try again...")

#open CURRENT month Sherpa Report
# currentReport = "SherpaReport(Nov).xlsm"
# currentwb = xlrd.open_workbook(currentReport)
# currentSheet = currentwb.sheet_by_index(0)

# Take a filename as Input.  If it doesn't work, try again until it does work.
PrevMonthSheetNumberGood = False

while PrevMonthSheetNumberGood == False:
  endOfPrevMonthSheetInput = input("How many cells are in the PREVIOUS month's sheet)> ")
  if endOfPrevMonthSheetInput == "exit" or endOfPrevMonthSheetInput == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endOfPrevMonthSheet = endOfPrevMonthSheetInput
      PrevMonthSheetNumberGood = True
    except:
      print("Not the right number, try again...)>")


# Take a filename as Input.  If it doesn't work, try again until it does work.
ThisMonthSheetNumberGood = False

while ThisMonthSheetNumberGood == False:
  endOfThisMonthSheetInput = input("How many cells are in the THIS month's sheet)> ")
  if endOfThisMonthSheetInput == "exit" or endOfThisMonthSheetInput == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endOfThisMonthSheet = endOfThisMonthSheetInput
      ThisMonthSheetNumberGood = True
    except:
      print("Not the right number, try again...)>")

# Take a filename as Input.  If it doesn't work, try again until it does work.
goodMonth = False

while goodMonth == False:
  endOfThisMonthSheetInput = input("What Month is this?)> ")
  if endOfThisMonthSheetInput == "exit" or endOfThisMonthSheetInput == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endOfThisMonthSheet = endOfThisMonthSheetInput
      goodMonth = True
    except:
      print("Not the right number, try again...)>")

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Leases Lost')
NewWorkbookName = "LeasesLost.xls"

#Globals
#endOfPrevMonthSheet = 3499
#endOfThisMonthSheet = 3569
newWorkbookPointer = 0
#Sherpa
SherpaReportAssetVolType = 8
SherpaReportSerialCol = 10
SherpaReportAssetPriceCol = 11
SherpaReportCustomerNameCol = 2
SherpaCustomerAddyCol = 12
SherpaReportCustomerModelCol = 9
SherpaReportTestFunderCol = 6
newWorkbookPointer = 0

if(endOfThisMonthSheet < endOfPrevMonthSheet):
  print("careful, this month has less total assets than last month. Check the code.")

#LOST
for x in range(1, endOfPrevMonthSheet):

  if endOfThisMonthSheet > endOfPrevMonthSheet:
    if x > endOfPrevMonthSheet:
      break
  elif endOfPrevMonthSheet > endOfThisMonthSheet:
    if x > endOfThisMonthSheet:
      break
  found = False
  # get serial to test from Sherpa Report
  try:
    testSerial = prevSheet.cell_value(x,SherpaReportSerialCol)
    testAssetPrice = prevSheet.cell_value(x,SherpaReportAssetPriceCol)
    testCustomerName = prevSheet.cell_value(x,SherpaReportCustomerNameCol)
    testCustomerAddy = prevSheet.cell_value(x,SherpaCustomerAddyCol)
    testCustomerModel = prevSheet.cell_value(x,SherpaReportCustomerModelCol)
    testFunder = prevSheet.cell_value(x,SherpaReportTestFunderCol)
    testAssetVol = prevSheet.cell_value(x,SherpaReportAssetVolType)
    testSerial = int(testSerial)
  except:
    testSerial = str(testSerial)

  # look in the other report for the serial
  for y in range(1, endOfThisMonthSheet):
    try:
      otherTestSerial = currentSheet.cell_value(y,SherpaReportSerialCol)
    except:
      continue

    if testSerial == "":
      worksheet.write(newWorkbookPointer, 0,"Blank Serial")
      worksheet.write(newWorkbookPointer, 1, testAssetPrice)
      worksheet.write(newWorkbookPointer, 2, testCustomerName)
      worksheet.write(newWorkbookPointer, 3, testCustomerModel)
      worksheet.write(newWorkbookPointer, 4, testCustomerAddy)
      worksheet.write(newWorkbookPointer, 5, testFunder)
      worksheet.write(newWorkbookPointer, 6, testAssetVol)
      newWorkbookPointer = newWorkbookPointer + 1
      continue
    if testSerial == otherTestSerial:
      found = True
      break
 
  #if found, go to next item
  if found == True:
    continue
  
  if found == False:
    worksheet.write(newWorkbookPointer, 0, testSerial)
    worksheet.write(newWorkbookPointer, 1, testAssetPrice)
    worksheet.write(newWorkbookPointer, 2, testCustomerName)
    worksheet.write(newWorkbookPointer, 3, testCustomerModel)
    worksheet.write(newWorkbookPointer, 4, testCustomerAddy)
    worksheet.write(newWorkbookPointer, 5, testFunder)
    worksheet.write(newWorkbookPointer, 6, testAssetVol)
    worksheet.write(newWorkbookPointer, 7, Device_Speed(testAssetVol))
    newWorkbookPointer = newWorkbookPointer + 1 


workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))