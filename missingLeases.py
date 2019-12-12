import math
import xlrd
import xlwt

#open PREVIOUS month Sherpa Report
prevReport = "SherpaReport(Oct).xlsm"
prevwb = xlrd.open_workbook(prevReport)
prevSheet = prevwb.sheet_by_index(0)

#open CURRENT month Sherpa Report
currentReport = "SherpaReport(Nov).xlsm"
currentwb = xlrd.open_workbook(currentReport)
currentSheet = currentwb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Leases Lost')
NewWorkbookName = "LeasesLost.xls"

#Globals
endOfPrevMonthSheet = 3499
endOfThisMonthSheet = 3569
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
    newWorkbookPointer = newWorkbookPointer + 1 

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))