import math
import xlrd
import xlwt


#open PREVIOUS month Sherpa Report
prevReport = "SherpaReport(Oct).xlsm"
prevwb = xlrd.open_workbook(SherpaReport)
prevSheet = prevwb.sheet_by_index(0)

#open CURRENT month Sherpa Report
currentReport = "SherpaReport(Sep).xlsm"
currentwb = xlrd.open_workbook(SherpaReport)
currentSheet = currentwb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Leases Lost')
NewWorkbookName = "LeasesLost.xls"

#Globals
endOfPrevMonthSheet = 3451
endOfThisMonthSheet = 3499
#Sherpa
SherpaReportSerialCol = 10
SherpaReportAssetPriceCol = 11
SherpaReportCustomerNameCol = 2
SherpaCustomerAddyCol = 12
SherpaReportCustomerModelCol = 9
SherpaReportTestFunderCol = 6

for x in range(1, endOfThisMonthSheet if endOfThisMonthSheet > endOfPrevMonthSheet else endOfPrevMonthSheet):
  # get serial to test from Sherpa Report
  found = False
  try:
    testSerial = currentSheet.cell_value(x,SherpaReportSerialCol)
    testAssetPrice = currentSheet.cell_value(x,SherpaReportAssetPriceCol)
    testCustomerName = currentSheet.cell_value(x,SherpaReportCustomerNameCol)
    testCustomerAddy = currentSheet.cell_value(x,SherpaCustomerAddyCol)
    testCustomerModel = currentSheet.cell_value(x,SherpaReportCustomerModelCol)
    testFunder = currentSheet.cell_value(x,SherpaReportTestFunderCol)
    testSerial = int(testSerial)
  except:
    testSerial = str(testSerial)

  # look in the DLL portfolio for the serial
  for y in range(1, endOfPrevMonthSheet if endOfThisMonthSheet > endOfPrevMonthSheet else endOfThisMonthSheet):
    try:
      otherTestSerial = prevSheet.cell_value(y,SherpaReportSerialCol) if endOfThisMonthSheet > endOfPrevMonthSheet else currentSheet.cell_value(y,SherpaReportSerialCol)
    except:
      continue

    if testSerial == "":
      worksheet.write(x,0,"Blank Serial (DLL)")
      worksheet.write(x,1, testAssetPrice)
      worksheet.write(x,2, testCustomerName)
      worksheet.write(x,3, testCustomerModel)
      worksheet.write(x,4, testCustomerAddy)
      continue
    if testSerial == otherTestSerial:
      found = True
      break
 
  #if found, go to next item
  if found == True: 
    continue
  
  if found == False:
    worksheet.write(x,0, testSerial)
    worksheet.write(x,1, testAssetPrice)
    worksheet.write(x,2, testCustomerName)
    worksheet.write(x,3, testCustomerModel)
    worksheet.write(x,4, testCustomerAddy)
    worksheet.write(x,5, testFunder)

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))