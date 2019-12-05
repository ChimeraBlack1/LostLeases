import math
import xlrd
import xlwt

#open wells workbook
WellsPortfolio = "WellsPortfolio(Oct2019).xlsx"
Wellswb = xlrd.open_workbook(WellsPortfolio)
WellsSheet = Wellswb.sheet_by_index(0)

#open DLL workbook
DLLPortfolio = "DLLPortfolio(Oct2019).xlsx"
DLLwb = xlrd.open_workbook(DLLPortfolio)
DLLSheet = DLLwb.sheet_by_index(0)

#open report for perry (Sherpa report)
SherpaReport = "SherpaReport.xlsm"
Sherpawb = xlrd.open_workbook(SherpaReport)
SherpaSheet = Sherpawb.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Leases Lost')
NewWorkbookName = "LeasesLost.xls"

#Globals
endOfSherpaSheet = 3499
endOfWells = 2152
endOfDLL = 1464
lostLeases = 0
#DLL
DLLAssetPriceColumn = 21
DLLSerialColumn = 22
DLLcustomerName = 5
DLLcustomerAddy = 6
DLLcustomerModel = 24
#wells
wellsAssetPriceColumn = 6
wellsSerialColumn = 5
wellscustomerName = 11
wellscustomerAddy = 20
wellscustomerModel = 3
#Sherpa
SherpaReportSerialCol = 10
SherpaReportAssetPriceCol = 11
SherpaReportCustomerNameCol = 2
SherpaCustomerAddyCol = 12
SherpaReportCustomerModelCol = 9
SherpaReportTestFunderCol = 6

for x in range(1, endOfSherpaSheet):
  # get serial to test from Sherpa Report
  found = False
  try:
    testSerial = SherpaSheet.cell_value(x,SherpaReportSerialCol)
    testAssetPrice = SherpaSheet.cell_value(x,SherpaReportAssetPriceCol)
    testCustomerName = SherpaSheet.cell_value(x,SherpaReportCustomerNameCol)
    testCustomerAddy = SherpaSheet.cell_value(x,SherpaCustomerAddyCol)
    testCustomerModel = SherpaSheet.cell_value(x,SherpaReportCustomerModelCol)
    testFunder = SherpaSheet.cell_value(x,SherpaReportTestFunderCol)
    testSerial = int(testSerial)
  except:
    testSerial = str(testSerial)

  # look in the DLL portfolio for the serial
  for y in range(1, endOfDLL):
    try:
      DLLAssetPrice = DLLSheet.cell_value(y,DLLAssetPriceColumn)
      DLLserial = DLLSheet.cell_value(y,DLLSerialColumn)
    except:
      continue

    if testSerial == "":
      worksheet.write(x,0,"Blank Serial (DLL)")
      worksheet.write(x,1, DLLAssetPrice)
      worksheet.write(x,2, DLLcustomerName)
      worksheet.write(x,3, DLLcustomerModel)
      worksheet.write(x,4, DLLcustomerAddy)
      continue
    if testSerial == DLLserial:
      found = True
      break
 
  #if found, go to next item
  if found == True: 
    continue

  # else, look in the wells portfolio for the serial
  for y in range(1, endOfWells):
    
    try:
      wellsAssetPrice = WellsSheet.cell_value(y, WellsAssetPriceColumn)
      wellsSerial = WellsSheet.cell_value(y, WellsSerialColumn)
    except:
      continue

    if testSerial == "":
      print("Writing to wb (wells)...")
      worksheet.write(x,0,"Blank Serial (wells)")
      worksheet.write(x,1, testAssetPrice)
      worksheet.write(x,2, testCustomerName)
      worksheet.write(x,3, testCustomerModel)
      worksheet.write(x,4, testCustomerAddy)
      continue
    if testSerial == wellsSerial:
      found = True
      break
  
  if found == False:
    worksheet.write(x,0, testSerial)
    worksheet.write(x,1, testAssetPrice)
    worksheet.write(x,2, testCustomerName)
    worksheet.write(x,3, testCustomerModel)
    worksheet.write(x,4, testCustomerAddy)
    worksheet.write(x,5, testFunder)

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))
#LOST LEASES
#if serialToTest is not blank:
  #for each serial number check DLL, then Wells to see if it's there
    #if you find a match, continue
    #if no match is found, write details to 'Lost Machines' and save

#NEW LEASES
#if serialToTest is not blank:
 #check all serials in DLL Portfolio against serials in Report for Perry
   #if no match is found, check Wells
   #check all serials in Wells portfolio against serials in Report for Perry
   #if no match is found, add to New Leases. 
