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

#LOST LEASES
#open report for perry
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
