import excelCompare as ec

prev = ec.OpenSheet("Wells(Nov).xlsx")
curr = ec.OpenSheet("Wells(Dec).xlsx")
newWBName = "Missing Leases Dec 2019"

serialColumn1 = 2
serialColumn2 = 4
priceCol = 4
customerCol = 8
modelCol = 3
addressCol = 11
rowStart = 1
totalRows1 = ec.FindLastRow(prev)
totalRows2 = ec.FindLastRow(curr)

newb = ec.Newb()
news = ec.News(newb, "Missing")

#Titles
news.write(0, 0, "Serial Number")
news.write(0, 1, "Asset Price")
news.write(0, 2, "Customer Name")
news.write(0, 3, "Model Name")
news.write(0, 4, "Address")

totalFound = 1

for i in range (1, totalRows1):
  found = False
  serial1 = ec.GetValue(prev, i, serialColumn1)
  price = ec.GetValue(prev, i, priceCol)
  customer = ec.GetValue(prev, i, customerCol)
  model = ec.GetValue(prev, i, modelCol)
  address = ec.GetValue(prev, i, addressCol)

  for x in range(1, totalRows2):
    serial2 = ec.GetValue(curr, x, serialColumn2)
    if serial1 == serial2:
      found = True
      break

  if found == False:
    news.write(totalFound, 0, serial1)
    news.write(totalFound, 1, price)
    news.write(totalFound, 2, customer)
    news.write(totalFound, 3, model)
    news.write(totalFound, 4, address)
    totalFound = totalFound + 1

ec.Save(newb, newWBName)

#write to workbook
# workbook = xlwt.Workbook()
# worksheet = workbook.add_sheet('Leases Lost')
# NewWorkbookName = "LeasesLost.xls
# workbook.save(NewWorkbookName)