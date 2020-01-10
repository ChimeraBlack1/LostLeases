import excelCompare as ec

prev = ec.OpenSheet("DLLPortfolio(Dec).xlsx")
curr = ec.OpenSheet("DLLPortfolio(Jan).xlsx")

serialColumn = 22
rowStart = 1
totalRows1 = ec.FindLastRow(prev)
totalRows2 = ec.FindLastRow(curr)

serial1 = ec.GetValue(prev, rowStart, serialColumn)
serial2 = ec.GetValue(curr, rowStart, serialColumn)
missing = []
matches = []

for i in range (1, totalRows1):
  found = False
  serial1 = ec.GetValue(prev, i, serialColumn)
  print(str(serial1))
  for x in range(1, totalRows2):
    serial2 = ec.GetValue(curr, x, serialColumn)
    if serial1 == serial2:
      found = True
      matches.append(serial1)
      break

  if found == False:
    missing.append(serial1)  
