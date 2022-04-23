from datetime import datetime
from openpyxl import load_workbook
wb = load_workbook(filename='MyTest.xlsx', data_only=True)
sheet = wb['Sheet1']
i = 0  # this is just for fun to break out of the loop after 5 iterations
for cell in sheet['A']:
    if isinstance(cell.value, datetime):  # Here we focus ony on cells that have dates
        if cell.value.month < 4 and cell.value.year == 2012:
            # usre coordinate to get the cell ID can also do row or column
            thisCell = cell.coordinate
            thisRow = str(cell.row)
            nextCell = 'K' + thisRow
            print(thisCell + ": " +
                  str(cell.value) + ": " + str(cell.value.month))
            cellRange = sheet[thisCell:nextCell]
            for cll in cellRange:
                for cll2 in cll:
                    print(cll2.coordinate + ":" + str(cll2.value), end=" ")
            i = i + 1
            print()

    if i > 0:
        break
