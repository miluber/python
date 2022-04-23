from datetime import datetime
from openpyxl import load_workbook
wb = load_workbook(filename='MyTest.xlsx', data_only=True)
sheet = wb['Sheet1']
i = 0  # this is just for fun to break out of the loop after 5 iterations
for cell in sheet['A']:
    if isinstance(cell.value, datetime):  # Here we focus ony on cells that have dates
        if cell.value.month < 4:
            # usre coordinate to get the cell ID can also do row or column
            print(cell.coordinate + ": " + str(cell.row) + ": " +
                  str(cell.value) + ": " + str(cell.value.month))
            i = i + 1
    else:  # Here we print data type of other cells
        print(cell.coordinate + ": " + str(type(cell.value)))
    if i > 5:
        break
