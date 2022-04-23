from openpyxl import Workbook
# create new workbook
wb = Workbook()
# open active worksheet
ws = wb.active
wb.create_sheet('Iwona')

for sheet in wb:
    print(sheet.title)
# wrie to some wb make sure it is closed first
wb.save('ivona3.xlsx')
