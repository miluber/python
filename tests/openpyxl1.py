from openpyxl import Workbook
# create new workbook
wb = Workbook()
# open active worksheet
ws = wb.active
wb.create_sheet('Iwona')

for sheet in wb:
    print(sheet.title)
wb.save('ivona3.xlsx')
