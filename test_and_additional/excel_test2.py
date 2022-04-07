import openpyxl

edrpou = '39808986'

file = 'customers.xlsx'
wb = openpyxl.load_workbook(file, read_only=True)
ws = wb.active
for row in ws.rows:
    for cell in row:
        if cell.value == int(edrpou):
            print(cell.coordinate)
