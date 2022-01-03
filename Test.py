import openpyxl
wb = openpyxl.load_workbook("testtt.xlsx")
sheet = wb.active
for row in range(2, sheet.max_row+1):
    title = sheet['A' + str(row)].value
    print(title)
    print(sheet['A' + str(row)].font.color.rgb)