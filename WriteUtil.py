import openpyxl

class WriteExcel:
    def __init__(self, fname, data):
        self.data = data
        self.wb = openpyxl.Workbook()
        self.sheet = self.wb.active
        self.fname = fname
    
    def write_excel(self):
        i, j = 1, 1
        for k,v in self.data:
            self.sheet.cell(row=i, column=j).value = k
            self.sheet.cell(row=i, column=j+1).value = v
            i += 1
        self.wb.save("Sortedfile.xlsx")
        
    