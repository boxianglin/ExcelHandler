from collections import defaultdict
from openpyxl.styles import fonts
from openpyxl.styles import colors

'''
Read the xlsx and store the data into a dictionary.

Dictionary format: {title: [year, font_color], title2:[year, font_color] ..... }
'''
class ReadStore:
    def __init__(self, wb, title_year_col):
        self.wb = wb
        self.title_year_col = title_year_col
        self.storage = defaultdict(list)
    
    '''
    Read the data and store into a dict
    '''
    def read_and_store(self):
        sheet = self.wb.get_sheet_by_name("Collected")
        for row in range(2, sheet.max_row+1):
            for t, y in self.title_year_col:
                
                title_cell = sheet[t + str(row)]
                year_cell = sheet[y + str(row)]
                
                title = title_cell.value
                year = year_cell.value
                
                # interested only for such cell with titles
                if title:
                    font_color = None
                    # record the font color
                    if title_cell.font.color and type(title_cell.font.color.rgb) == str:
                        font_color = title_cell.font.color.rgb
                        
                    if not year: year = 0
                    self.storage[title.strip()] = [int(year), font_color]
    
    '''
    Sort the data by alphabet then by year
    '''
    def data_year_sorted(self):
        data = list(self.storage.items())
        return sorted(sorted(data), key=lambda x:x[1][0], reverse=True)