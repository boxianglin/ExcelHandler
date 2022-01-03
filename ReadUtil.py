from collections import defaultdict

class ReadStore:
    def __init__(self, wb, title_year_col):
        self.wb = wb
        self.title_year_col = title_year_col
        self.storage = defaultdict(int)
    
    def read_and_store(self):
        sheet = self.wb.active
        for row in range(2, sheet.max_row+1):
            for t, y in self.title_year_col:
                title = sheet[t + str(row)].value
                year = sheet[y + str(row)].value
                if title:
                    if not year: year = 0
                    self.storage[title.strip()] = int(year)
    
    def data_year_sorted(self):
        data = list(self.storage.items())
        # sorted by name then by year
        return sorted(sorted(data), key=lambda x:x[1], reverse=True)