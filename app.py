import openpyxl
from ReadUtil import ReadStore
from WriteUtil import WriteExcel

if __name__ == '__main__':
    
    
    wb = openpyxl.load_workbook("Survey_snowballing.xlsx")
    title_year_col = [('A', 'B'), ('E', 'F')]
    
    
    rs = ReadStore(wb, title_year_col)
    rs.read_and_store()
    sorted_data = rs.data_year_sorted()
    
    we = WriteExcel('sorted-cite-citation', sorted_data)
    we.write_excel()
    
