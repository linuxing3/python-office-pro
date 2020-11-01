import xlrd
import xlwt

file_path = 'd:\workspace\python-office-hallow\sample.xlsx'
tem_excel = xlrd.open_workbook(file_path, formatting_info=True)
# tem_sheet = tem_excel.sheet_by_index(0)
tem_sheet = tem_excel.sheet_by_name('data')
# Write with style

for row in range(10):
  print(tem_sheet.cell_value(0, row))
  
# other manners to read cell value
# print(tem_sheet.cell(0, 3).value)
# print(tem_sheet.row(1)[2].value)
