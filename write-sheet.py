import xlwt

file_path = 'd:\workspace\python-office-hallow\sample.xlsx'
new_workbook = xlwt.Workbook()

data_sheet = new_workbook.add_sheet('data')

for row in range(10):
  data_sheet.write(0, row, '数据'+ str(row))

new_workbook.save(file_path)