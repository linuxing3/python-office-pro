import xlwt

file_path = 'd:\workspace\python-office-hallow\sample.xlsx'
new_workbook = xlwt.Workbook()

data_sheet = new_workbook.add_sheet('data')

for row in range(10):
  for col in range(10):
    data_sheet.write(row, col, '数据'+ str(row) + '---' + str(col))

new_workbook.save(file_path)