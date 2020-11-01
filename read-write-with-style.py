from xlutils.copy import copy
import xlrd
import xlwt

sample_file_path = 'd:\workspace\python-office-hallow\sample.xlsx'
# work_file_path = 'd:\workspace\python-office-hallow\homework.xlsx'

tem_excel = xlrd.open_workbook(sample_file_path, formatting_info=True)
tem_sheet = tem_excel.sheet_by_index(0)

new_excel = copy(tem_excel)
new_sheet = new_excel.get_sheet(0)

style = xlwt.XFStyle()
# font
font = xlwt.Font()
font.name = 'yahei'
font.bold = True
# 260 = 20 * 18 
font.height = 260
style.font = font

# bolder
borders = xlwt.Borders()
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN
borders.left = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
style.borders = borders

# alignment
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER
style.alignment = alignment

# Write with style
new_sheet.write(2, 1, 12, style)
new_excel.save()