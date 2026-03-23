import xlrd
from openpyxl import Workbook

# 使用xlrd读取原始xls文件
workbook = xlrd.open_workbook('期初资料/附件3 课程组教学资料检查情况记录表.xls')
sheet = workbook.sheet_by_index(0)

print('原始xls文件内容:')
print('=' * 80)

# 打印内容
for row_idx in range(sheet.nrows):
    row_values = []
    for col_idx in range(sheet.ncols):
        cell_value = sheet.cell_value(row_idx, col_idx)
        row_values.append(str(cell_value))
    print(f'第{row_idx+1}行: ' + ' | '.join(row_values))

# 检查单元格格式
print('\n单元格格式信息:')
for row_idx in range(min(5, sheet.nrows)):
    for col_idx in range(min(10, sheet.ncols)):
        cell_type = sheet.cell_type(row_idx, col_idx)
        cell_value = sheet.cell_value(row_idx, col_idx)
        print(f'第{row_idx+1}行第{col_idx+1}列: 类型={cell_type}, 值={cell_value}')

# 检查合并单元格
print('\n合并单元格信息:')
if hasattr(sheet, 'merged_cells'):
    for merged_cell in sheet.merged_cells:
        print(f'合并单元格: {merged_cell}')
else:
    print('xlrd不支持合并单元格信息')