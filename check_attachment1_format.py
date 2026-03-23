import xlrd

# 读取附录1的xls文件
file_path = '期初资料/附件1 教学大纲的审核汇总表-大数据信息分析_任渝.xls'
wb = xlrd.open_workbook(file_path)
ws = wb.sheet_by_index(0)

print('附录1格式检查:')
print('=' * 80)

# 检查基本信息
print(f'工作表名称: {ws.name}')
print(f'行数: {ws.nrows}')
print(f'列数: {ws.ncols}')

# 检查前20行的内容
print('\n前20行内容:')
for row in range(min(20, ws.nrows)):
    row_data = []
    for col in range(ws.ncols):
        cell_value = ws.cell_value(row, col)
        if cell_value != '':
            row_data.append(f'列{col+1}: {cell_value}')
    if row_data:
        print(f'第{row+1}行: {", ".join(row_data)}')

# 注意：xlrd无法直接获取合并单元格信息，需要通过其他方式识别

print('\n格式检查完成！')