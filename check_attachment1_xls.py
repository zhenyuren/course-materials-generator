import xlrd

# 打开附件1.xls文件
wb = xlrd.open_workbook('期初资料/附件1 教学大纲的审核汇总表-大数据信息分析_任渝.xls')
sheet = wb.sheet_by_index(0)

print(f'行数: {sheet.nrows}, 列数: {sheet.ncols}')
print('前20行内容:')
for i in range(min(20, sheet.nrows)):
    values = sheet.row_values(i)
    # 过滤掉空值
    non_empty_values = [v for v in values if v != '']
    if non_empty_values:
        print(f'行{i+1}: {non_empty_values}')
