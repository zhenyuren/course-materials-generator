from openpyxl import load_workbook

# 加载模板文件
template_file = "期初资料/附件1 教学大纲审核汇总表-完全一致模板.xlsx"
wb = load_workbook(template_file)
ws = wb.active

# 获取所有合并单元格
merged_cells = list(ws.merged_cells.ranges)

print("合并单元格列表：")
for merged_range in merged_cells:
    print(f"{merged_range}")
