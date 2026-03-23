import os
from win32com.client import Dispatch
import pythoncom

def convert_xls_to_xlsx(xls_path, xlsx_path):
    """使用Excel COM对象转换xls到xlsx，保持所有格式"""
    pythoncom.CoInitialize()
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        workbook = excel.Workbooks.Open(xls_path)
        workbook.SaveAs(xlsx_path, FileFormat=51)  # 51 is xlsx format
        workbook.Close()
        print(f"✅ 成功转换: {xls_path} -> {xlsx_path}")
        return True
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        return False
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

# 转换原始xls文件
xls_file = '期初资料/附件3 课程组教学资料检查情况记录表.xls'
xlsx_file = '期初资料/附件3 课程组教学资料检查情况记录表-原始格式.xlsx'

if os.path.exists(xls_file):
    convert_xls_to_xlsx(xls_file, xlsx_file)
else:
    print(f"❌ 文件不存在: {xls_file}")