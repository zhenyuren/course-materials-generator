from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os

# 测试添加电子签名图片
def test_add_signature():
    # 读取生成的文件
    file_path = '智能金融学院_未知教师/FIN3004A_大数据分析基础/大数据分析基础-课程组期初教学资料检查情况记录表-任渝.xlsx'
    
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return
    
    # 加载工作簿
    wb = load_workbook(file_path)
    ws = wb.active
    
    # 假设电子签名图片路径（用户需要提供实际的图片文件）
    signature_image_path = 'signature.png'  # 用户需要将签名图片保存为这个文件名
    
    if os.path.exists(signature_image_path):
        # 创建图片对象
        img = Image(signature_image_path)
        
        # 设置图片大小（根据需要调整）
        img.width = 100
        img.height = 40
        
        # 将图片添加到课程组负责人签字后面的位置（第20行第8列附近）
        ws.add_image(img, 'H20')
        
        # 保存文件
        output_path = 'test_with_signature.xlsx'
        wb.save(output_path)
        print(f"✅ 电子签名已添加到: {output_path}")
    else:
        print(f"❌ 电子签名图片不存在: {signature_image_path}")
        print("请将电子签名图片保存为 signature.png 放在当前目录")

if __name__ == '__main__':
    test_add_signature()
    print("\n提示：要自动添加电子签名，请将电子签名图片保存为 signature.png 放在当前目录")