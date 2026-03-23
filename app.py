from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import json
from generate_renyu_materials import RenYuMaterialsGenerator
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

# 创建目录
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_materials():
    # 获取上传的JSON文件
    if 'json_file' not in request.files:
        return "请选择JSON文件"
    
    json_file = request.files['json_file']
    if json_file.filename == '':
        return "请选择JSON文件"
    
    # 保存上传的文件
    json_path = os.path.join(app.config['UPLOAD_FOLDER'], json_file.filename)
    json_file.save(json_path)
    
    # 读取JSON数据
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
    except Exception as e:
        return f"读取JSON文件失败: {str(e)}"
    
    # 检查JSON格式
    if isinstance(json_data, dict) and 'courses' in json_data:
        course_data = json_data['courses']
    elif isinstance(json_data, list):
        course_data = json_data
    else:
        return "JSON文件格式不正确，需要包含courses列表"
    
    # 获取教师姓名
    if course_data:
        teacher_name = course_data[0].get('teacherName', '未知教师')
    else:
        teacher_name = '未知教师'
    
    # 创建生成器实例
    generator = RenYuMaterialsGenerator()
    generator.template_dir = '期初资料1'
    
    # 在应用程序目录创建文件夹：姓名_天府学院期初资料
    generator.output_base_dir = os.path.join(app.config['OUTPUT_FOLDER'], f"{teacher_name}_天府学院期初资料")
    
    # 设置课程数据
    generator.courses = course_data
    
    # 设置元数据
    generator.metadata = {
        'date': datetime.now().strftime('%Y年%m月%d日'),
        'semester': '2026年春季学期'
    }
    
    # 打印调试信息
    print(f"📋 app.py - course_data类型: {type(course_data)}")
    print(f"📋 app.py - course_data长度: {len(course_data)}")
    if course_data:
        print(f"📋 app.py - 第一个course类型: {type(course_data[0])}")
        print(f"📋 app.py - 第一个course内容: {course_data[0]}")
    
    # 生成资料
    try:
        generator.generate_all_materials()
        return redirect(url_for('download', teacher_name=teacher_name))
    except Exception as e:
        print(f"❌ 生成失败异常: {str(e)}")
        import traceback
        traceback.print_exc()
        return f"生成资料失败: {str(e)}"

@app.route('/download/<teacher_name>')
def download(teacher_name):
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], f"{teacher_name}_天府学院期初资料")
    
    # 检查目录是否存在
    if not os.path.exists(output_dir):
        return "生成的资料不存在"
    
    # 获取所有文件夹及其文件
    course_data = []
    for item in os.listdir(output_dir):
        item_path = os.path.join(output_dir, item)
        if os.path.isdir(item_path):
            files = os.listdir(item_path)
            course_data.append({
                'folder': item,
                'files': files
            })
    
    return render_template('download.html', teacher_name=teacher_name, course_data=course_data)

@app.route('/download_file/<teacher_name>/<path:folder>/<path:filename>')
def download_file(teacher_name, folder, filename):
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], f"{teacher_name}_天府学院期初资料", folder)
    return send_from_directory(output_dir, filename)

if __name__ == '__main__':
    app.run(debug=True)
