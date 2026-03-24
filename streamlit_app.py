import streamlit as st
import json
import os
import zipfile
import urllib.parse
from generate_renyu_materials import RenYuMaterialsGenerator
from datetime import datetime
from flask import Flask, request, jsonify
import threading
import traceback

# 创建Flask应用实例
flask_app = Flask(__name__)

# 设置页面标题和布局
st.set_page_config(page_title="课程期初资料生成工具", layout="wide")

# 获取URL参数
query_params = st.query_params
json_data_param = query_params.get('json_data', None)

# 页面标题
st.title("📚 课程期初资料生成工具")

# 侧边栏配置
with st.sidebar:
    st.subheader("配置选项")
    template_dir = st.text_input("模板目录", value="期初资料1")
    output_dir = st.text_input("输出目录", value="output")

# 主内容区
st.markdown("---")

# 方式1：从URL参数获取JSON（智能体调用）
if json_data_param:
    try:
        st.info("📥 接收到智能体传递的课程数据")
        # URL解码
        json_str = urllib.parse.unquote(json_data_param)
        course_data = json.loads(json_str)
        
        # 合并相同课程名称的课程信息
        merged_courses = {}
        for course in course_data:
            course_name = course.get('courseName', '')
            if course_name not in merged_courses:
                merged_courses[course_name] = course.copy()
        
        # 显示课程信息
        st.subheader("📋 课程信息")
        for i, (course_name, course) in enumerate(merged_courses.items(), 1):
            with st.expander(f"课程 {i}: {course_name}"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**课程代码**: {course.get('courseCode', '')}")
                    st.write(f"**教师姓名**: {course.get('teacherName', '')}")
                    st.write(f"**开课单位**: {course.get('department', '')}")
                with col2:
                    st.write(f"**课程性质**: {course.get('courseNature', '')}")
                    st.write(f"**学分**: {course.get('credits', '')}")
                    st.write(f"**总学时**: {course.get('totalHours', '')}")
        
        # 自动生成按钮
        if st.button("🚀 生成课程资料", type="primary", key="auto_generate"):
            with st.spinner("⏳ 正在生成资料，请稍候..."):
                try:
                    # 创建生成器实例
                    generator = RenYuMaterialsGenerator()
                    generator.template_dir = template_dir
                    generator.output_base_dir = output_dir
                    
                    # 设置课程数据（使用合并后的课程数据）
                    generator.courses = list(merged_courses.values())
                    
                    # 设置元数据
                    generator.metadata = {
                        'date': datetime.now().strftime('%Y年%m月%d日'),
                        'semester': '2026年春季学期'
                    }
                    
                    # 生成资料
                    generator.generate_all_materials()
                    
                    # 获取生成的文件
                    teacher_name = course_data[0].get('teacherName', '未知教师')
                    output_folder = os.path.join(generator.output_base_dir, f"{teacher_name}_天府学院期初资料")
                    
                    # 创建ZIP文件
                    zip_filename = f"{teacher_name}_课程资料.zip"
                    zip_path = os.path.join(output_dir, zip_filename)
                    
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for root, dirs, files in os.walk(output_folder):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, output_dir)
                                zipf.write(file_path, arcname)
                    
                    # 显示成功信息
                    st.success(f"✅ 资料生成成功！共生成 {len(course_data)} 门课程的资料")
                    
                    # 提供下载
                    st.subheader("📥 下载生成的资料")
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label=f"下载 {zip_filename}",
                            data=f,
                            file_name=zip_filename,
                            mime="application/zip",
                            key="download_auto"
                        )
                    
                except Exception as e:
                    st.error(f"❌ 生成失败: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
    
    except Exception as e:
        st.error(f"❌ JSON数据解析失败: {str(e)}")

# 分隔线
st.markdown("---")

# 方式2：手动上传（原有功能）
st.subheader("📁 手动上传JSON文件")
uploaded_file = st.file_uploader("选择JSON文件", type="json")

if uploaded_file is not None:
    # 读取JSON数据
    try:
        json_data = json.load(uploaded_file)
        
        # 检查JSON格式
        if isinstance(json_data, dict) and 'courses' in json_data:
            course_data = json_data['courses']
        elif isinstance(json_data, list):
            course_data = json_data
        else:
            st.error("❌ JSON文件格式不正确，需要包含courses列表")
            st.stop()
        
        # 合并相同课程名称的课程信息
        merged_courses = {}
        for course in course_data:
            course_name = course.get('courseName', '')
            if course_name not in merged_courses:
                merged_courses[course_name] = course.copy()
        
        # 显示课程信息
        st.subheader("📋 课程信息")
        for i, (course_name, course) in enumerate(merged_courses.items(), 1):
            with st.expander(f"课程 {i}: {course_name}"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**课程代码**: {course.get('courseCode', '')}")
                    st.write(f"**教师姓名**: {course.get('teacherName', '')}")
                    st.write(f"**开课单位**: {course.get('department', '')}")
                with col2:
                    st.write(f"**课程性质**: {course.get('courseNature', '')}")
                    st.write(f"**学分**: {course.get('credits', '')}")
                    st.write(f"**总学时**: {course.get('totalHours', '')}")
        
        # 生成按钮
        if st.button("🚀 生成课程资料", type="primary", key="manual_generate"):
            with st.spinner("⏳ 正在生成资料，请稍候..."):
                try:
                    # 创建生成器实例
                    generator = RenYuMaterialsGenerator()
                    generator.template_dir = template_dir
                    generator.output_base_dir = output_dir
                    
                    # 设置课程数据（使用合并后的课程数据）
                    generator.courses = list(merged_courses.values())
                    
                    # 设置元数据
                    generator.metadata = {
                        'date': datetime.now().strftime('%Y年%m月%d日'),
                        'semester': '2026年春季学期'
                    }
                    
                    # 生成资料
                    generator.generate_all_materials()
                    
                    # 获取生成的文件
                    teacher_name = course_data[0].get('teacherName', '未知教师')
                    output_folder = os.path.join(generator.output_base_dir, f"{teacher_name}_天府学院期初资料")
                    
                    # 创建ZIP文件
                    zip_filename = f"{teacher_name}_课程资料.zip"
                    zip_path = os.path.join(output_dir, zip_filename)
                    
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for root, dirs, files in os.walk(output_folder):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, output_dir)
                                zipf.write(file_path, arcname)
                    
                    # 显示成功信息
                    st.success(f"✅ 资料生成成功！共生成 {len(course_data)} 门课程的资料")
                    
                    # 提供下载
                    st.subheader("📥 下载生成的资料")
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label=f"下载 {zip_filename}",
                            data=f,
                            file_name=zip_filename,
                            mime="application/zip",
                            key="download_manual"
                        )
                    
                except Exception as e:
                    st.error(f"❌ 生成失败: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
                    
    except json.JSONDecodeError:
        st.error("❌ JSON文件格式错误，请检查文件内容")
    except Exception as e:
        st.error(f"❌ 处理文件时出错: {str(e)}")
else:
    st.info("请上传包含课程信息的JSON文件")

# 底部信息
st.markdown("---")
st.caption("💡 提示：请确保JSON文件包含完整的课程信息，包括课程名称、代码、教师姓名等字段")


# Flask API端点
@flask_app.route('/api/generate', methods=['POST'])
def api_generate():
    """API端点：生成课程资料"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'success': False, 'message': '请求数据为空'}), 400
        
        courses = data.get('courses', [])
        
        if not courses:
            return jsonify({'success': False, 'message': '课程数据为空'}), 400
        
        # 创建生成器实例
        generator = RenYuMaterialsGenerator()
        generator.template_dir = "期初资料1"
        generator.output_base_dir = "output"
        generator.courses = courses
        generator.metadata = {
            'date': datetime.now().strftime('%Y年%m月%d日'),
            'semester': '2026年春季学期'
        }
        
        # 生成资料
        generator.generate_all_materials()
        
        # 获取生成的文件
        teacher_name = courses[0].get('teacherName', '未知教师')
        output_folder = os.path.join(generator.output_base_dir, f"{teacher_name}_天府学院期初资料")
        
        # 创建ZIP文件
        zip_filename = f"{teacher_name}_课程资料.zip"
        zip_path = os.path.join("output", zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, "output")
                    zipf.write(file_path, arcname)
        
        # 返回下载链接
        download_url = f"/api/download/{zip_filename}"
        
        return jsonify({
            'success': True,
            'zipUrl': download_url,
            'filename': zip_filename,
            'message': f'成功生成 {len(courses)} 门课程的资料'
        })
    
    except Exception as e:
        print(f"API生成失败: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'success': False, 'message': f'生成失败: {str(e)}'}), 500


@flask_app.route('/api/download/<filename>', methods=['GET'])
def api_download(filename):
    """API端点：下载文件"""
    try:
        file_path = os.path.join("output", filename)
        
        if not os.path.exists(file_path):
            return jsonify({'success': False, 'message': '文件不存在'}), 404
        
        return flask_app.send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/zip'
        )
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'下载失败: {str(e)}'}), 500


# 在单独的线程中启动Flask服务器
def start_flask_server():
    """在单独的线程中启动Flask服务器"""
    try:
        flask_app.run(host='0.0.0.0', port=5001, debug=False)
    except Exception as e:
        print(f"Flask服务器启动失败: {str(e)}")
        print(traceback.format_exc())


# 启动Flask服务器
flask_thread = threading.Thread(target=start_flask_server, daemon=True)
flask_thread.start()
