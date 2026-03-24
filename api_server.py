from flask import Flask, request, jsonify, send_file
import os
import json
import tempfile
import zipfile
from flask_cors import CORS
from process_all_json import process_all_json_files
from mcp_uploader import batch_upload_files

app = Flask(__name__)
CORS(app)  # 启用跨域支持

@app.route('/api/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({"status": "ok", "message": "API服务正常运行"})

@app.route('/api/generate', methods=['POST'])
def generate_materials():
    """
    生成期初教学资料API
    
    请求体格式:
    {
        "courses": [...],
        "metadata": {...}
    }
    
    返回:
    {
        "success": true/false,
        "message": "...",
        "download_url": "..."  # 生成的zip文件下载链接
    }
    """
    try:
        # 获取请求数据
        data = request.get_json()
        
        if not data or 'courses' not in data:
            return jsonify({
                "success": False,
                "message": "请求数据格式错误，缺少courses字段"
            }), 400
        
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        
        # 保存JSON数据到临时文件
        json_file_path = os.path.join(temp_dir, 'course_data.json')
        with open(json_file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # 生成课程资料
        process_all_json_files(temp_dir)
        
        # 创建zip文件
        teacher_name = data.get('metadata', {}).get('teacherName', '未知教师')
        zip_filename = f"{teacher_name}_期初教学资料.zip"
        zip_path = os.path.join(temp_dir, zip_filename)
        
        # 查找生成的教师文件夹
        teacher_folders = [d for d in os.listdir(temp_dir) 
                          if os.path.isdir(os.path.join(temp_dir, d)) 
                          and not d.startswith('.')]
        
        if teacher_folders:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for folder in teacher_folders:
                    folder_path = os.path.join(temp_dir, folder)
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            zipf.write(file_path, arcname)
            
            return jsonify({
                "success": True,
                "message": "课程资料生成成功",
                "teacher_name": teacher_name,
                "course_count": len(data.get('courses', [])),
                "download_url": f"/api/download/{os.path.basename(temp_dir)}/{zip_filename}"
            })
        else:
            return jsonify({
                "success": False,
                "message": "未找到生成的课程资料"
            }), 500
            
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"生成失败: {str(e)}"
        }), 500

@app.route('/api/download/<temp_id>/<filename>', methods=['GET'])
def download_file(temp_id, filename):
    """下载生成的文件"""
    try:
        # 这里需要实现安全的文件下载逻辑
        # 实际部署时应该使用云存储或临时文件服务
        return jsonify({
            "success": True,
            "message": "文件下载功能需要配置云存储服务"
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"下载失败: {str(e)}"
        }), 500

@app.route('/api/upload-to-mcp', methods=['POST'])
def upload_to_mcp():
    """
    上传文件到MCP平台
    
    请求体格式:
    {
        "file_path": "...",
        "mcp_endpoint": "...",
        "api_key": "...",
        "agent_id": "..."
    }
    """
    try:
        data = request.get_json()
        
        file_path = data.get('file_path')
        mcp_endpoint = data.get('mcp_endpoint')
        api_key = data.get('api_key')
        agent_id = data.get('agent_id', '1806')
        
        if not all([file_path, mcp_endpoint, api_key]):
            return jsonify({
                "success": False,
                "message": "缺少必要参数"
            }), 400
        
        # 调用MCP上传功能
        result = batch_upload_files(
            os.path.dirname(file_path),
            mcp_endpoint,
            api_key,
            agent_id
        )
        
        return jsonify({
            "success": result['success'] > 0,
            "message": f"上传完成，成功: {result['success']}, 失败: {result['failed']}",
            "details": result
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"上传失败: {str(e)}"
        }), 500

if __name__ == '__main__':
    # 生产环境应该使用gunicorn等WSGI服务器
    # 开发环境使用Flask内置服务器
    app.run(host='0.0.0.0', port=5000, debug=True)