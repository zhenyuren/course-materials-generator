import json
import os
from datetime import datetime

def main(args: dict) -> dict:
    params = args.get("params")
    
    # 从params中获取课程信息占位变量
    # 直接使用params作为课程信息，因为代码节点的输出是直接的课程信息对象
    course_info = params
    
    # 添加生成时间戳
    course_info["generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 定义保存路径
    save_dir = r"D:\智能金融学院\智能体设计"
    
    # 确保目录存在
    os.makedirs(save_dir, exist_ok=True)
    
    # 生成文件名（使用时间戳避免重复）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"course_info_{timestamp}.json"
    filepath = os.path.join(save_dir, filename)
    
    # 保存JSON文件
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(course_info, f, ensure_ascii=False, indent=2)
        
        ret = {
            "status": "success",
            "message": "JSON文件生成成功",
            "file_path": filepath,
            "filename": filename,
            "course_info": course_info
        }
    except Exception as e:
        ret = {
            "status": "error",
            "message": f"JSON文件生成失败: {str(e)}",
            "file_path": None,
            "filename": None,
            "course_info": course_info
        }
    
    return ret