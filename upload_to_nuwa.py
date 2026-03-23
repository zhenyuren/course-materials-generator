import requests
import os
import json

def upload_file_to_nuwa(file_path, api_key, file_type='store'):
    """
    上传文件到女娲智能体平台
    
    Args:
        file_path: 文件路径
        api_key: 女娲平台的API Key
        file_type: 存储类型，'tmp'临时文件，'store'永久存储
    
    Returns:
        dict: 上传结果
    """
    try:
        print(f"正在上传文件: {file_path}")
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件不存在: {file_path}")
            return {"code": "error", "message": "文件不存在"}
        
        # 检查文件大小
        file_size = os.path.getsize(file_path)
        print(f"文件大小: {file_size} 字节")
        
        # 准备请求头
        headers = {
            "Authorization": f"Bearer {api_key}"
        }
        
        # 准备请求数据
        files = {
            "file": open(file_path, 'rb')
        }
        
        data = {
            "type": file_type
        }
        
        # 发送请求
        url = "http://localhost:8000/api/file/upload"  # 假设本地部署的女娲平台
        # 如果是云部署，需要修改为实际的API地址
        # url = "https://your-nuwa-platform/api/file/upload"
        
        print(f"上传到: {url}")
        response = requests.post(url, headers=headers, files=files, data=data)
        
        # 关闭文件
        files["file"].close()
        
        # 处理响应
        if response.status_code == 200:
            result = response.json()
            print(f"上传成功! 响应: {result}")
            return result
        else:
            print(f"上传失败! 状态码: {response.status_code}")
            print(f"响应内容: {response.text}")
            return {"code": "error", "message": f"上传失败，状态码: {response.status_code}"}
            
    except Exception as e:
        print(f"上传时出错: {str(e)}")
        return {"code": "error", "message": str(e)}

def upload_all_outlines(base_dir, api_key):
    """
    上传所有生成的大纲文件到女娲平台
    
    Args:
        base_dir: 基础目录
        api_key: 女娲平台的API Key
    """
    print("=" * 80)
    print("🚀 开始上传大纲文件到女娲平台")
    print("=" * 80)
    
    # 遍历所有教师文件夹
    uploaded_files = []
    failed_files = []
    
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.docx') and '大纲' in file:
                file_path = os.path.join(root, file)
                
                # 上传文件
                result = upload_file_to_nuwa(file_path, api_key)
                
                if result.get('code') == '0000' or result.get('code') == 0:
                    uploaded_files.append({
                        "file": file,
                        "path": file_path,
                        "url": result.get('data', {}).get('url', '')
                    })
                else:
                    failed_files.append({
                        "file": file,
                        "path": file_path,
                        "error": result.get('message', '未知错误')
                    })
    
    print("=" * 80)
    print(f"上传完成！")
    print(f"成功上传: {len(uploaded_files)} 个文件")
    print(f"上传失败: {len(failed_files)} 个文件")
    
    if uploaded_files:
        print("\n成功上传的文件:")
        for item in uploaded_files:
            print(f"  - {item['file']}")
            if item['url']:
                print(f"    URL: {item['url']}")
    
    if failed_files:
        print("\n上传失败的文件:")
        for item in failed_files:
            print(f"  - {item['file']}: {item['error']}")
    
    print("=" * 80)
    
    return {
        "success": len(uploaded_files),
        "failed": len(failed_files),
        "uploaded_files": uploaded_files,
        "failed_files": failed_files
    }

def main():
    # 定义路径
    base_dir = r"D:\智能金融学院\智能体设计"
    
    # 请替换为实际的API Key
    api_key = "ak-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    
    # 上传所有大纲文件
    result = upload_all_outlines(base_dir, api_key)
    
    # 保存上传结果
    with open(os.path.join(base_dir, "upload_result.json"), 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"上传结果已保存到: upload_result.json")

if __name__ == "__main__":
    main()