import requests
import os
import json
import base64

def upload_file_via_mcp(file_path, mcp_endpoint, api_key, agent_id):
    """
    通过MCP API接口上传文件到智能体平台
    
    Args:
        file_path: 文件路径
        mcp_endpoint: MCP API端点
        api_key: API密钥
        agent_id: 智能体ID
    
    Returns:
        dict: 上传结果
    """
    try:
        print(f"正在通过MCP上传文件: {file_path}")
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件不存在: {file_path}")
            return {"code": "error", "message": "文件不存在"}
        
        # 准备请求头
        headers = {
            "Authorization": f"Bearer {api_key}"
        }
        
        # 准备请求数据 (form-data格式)
        with open(file_path, 'rb') as f:
            files = {
                "file": (os.path.basename(file_path), f)
            }
            
            data = {
                "type": "store"
            }
        
        # 发送请求
        print(f"上传到: {mcp_endpoint}")
        response = requests.post(mcp_endpoint, headers=headers, files=files, data=data)
        
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

def batch_upload_files(base_dir, mcp_endpoint, api_key, agent_id):
    """
    批量上传生成的教学资料到智能体平台
    
    Args:
        base_dir: 基础目录
        mcp_endpoint: MCP API端点
        api_key: API密钥
        agent_id: 智能体ID
    """
    print("=" * 80)
    print("🚀 开始通过MCP上传教学资料到智能体平台")
    print("=" * 80)
    
    # 遍历所有zip文件
    uploaded_files = []
    failed_files = []
    
    for file in os.listdir(base_dir):
        if file.endswith('.zip') and '期初教学资料' in file:
            file_path = os.path.join(base_dir, file)
            
            # 上传文件
            result = upload_file_via_mcp(file_path, mcp_endpoint, api_key, agent_id)
            
            if result.get('code') == '0000' or result.get('code') == 0 or result.get('success'):
                uploaded_files.append({
                    "file": file,
                    "path": file_path,
                    "url": result.get('data', {}).get('url', '') or result.get('url', '')
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
    
    # MCP API配置
    mcp_endpoint = "http://localhost:8000/api/mcp/file/upload"  # 替换为实际的MCP API端点
    api_key = "ak-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"  # 替换为实际的API密钥
    agent_id = "your_agent_id"  # 替换为实际的智能体ID
    
    # 批量上传文件
    result = batch_upload_files(base_dir, mcp_endpoint, api_key, agent_id)
    
    # 保存上传结果
    with open(os.path.join(base_dir, "mcp_upload_result.json"), 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"上传结果已保存到: mcp_upload_result.json")

if __name__ == "__main__":
    main()