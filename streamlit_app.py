import streamlit as st
import os
import json
import tempfile
import zipfile
from process_all_json import process_all_json_files
from mcp_uploader import batch_upload_files

# 设置页面配置
st.set_page_config(
    page_title="期初教学资料填制助手",
    page_icon="📚",
    layout="wide"
)

# 页面标题
st.title("📚 期初教学资料填制助手")

# 侧边栏配置
st.sidebar.title("配置")

# MCP API配置
mcp_endpoint = st.sidebar.text_input("MCP API端点", "https://znt.tfswufe.edu.cn/api/file/upload")
api_key = st.sidebar.text_input("API密钥", "ak-e7cc7011dce24aa88ed0ed2bba49c90f", type="password")
# 智能体ID写死为1806
agent_id = "1806"
st.sidebar.text_input("智能体ID", agent_id, disabled=True)

# 上传JSON文件
st.header("1. 上传课程信息JSON文件")
uploaded_files = st.file_uploader("选择JSON文件", accept_multiple_files=True, type="json")

# 临时目录
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

# 保存上传的文件
if uploaded_files:
    st.success(f"成功上传 {len(uploaded_files)} 个文件")
    for file in uploaded_files:
        file_path = os.path.join(st.session_state.temp_dir, file.name)
        with open(file_path, 'wb') as f:
            f.write(file.getbuffer())
        st.write(f"- {file.name}")

# 生成期初教学填报资料
st.header("2. 生成期初教学填报资料")
if st.button("生成期初教学填报资料"):
    if uploaded_files:
        with st.spinner("正在生成期初教学填报资料..."):
            process_all_json_files(st.session_state.temp_dir)
        st.success("期初教学填报资料生成完成！")
    else:
        st.error("请先上传JSON文件")

# 推送文件到智能体平台
st.header("3. 推送文件到智能体平台")
if st.button("推送文件"):
    if uploaded_files and api_key != "ak-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx":
        with st.spinner("正在准备文件..."):
            # 确保zip文件已创建
            create_zip_files(st.session_state.temp_dir)
        with st.spinner("正在推送文件..."):
            result = batch_upload_files(
                st.session_state.temp_dir,
                mcp_endpoint,
                api_key,
                agent_id
            )
        
        st.success(f"推送完成！成功: {result['success']}, 失败: {result['failed']}")
        
        if result['uploaded_files']:
            st.subheader("成功推送的文件:")
            for item in result['uploaded_files']:
                st.write(f"- {item['file']}")
                if item.get('url'):
                    st.write(f"  下载链接: {item['url']}")
        
        if result['failed_files']:
            st.subheader("推送失败的文件:")
            for item in result['failed_files']:
                st.write(f"- {item['file']}: {item['error']}")
    else:
        if not uploaded_files:
            st.error("请先上传JSON文件")
        else:
            st.error("请设置正确的API密钥")

def create_zip_files(base_dir):
    """为每个教师创建zip文件，包含所有生成的教学资料"""
    zip_files = []
    
    # 遍历所有教师文件夹
    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)
        if os.path.isdir(item_path):
            # 提取教师姓名
            teacher_name = item.split('_')[1] if len(item.split('_')) > 1 else item
            
            # 创建zip文件
            zip_filename = f"{teacher_name}_期初教学资料.zip"
            zip_path = os.path.join(base_dir, zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 遍历教师文件夹中的所有文件
                for root, dirs, files in os.walk(item_path):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, base_dir)
                        zipf.write(file_path, arcname)
            
            zip_files.append(zip_path)
    
    return zip_files

# 显示生成的文件
st.header("4. 生成的教学资料")
if os.path.exists(st.session_state.temp_dir):
    # 创建zip文件
    zip_files = create_zip_files(st.session_state.temp_dir)
    
    # 显示zip文件
    for zip_path in zip_files:
        zip_filename = os.path.basename(zip_path)
        st.write(f"- {zip_filename}")
        with open(zip_path, 'rb') as f:
            st.download_button(
                label=f"下载 {zip_filename}",
                data=f,
                file_name=zip_filename,
                mime="application/zip"
            )

# 清理临时文件
if st.button("清理临时文件"):
    import shutil
    shutil.rmtree(st.session_state.temp_dir)
    st.session_state.temp_dir = tempfile.mkdtemp()
    st.success("临时文件已清理")

# 说明
st.sidebar.header("使用说明")
st.sidebar.markdown("""
1. 上传包含课程信息的JSON文件
2. 点击"生成期初教学填报资料"按钮生成相关资料
3. 配置MCP API参数
4. 点击"推送文件"按钮将资料推送到智能体平台
5. 可以直接下载生成的压缩文件
""")