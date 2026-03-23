#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
课程期初资料生成工具 - 打包脚本
使用PyInstaller将应用打包成独立可执行文件
"""

import os
import shutil
import subprocess
import zipfile
import platform

def build_app():
    """使用PyInstaller打包应用"""
    
    print("开始打包课程期初资料生成工具...")
    
    # 清理旧的构建文件
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # 检查PyInstaller是否安装
    try:
        subprocess.run(['pyinstaller', '--version'], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("安装PyInstaller...")
        subprocess.run(['pip', 'install', 'pyinstaller'], check=True)
    
    # 根据操作系统确定分隔符
    separator = ':' if platform.system() == 'Darwin' else ';'
    
    # 构建应用
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name=课程期初资料生成工具',
        f'--add-data=期初资料1{separator}期初资料1',
        f'--add-data=课程信息_2026年03月17日.json{separator}.',
        'course_materials_app.py']
    
    try:
        subprocess.run(cmd, check=True)
        print("✅ 应用打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失败: {e}")
        return False

def create_distribution():
    """创建分发包"""
    
    print("创建分发包...")
    
    # 创建分发目录
    dist_dir = '课程期初资料生成工具_分发版'
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    os.makedirs(dist_dir)
    
    # 复制可执行文件
    if platform.system() == 'Darwin':  # macOS
        app_path = os.path.join('dist', '课程期初资料生成工具.app')
        if os.path.exists(app_path):
            shutil.copytree(app_path, os.path.join(dist_dir, '课程期初资料生成工具.app'))
    elif platform.system() == 'Windows':  # Windows
        exe_path = os.path.join('dist', '课程期初资料生成工具.exe')
        if os.path.exists(exe_path):
            shutil.copy2(exe_path, dist_dir)
    
    # 复制必要文件
    files_to_copy = [
        '使用说明.md',
        'requirements.txt',
        '期初资料1',
        '课程信息_2026年03月17日.json'
    ]
    
    for file_path in files_to_copy:
        if os.path.exists(file_path):
            if os.path.isdir(file_path):
                shutil.copytree(file_path, os.path.join(dist_dir, os.path.basename(file_path)))
            else:
                shutil.copy2(file_path, dist_dir)
    
    # 创建ZIP包
    zip_file = '课程期初资料生成工具_分发版.zip'
    with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(dist_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, '.')
                zipf.write(file_path, arcname)
    
    print(f"✅ 分发包创建成功: {zip_file}")
    return zip_file

def main():
    """主函数"""
    print("=" * 60)
    print("课程期初资料生成工具 - 打包与分发")
    print("=" * 60)
    
    # 打包应用
    if build_app():
        # 创建分发包
        zip_file = create_distribution()
        print(f"\n🎉 打包完成！")
        print(f"分发包已保存为: {zip_file}")
        print("\n使用说明：")
        print("1. 将压缩包发送给其他老师")
        print("2. 解压后运行'课程期初资料生成工具'")
        print("3. 按照'使用说明.md'操作即可")

if __name__ == "__main__":
    main()
