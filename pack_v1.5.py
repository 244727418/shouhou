#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v1.5版本打包脚本
确保所有UI文件和资源文件正确打包
"""

import os
import sys
import subprocess
import shutil

def main():
    print("🚀 开始打包售后管理工具 v1.5...")
    
    # 清理之前的打包文件
    if os.path.exists('build'):
        shutil.rmtree('build')
        print("✅ 清理build目录")
    
    if os.path.exists('dist'):
        # 只删除旧的exe文件，保留数据库文件
        for file in os.listdir('dist'):
            if file.endswith('.exe'):
                os.remove(os.path.join('dist', file))
        print("✅ 清理dist目录中的exe文件")
    
    # 创建PyInstaller命令
    cmd = [
        'pyinstaller',
        '--onefile',           # 打包为单个exe文件
        '--noconsole',         # 不显示控制台窗口
        '--name=售后管理工具_v1.5',  # 输出文件名
        '--add-data=dialog_add_store.ui;.',      # 添加UI文件
        '--add-data=dialog_store_settings.ui;.',  # 添加UI文件
        '--add-data=input_panel.ui;.',           # 添加UI文件
        '--add-data=main_window.ui;.',           # 添加UI文件
        '--add-data=quick_date_panel.ui;.',      # 添加UI文件
        '--add-data=search_panel.ui;.',          # 添加UI文件
        '--add-data=table_panel.ui;.',           # 添加UI文件
        '--add-data=dopamine_styles.qss;.',      # 添加样式文件
        '--add-data=help_dialog.py;.',           # 添加帮助对话框模块
        '--add-data=icons;icons',                # 添加图标文件夹
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=sqlite3',
        '--hidden-import=requests',
        '--hidden-import=openpyxl',
        '--hidden-import=xlrd',
        '--hidden-import=matplotlib',
        '--hidden-import=matplotlib.backends.backend_qt5agg',
        '--hidden-import=numpy',
        '--hidden-import=markdown',
        '--hidden-import=datetime',
        '--hidden-import=re',
        '--hidden-import=json',
        '--hidden-import=os',
        '--hidden-import=sys',
        '--hidden-import=threading',
        '--hidden-import=tempfile',
        '--hidden-import=shutil',
        '--hidden-import=subprocess',
        '--hidden-import=socket',
        '--hidden-import=urllib.parse',
        'dj.py'
    ]
    
    print("📦 开始打包过程...")
    print(f"执行命令: {' '.join(cmd)}")
    
    # 执行打包命令
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        
        if result.returncode == 0:
            print("✅ 打包成功！")
            
            # 检查生成的文件
            exe_path = os.path.join('dist', '售后管理工具_v1.5.exe')
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
                print(f"📁 生成文件: {exe_path}")
                print(f"📊 文件大小: {file_size:.2f} MB")
                
                if file_size > 100:
                    print("✅ 文件大小正常（>100MB）")
                else:
                    print("⚠️ 文件大小可能偏小，可能缺少某些依赖")
                    
                # 列出dist目录内容
                print("\n📂 dist目录内容:")
                for file in os.listdir('dist'):
                    file_path = os.path.join('dist', file)
                    if os.path.isfile(file_path):
                        size = os.path.getsize(file_path) / (1024 * 1024)
                        print(f"   {file} ({size:.2f} MB)")
                    else:
                        print(f"   {file}/ (目录)")
                        
            else:
                print("❌ 打包失败：exe文件未生成")
                
        else:
            print("❌ 打包失败！")
            print("错误输出:")
            print(result.stderr)
            print("标准输出:")
            print(result.stdout)
            
    except Exception as e:
        print(f"❌ 打包过程出错: {e}")
        
    print("\n🎯 打包完成！")

if __name__ == "__main__":
    main()