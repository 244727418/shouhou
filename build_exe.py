#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
售后登记表软件打包脚本 - 版本 1.3
打包成独立的exe文件，不显示控制台窗口
包含所有必需文件：UI文件、样式表、图标、数据库
"""

import os
import sys
import subprocess
import shutil
from datetime import datetime

def main():
    print("=== 售后登记表软件打包工具 v1.3 ===")
    print("开始打包...")
    
    # 创建dist目录（如果不存在）
    if not os.path.exists("dist"):
        os.makedirs("dist")
    
    # 打包命令
    cmd = [
        "pyinstaller",
        "--name=售后登记表_v1.3",
        "--onefile",  # 打包成单个exe文件
        "--windowed",  # 不显示控制台窗口
        "--icon=NONE",  # 不使用图标
        # 添加所有UI文件
        "--add-data=input_panel.ui;.",
        "--add-data=search_panel.ui;.",
        "--add-data=quick_date_panel.ui;.",
        "--add-data=table_panel.ui;.",
        "--add-data=dialog_add_store.ui;.",
        "--add-data=dialog_store_settings.ui;.",
        "--add-data=main_window.ui;.",
        # 添加样式表文件
        "--add-data=dopamine_styles.qss;.",
        # 添加图标文件
        "--add-data=icons/check_green.svg;icons",
        "--add-data=icons/cross_red.svg;icons",
        # 注意：数据库文件不打包，使用用户本地的refund_data.db文件
        # 隐藏导入
        "--hidden-import=PyQt5.QtWidgets",
        "--hidden-import=PyQt5.QtCore", 
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=matplotlib.backends.backend_qt5agg",
        "--hidden-import=matplotlib.figure",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--clean",  # 清理临时文件
        "dj.py"
    ]
    
    print("执行打包命令:")
    print(" ".join(cmd))
    
    try:
        # 执行打包命令
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())
        
        if result.returncode == 0:
            print("✅ 打包成功!")
            
            # 检查生成的exe文件
            exe_path = os.path.join("dist", "售后登记表_v1.3.exe")
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
                print(f"✅ 生成的exe文件: {exe_path}")
                print(f"✅ 文件大小: {file_size:.2f} MB")
                print(f"✅ 打包时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                print("\n🎉 打包完成!")
                print("📁 生成的文件在 'dist' 目录中:")
                print("   - 售后登记表_v1.3.exe (单个独立程序)")
                print("\n💡 使用说明:")
                print("   1. 双击 '售后登记表_v1.3.exe' 运行程序")
                print("   2. 所有必需文件已内嵌在exe中，无需额外文件")
                print("   3. 程序启动后不会显示黑色命令提示符")
                print("   4. 这是一个真正的独立exe文件，可以单独分发")
                
            else:
                print("❌ 打包失败: 未找到生成的exe文件")
                
        else:
            print("❌ 打包失败!")
            print("错误信息:")
            print(result.stdout)
            print(result.stderr)
            
    except Exception as e:
        print(f"❌ 打包过程中出现错误: {e}")
        print("请确保已安装pyinstaller: pip install pyinstaller")

if __name__ == "__main__":
    main()