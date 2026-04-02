#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
售后登记表软件打包脚本 - 版本 1.2
打包成独立的exe文件，不显示控制台窗口
"""

import os
import sys
import subprocess
import shutil
from datetime import datetime

def main():
    print("=== 售后登记表软件打包工具 v1.2 ===")
    print("开始打包...")
    
    # 创建dist目录（如果不存在）
    if not os.path.exists("dist"):
        os.makedirs("dist")
    
    # 打包命令
    cmd = [
        "pyinstaller",
        "--name=售后登记表_v1.2",
        "--onefile",  # 打包成单个exe文件
        "--windowed",  # 不显示控制台窗口
        "--icon=NONE",  # 不使用图标
        "--add-data=input_panel.ui;.",  # 添加UI文件
        "--add-data=search_panel.ui;.",  # 添加UI文件
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
            exe_path = os.path.join("dist", "售后登记表_v1.2.exe")
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
                print(f"✅ 生成的exe文件: {exe_path}")
                print(f"✅ 文件大小: {file_size:.2f} MB")
                print(f"✅ 打包时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                # 复制必要的文件到dist目录
                files_to_copy = ["input_panel.ui", "search_panel.ui"]
                for file in files_to_copy:
                    if os.path.exists(file):
                        shutil.copy2(file, "dist")
                        print(f"✅ 已复制: {file}")
                
                print("\n🎉 打包完成!")
                print("📁 生成的文件在 'dist' 目录中:")
                print("   - 售后登记表_v1.2.exe (主程序)")
                print("   - input_panel.ui (UI文件)")
                print("   - search_panel.ui (UI文件)")
                print("\n💡 使用说明:")
                print("   1. 双击 '售后登记表_v1.2.exe' 运行程序")
                print("   2. 确保所有UI文件与exe在同一目录")
                print("   3. 程序启动后不会显示黑色命令提示符")
                
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