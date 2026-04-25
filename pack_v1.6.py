#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
售后登记表软件打包脚本 - 版本 1.6
打包成独立的exe文件，不显示控制台窗口
包含所有必需文件：UI文件、样式表、图标、数据库
"""

import os
import sys
import subprocess
import shutil
from datetime import datetime

def main():
    print("=== 售后登记表软件打包工具 v1.6 ===")
    print("开始打包...")

    project_name = "售后登记表_v1.6"
    dist_path = "dist"
    build_path = "build"

    if not os.path.exists(dist_path):
        os.makedirs(dist_path)

    if os.path.exists(build_path):
        shutil.rmtree(build_path)
        print("[OK] 清理build目录")

    existing_exe = os.path.join(dist_path, f"{project_name}.exe")
    if os.path.exists(existing_exe):
        os.remove(existing_exe)
        print(f"[OK] 清理旧的exe文件")

    cmd = [
        "pyinstaller",
        f"--name={project_name}",
        "--onefile",
        "--windowed",
        "--icon=NONE",
        "--add-data=input_panel.ui;.",
        "--add-data=search_panel.ui;.",
        "--add-data=quick_date_panel.ui;.",
        "--add-data=table_panel.ui;.",
        "--add-data=dialog_add_store.ui;.",
        "--add-data=dialog_store_settings.ui;.",
        "--add-data=main_window.ui;.",
        "--add-data=dopamine_styles.qss;.",
        "--add-data=icons;icons",
        "--add-data=help_dialog.py;.",
        "--hidden-import=PyQt5",
        "--hidden-import=PyQt5.QtWidgets",
        "--hidden-import=PyQt5.QtCore",
        "--hidden-import=PyQt5.QtGui",
        "--hidden-import=PyQt5.uic",
        "--hidden-import=matplotlib.backends.backend_qt5agg",
        "--hidden-import=matplotlib.figure",
        "--hidden-import=matplotlib",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=numpy",
        "--hidden-import=requests",
        "--hidden-import=markdown",
        "--hidden-import=sqlite3",
        "--hidden-import=datetime",
        "--hidden-import=re",
        "--hidden-import=json",
        "--hidden-import=os",
        "--hidden-import=sys",
        "--hidden-import=threading",
        "--hidden-import=tempfile",
        "--hidden-import=shutil",
        "--hidden-import=subprocess",
        "--hidden-import=socket",
        "--hidden-import=urllib.parse",
        "--clean",
        "dj.py"
    ]

    print("执行打包命令:")
    print(" ".join(cmd))

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=os.getcwd())

        if result.returncode == 0:
            print("[OK] 打包成功!")

            exe_path = os.path.join(dist_path, f"{project_name}.exe")
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"[OK] 生成的exe文件: {exe_path}")
                print(f"[OK] 文件大小: {file_size:.2f} MB")
                print(f"[OK] 打包时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"[OK] 版本标识: v1.6")

                print("\n*** 打包完成! ***")
                print("生成的文件在 'dist' 目录中:")
                print(f"   - {project_name}.exe (单个独立程序)")
                print("\n使用说明:")
                print(f"   1. 双击 '{project_name}.exe' 运行程序")
                print("   2. 所有必需文件已内嵌在exe中，无需额外文件")
                print("   3. 程序启动后不会显示黑色命令提示符")
                print("   4. 这是一个真正的独立exe文件，可以单独分发")
                print("   5. 数据文件(refund_data.db)使用用户本地文件，不打包")
                print("   6. 数据库文件使用.db后缀格式")

            else:
                print("[ERROR] 打包失败: 未找到生成的exe文件")

        else:
            print("[ERROR] 打包失败!")
            print("错误信息:")
            print(result.stdout)
            print(result.stderr)

    except Exception as e:
        print(f"[ERROR] 打包过程中出现错误: {e}")
        print("请确保已安装pyinstaller: pip install pyinstaller")

if __name__ == "__main__":
    main()