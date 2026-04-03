#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v1.5版本打包脚本 - 修复版
确保所有UI文件和资源文件正确打包
"""

import os
import sys
import subprocess
import shutil

def create_spec_file():
    """创建PyInstaller spec文件，确保正确包含所有资源"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 添加当前目录到路径
import sys
sys.path.insert(0, '.')

a = Analysis(
    ['dj.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # 包含所有UI文件
        ('dialog_add_store.ui', '.'),
        ('dialog_store_settings.ui', '.'),
        ('input_panel.ui', '.'),
        ('main_window.ui', '.'),
        ('quick_date_panel.ui', '.'),
        ('search_panel.ui', '.'),
        ('table_panel.ui', '.'),
        # 包含样式文件
        ('dopamine_styles.qss', '.'),
        # 包含帮助对话框模块
        ('help_dialog.py', '.'),
        # 包含图标文件夹
        ('icons/', 'icons/'),
    ],
    hiddenimports=[
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore', 
        'PyQt5.QtGui',
        'sqlite3',
        'requests',
        'openpyxl',
        'xlrd',
        'matplotlib',
        'matplotlib.backends.backend_qt5agg',
        'numpy',
        'markdown',
        'datetime',
        're',
        'json',
        'os',
        'sys',
        'threading',
        'tempfile',
        'shutil',
        'subprocess',
        'socket',
        'urllib.parse',
        'PyQt5.uic',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 添加数据文件（排除数据文件）
# 注意：我们不包含数据库文件和其他数据文件，只包含程序文件

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='售后管理工具_v1.5',  # 生成的exe文件名
    debug=False,  # 不包含调试信息
    bootloader_ignore_signals=False,
    strip=False,  # 不剥离符号
    upx=True,  # 使用UPX压缩
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 重要：设置为False去掉黑框！
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以设置图标文件路径
)
'''
    
    with open('售后管理工具_v1.5.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ 创建spec文件成功")

def main():
    print("🚀 开始打包售后管理工具 v1.5（修复版）...")
    
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
    
    # 创建spec文件
    create_spec_file()
    
    # 使用spec文件打包
    cmd = ['pyinstaller', '售后管理工具_v1.5.spec']
    
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
    
    # 清理临时文件
    if os.path.exists('售后管理工具_v1.5.spec'):
        os.remove('售后管理工具_v1.5.spec')
        print("✅ 清理临时spec文件")
        
    print("\n🎯 打包完成！")

if __name__ == "__main__":
    main()