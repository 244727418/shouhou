# -*- mode: python ; coding: utf-8 -*-

from scripts.pyinstaller_spec_common import (
    APP_NAME,
    DATA_FILES,
    ENTRY_SCRIPT,
    EXCLUDES,
    HIDDEN_IMPORTS,
    ICON_PATH,
    prune_analysis_datas,
)


a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[],
    binaries=[],
    datas=DATA_FILES,
    hiddenimports=HIDDEN_IMPORTS,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

a.datas = prune_analysis_datas(a.datas)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=ICON_PATH,
)
