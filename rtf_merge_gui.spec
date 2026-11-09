# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置 —— 图形界面版（onedir + windowed）
# 用法：pyinstaller rtf_merge_gui.spec
# 说明：rtfmerge/ps/*.ps1 必须作为数据文件随包分发（后处理脚本外置于包内）；
#       图标同时以 --add-data 带一份，供运行时窗口图标使用。

from pathlib import Path

ROOT = Path(r"F:\rtf_merge")

a = Analysis(
    [str(ROOT / "rtf_merge_gui.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        (str(ROOT / "rtfmerge" / "ps"), "rtfmerge/ps"),
        (str(ROOT / "rtf_merge.ico"), "."),
    ],
    hiddenimports=["rtfmerge", "rtfmerge.merge_core", "rtfmerge.word_post",
                   "rtfmerge.docx_xml", "rtfmerge.rtf_lint", "rtfmerge.gui"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="rtf_merge_gui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=str(ROOT / "rtf_merge.ico"),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="rtf_merge_gui",
)
