# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置 —— 命令行版（onefile）
# 用法：pyinstaller rtf_merge_cli.spec
# 说明：rtfmerge/ps/*.ps1 必须作为数据文件随包分发（后处理脚本外置于包内）。

from pathlib import Path

ROOT = Path(r"F:\rtf_merge")

a = Analysis(
    [str(ROOT / "rtf_merge.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        (str(ROOT / "rtfmerge" / "ps"), "rtfmerge/ps"),
        (str(ROOT / "rtf_merge.ico"), "."),
    ],
    hiddenimports=["rtfmerge", "rtfmerge.merge_core", "rtfmerge.word_post",
                   "rtfmerge.docx_xml", "rtfmerge.rtf_lint", "rtfmerge.cli"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="rtf_merge",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    icon=str(ROOT / "rtf_merge.ico"),
)
