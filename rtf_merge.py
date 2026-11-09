#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""rtf_merge —— 兼容入口（v1.1.0 起拆分为 rtfmerge 包）。

此文件仅保留旧调用路径，实际实现位于 rtfmerge/：
    rtfmerge/merge_core.py  收集 / 排序 / 合并 / 原子写盘 / 合并清单
    rtfmerge/word_post.py   Word 后处理编排（ps/*.ps1 外置脚本）
    rtfmerge/docx_xml.py    字体 / 引号 / 加粗 / 验收
    rtfmerge/cli.py         命令行入口
    rtfmerge/gui.py         图形界面

兼容用法（与旧版一致）：
    from rtf_merge import merge_rtf, fix_fonts_via_word
    python rtf_merge.py --inpath D:\\xx\\output_rtf --outfile D:\\xx\\all.rtf
"""
from __future__ import annotations

import sys
from pathlib import Path

# 源码目录运行时，确保同目录的 rtfmerge 包可被导入
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from rtfmerge import __version__, fix_fonts_via_word, merge_rtf   # noqa: E402,F401
from rtfmerge.cli import main                                     # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
