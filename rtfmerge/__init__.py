# -*- coding: utf-8 -*-
"""rtfmerge —— RTF 合并工具（Python 版 reimplementation of SAS 宏
%rtf_mergeV4_1）。

模块划分：
  merge_core —— 收集 / 排序 / 逐行合并 / 原子写盘 / 合并清单（纯文本，可单测）
  word_post  —— PowerShell/Word COM 后处理编排（脚本外置在 ps/ 下）
  docx_xml   —— python-docx 层的字体 / 引号 / 加粗 / 验收（唯一最终写盘者）
  cli        —— 命令行入口
  gui        —— tkinter 图形界面

对外兼容接口（与旧单文件版一致）：
    from rtf_merge import merge_rtf, fix_fonts_via_word
"""

__version__ = "1.1.2"
__all__ = ["merge_rtf", "fix_fonts_via_word", "__version__"]

from .merge_core import merge_rtf                     # noqa: E402
from .word_post import fix_fonts_via_word            # noqa: E402
