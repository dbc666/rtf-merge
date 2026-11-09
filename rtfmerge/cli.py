# -*- coding: utf-8 -*-
"""命令行入口：参数与 SAS 宏一一对应。"""
from __future__ import annotations

import argparse
import logging
import sys
import time

from . import __version__
from .merge_core import merge_rtf, update_manifest_postprocess
from .word_post import fix_fonts_via_word


def _print_usage_hint() -> None:
    """双击 exe（无任何参数）时打印使用说明，避免窗口一闪而过。"""
    print("=" * 64)
    print(f" rtf_merge —— RTF 合并工具 v{__version__}"
          "（对应 SAS 宏 rtf_merge4.1_utf8.sas）")
    print("=" * 64)
    print("本程序是命令行工具，需要带参数运行；双击本窗口只会显示本说明。")
    print()
    print("用法：")
    print("    rtf_merge.exe --inpath <输入文件夹> --outfile <输出.rtf> "
          "[--fix-fonts] [--order 0|1] [--delline \"行号\"]")
    print()
    print("常用示例：")
    print("  1) 仅合并：")
    print("     rtf_merge.exe --inpath D:\\数据\\输出 --outfile D:\\数据\\merged.rtf")
    print("  2) 合并并做全部后处理（字体修正/表格调整/插入目录/引号等）：")
    print("     rtf_merge.exe --inpath D:\\数据\\输出 --outfile D:\\数据\\merged.rtf "
          "--fix-fonts")
    print()
    print("提示：")
    print("  · --fix-fonts 需要本机装有 Microsoft Word（内部经 PowerShell COM 调用）")
    print("  · 也可以把命令写进 .bat 文件后双击运行（同目录已附 run_rtf_merge.bat 示例）")
    print("  · 合并时会同时生成 xxx_manifest.json（输入清单+SHA256+顺序+参数，可追溯）")
    print("=" * 64)


def main(argv=None) -> int:
    """入口包装：无论正常结束、结构失败还是异常返回，都记一条总耗时。"""
    t0 = time.time()
    try:
        return _main(argv)
    finally:
        elapsed = time.time() - t0
        logging.getLogger("rtf_merge").info(
            "总耗时：%.1f 秒（约 %.1f 分钟）", elapsed, elapsed / 60.0)


def _main(argv=None) -> int:
    # 输出编码自适应：交互控制台（TTY）用系统本地代码页（中文 Windows 为
    # GBK，cmd 显示正常）；重定向/管道（日志落盘、被其它程序捕获）用
    # UTF-8，避免 PyInstaller 打包的 exe 在 pwsh 等 UTF-8 管道里出现乱码。
    try:
        for _s in (sys.stdout, sys.stderr):
            if hasattr(_s, "reconfigure"):
                _enc = "utf-8"
                if _s.isatty():
                    try:
                        import locale
                        _enc = locale.getpreferredencoding(False) or "utf-8"
                    except Exception:
                        _enc = "utf-8"
                _s.reconfigure(encoding=_enc, errors="replace")
    except Exception:
        pass

    if argv is None:
        argv = sys.argv[1:]
    if not argv:
        # 无参数（典型：双击 exe）：打印说明并停留，防止窗口一闪而过
        _print_usage_hint()
        try:
            if sys.stdin is not None and sys.stdin.isatty():
                input("按回车键退出...")
        except EOFError:
            pass
        return 0

    parser = argparse.ArgumentParser(
        prog="rtf_merge",
        description="Python 版 RTF 合并工具（对应 SAS 宏 rtf_merge4.1_utf8.sas）",
    )
    parser.add_argument("--inpath", required=True, help="输入文件夹（递归查找 *.rtf）")
    parser.add_argument("--outfile", required=True, help="输出合并后的 RTF 文件")
    parser.add_argument("--order", type=int, default=0, choices=(0, 1),
                        help="0=文件名自然排序（默认）；1=读取 file_order.xls/.xlsx 人工排序")
    parser.add_argument("--pageyn", type=int, default=1, choices=(0, 1),
                        help="1=修正页码（删除 \\pgnrestart，默认）；0=不修正")
    parser.add_argument("--delline", default=None,
                        help="删除指定行号，如 \"85 86 87\" 或 \"85-87\"（对应宏 delline 参数）")
    parser.add_argument("--encoding", default="utf-8", help="输入编码，默认 utf-8（失败回退 gbk）")
    parser.add_argument("--fix-fonts", action="store_true",
                        help="合并后用 Word 修正字体并另存 .docx（对应 Font.vbs 步骤，"
                             "经 PowerShell COM 调用 Word，无需 pywin32；合并关键词"
                             "单元格与右侧空单元格、按内容铺满窗口自动调整正文表格、"
                             "空白页检测、中文引号以宋体渲染、在“目录”下方插入"
                             "目录，页眉页脚表格不参与单元格合并）")
    parser.add_argument("--keep-sect-gap", action="store_true",
                        help="保留文件间分节符前的空行（完全复刻 SAS 宏原行为，"
                             "可能会产生空白页；默认已从源头丢弃该空行）")
    parser.add_argument("--empty-page-auto", choices=("NO", "YES", "ASK"), default="NO",
                        help="【已废弃，无实际作用】v1.1.2 起空白页环节为纯只读"
                             "检测（不删页、不弹窗）。该参数仅为兼容旧脚本保留，"
                             "传 YES/ASK 会打印一条提示，但行为与 NO 完全相同")
    parser.add_argument("--allow-partial-order", action="store_true",
                        help="order=1 时，排序表与目录内文件不一致也继续合并"
                             "（默认拒绝，避免漏合并；不建议用于申报资料）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    logger = logging.getLogger("rtf_merge")

    # 兼容保留参数的显式提示：老脚本传 YES/ASK 不会被拒（不破坏兼容），
    # 但一定要说清"不会删页"，避免有人以为打开这个开关就会清空白页。
    if args.empty_page_auto != "NO":
        logger.warning("提示: --empty-page-auto=%s 为已废弃参数，无实际作用；"
                       "空白页只做只读检测并把结果写进清单，不会删除任何页面。",
                       args.empty_page_auto)

    try:
        result = merge_rtf(inpath=args.inpath, outfile=args.outfile, order=args.order,
                           pageyn=args.pageyn, delline=args.delline,
                           encoding=args.encoding,
                           drop_sect_gap=not args.keep_sect_gap,
                           allow_partial_order=args.allow_partial_order)
    except (ValueError, FileNotFoundError, UnicodeDecodeError, RuntimeError) as exc:
        logger.error("ERROR: %s", exc)
        return 2

    for w in result.get("warnings", []):
        logger.warning("警告: %s", w)
    # 结构校验未通过（花括号不配平/首末行异常）时不再往后处理 Word：
    # 这种文件 Word 可能打不开或吞内容，宁可早停也不产出"看似成功"的交付物。
    # status 由 merge_rtf 统一给出（structure_invalid），CLI/GUI/bat 判据一致。
    if result["status"] == "structure_invalid":
        issues = (result.get("rtf_validation") or {}).get("issues")
        logger.error("ERROR: 合并输出结构校验未通过（文件已写出：%s），"
                     "已停止 Word 后处理：%s", result.get("outfile"), issues)
        return 4
    if result["status"] == "merged":
        logger.info("成功合并 %d 个 RTF 文件 -> %s",
                    len(result["files"]), result["outfile"])
        if result.get("manifest"):
            logger.info("合并清单：%s", result["manifest"])
        if args.fix_fonts:
            try:
                post = fix_fonts_via_word(result["outfile"])
                logger.info("已用 Word 修正字体并另存为：%s", post)
                if result.get("manifest"):
                    post["command_line"] = " ".join(sys.argv)
                    post["mode"] = "cli"
                    update_manifest_postprocess(result["manifest"], post)
                    logger.info("合并清单（含后处理留痕）：%s", result["manifest"])
            except (RuntimeError, ImportError) as exc:
                logger.error("ERROR: 字体修正失败：%s（合并结果不受影响）", exc)
                return 3
        return 0
    # template_created / no_files / no_ordered_files / order_incomplete /
    # empty_output / error 均已由 merge_rtf 输出日志
    return 1
