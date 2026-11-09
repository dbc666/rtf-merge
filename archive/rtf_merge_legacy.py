#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rtf_merge.py —— Python 版 RTF 合并工具
========================================
对应 SAS 宏 rtf_merge4.1_utf8.sas（%rtf_mergeV4_1），用于合并 SAS ODS RTF
语句输出的多个 *.rtf 文件为一个 RTF 文件，逐行忠实还原宏的合并逻辑：

  1. 递归收集 inpath 下所有 *.rtf（对应 SAS: dir /b/s "inpath\\*.rtf"）；
  2. 按文件名自然排序（对应 SAS: sortseq=linguistic(numeric_collation=on)，
     即 t_1_2 排在 t_1_10 之前）；
  3. order=1 时读取排序文件 file_order.xls/.xlsx（sheet=order，列 fname/order）
     人工指定合并顺序；文件不存在时自动生成模板并退出；
  4. 合并规则（逐条对应宏内注释）：
       - 除首个 RTF 外，其余 RTF 第一行的首个字符“{”删除；
       - 除最后一个 RTF 外，其余 RTF 最后一行的“}”删除（末行单独为“}”时
         置空行；末行为“...}”时去掉末字符）；
       - 每两个 RTF 之间插入分节行：
         \\sect\\sectd\\linex0\\endnhere\\pgwsxn15840\\pghsxn12240\\lndscpsxn\\
         headery1440\\footery1440\\marglsxn1440\\margrsxn1440\\margtsxn1440\\margbsxn1440
       - 最后一个 RTF 的末尾补回“}”；pageyn=1 时删除所有行中的 \\pgnrestart
         （解决页码错乱）；
       - 合并完成后整文档执行花式表格修正：若某行是 \\pard{\\par} 且上一行是
         {\\row}，则该行改为 \\pard（对应宏最后的 data want 步）；
  5. 输出：每行一个记录、行尾 CRLF、UTF-8 无 BOM（与 SAS Windows 端 put 行为一致）。

用法（命令行，参数与 SAS 宏一一对应）：
  python rtf_merge.py --inpath D:\\xx\\output_rtf --outfile D:\\xx\\all.rtf
  python rtf_merge.py --inpath D:\\xx\\output_rtf --outfile D:\\xx\\all.rtf ^
                      --order 1 --pageyn 1 --delline "85 86 87"

模块方式调用：
  from rtf_merge import merge_rtf
  result = merge_rtf(inpath=r"D:\\xx\\output_rtf", outfile=r"D:\\xx\\all.rtf",
                     order=1, pageyn=1, delline="85-87")

与 SAS 宏的差异（均为有意为之，请注意）：
  * order=1 且无排序文件时，宏生成 file_order.xls（Excel 97 二进制）；
    Python 版生成 file_order.xlsx（openpyxl 写入）。读取时优先
    file_order.xlsx（openpyxl 可直接读），仅当目录里只有宏生成的老版二进制
    file_order.xls 时回退读取，此时需先用 Excel 将其另存为 .xlsx（或安装
    pandas+xlrd）。模板列：fname（完整路径）、order（数值）。
  * 宏的 delline 参数是任意 SAS 语句（如 if del in (85 86 87) then delete;），
    Python 版支持行号列表/范围字符串（如 "85 86 87"、"85-87,90"）或
    callable(del行号, line)->bool（True 表示删除）。
  * 宏合并后还会调用 Font.vbs 用 Word 打开 RTF、另存 .docx 并把全文与页眉
    页脚字体设为宋体/Times New Roman；Python 版提供可选函数
    fix_fonts_via_word()（通过 PowerShell COM 调用 Word，无需 pywin32），
    命令行加 --fix-fonts 开启（修正字体 + 按内容自动调整正文表格），默认关闭。

作者：dbc 移植自 rtf_merge4.1_utf8.sas（2026-11-03）
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

__version__ = "1.1.0"

# 宏中插入两段 RTF 之间的分节行（rtf_merge4.1_utf8.sas 第 137 行）
SECTION_LINE = (
    r"\sect\sectd\linex0\endnhere\pgwsxn15840\pghsxn12240\lndscpsxn"
    r"\headery1440\footery1440\marglsxn1440\margrsxn1440\margtsxn1440\margbsxn1440"
)

# 与宏中 lrecl=30000 / informat line $30000 对应：单行超过 30000 字符截断
LRECL = 30000

ORDER_XLS = "file_order.xls"
ORDER_XLSX = "file_order.xlsx"


def natural_sort_key(s: str):
    """自然排序键，对应 SAS sortseq=linguistic(numeric_collation=on)。

    将字符串按数字段切分：数字段按数值比较、其余段按字符比较（忽略大小写），
    使 t_1_2.rtf 排在 t_1_10.rtf 之前。
    """
    return tuple(
        (0, int(chunk)) if chunk.isdigit() else (1, chunk.casefold())
        for chunk in re.split(r"(\d+)", s)
        if chunk != ""
    )


def collect_rtf_files(inpath: str | os.PathLike,
                     outfile: str | os.PathLike | None = None) -> list[Path]:
    """递归收集 inpath 下全部 *.rtf 文件并自然排序（对应 dir /b/s + proc sort）。

    排除：Word 锁文件（~$ 前缀）、隐藏文件（. 前缀）、以及本次输出文件自身
    （避免输出落在输入目录内时二次运行把自己的产物当输入）。
    """
    root = Path(inpath)
    out_resolved = None
    if outfile:
        try:
            out_resolved = Path(outfile).resolve()
        except OSError:
            out_resolved = None
    files: list[Path] = []
    for p in root.rglob("*.rtf"):
        if not p.is_file():
            continue
        name = p.name
        if name.startswith("~$") or name.startswith("."):
            continue
        if out_resolved is not None:
            try:
                if p.resolve() == out_resolved:
                    continue
            except OSError:
                pass
        files.append(p)
    files.sort(key=lambda p: natural_sort_key(os.path.abspath(str(p))))
    return files


def parse_delline(delline) -> callable | None:
    """解析 delline 参数，返回 (行号, 行内容)->bool（True=删除）的判断函数。

    支持：
      * None                                  —— 不删除
      * 字符串 "85 86 87" / "85-87" / "85-87,90" —— 按行号（1 起）删除
      * callable(del, line) -> bool            —— 自定义判断（对应宏内任意 SAS 语句）
    """
    if delline is None:
        return None
    if callable(delline):
        return delline
    nums: set[int] = set()
    text = str(delline).strip()
    for tok in re.split(r"[,\s]+", text):
        if not tok:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", tok)
        if not m:
            raise ValueError(
                f"delline 无法解析: {tok!r}，仅支持行号或范围，如 \"85 86 87\" 或 \"85-87\""
            )
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        if b < a:
            a, b = b, a
        nums.update(range(a, b + 1))
    return lambda del_no, line: del_no in nums


def read_rtf_lines(path: Path, encoding: str = "utf-8",
                   delline: callable | None = None) -> tuple[list[str], str]:
    """按 SAS 的读入方式读取一个 RTF 文件的行列表。

    SAS 语义：
      * infile ... truncover; input line $1-30000; —— 每条记录取前 30000 字符
      * line=strip(line);                          —— 去掉首尾空格（仅空格）
      * del=_n_; &delline.                         —— 按行号删除指定行
    返回 (行列表, 实际使用的编码)。编码优先按 encoding 参数，失败时回退
    gbk（文件头 ansicpg936），再回退 utf-8。
    """
    raw = Path(path).read_bytes()
    text = None
    used = None
    attempts = []
    for enc in (encoding, "gbk", "utf-8"):
        if enc not in attempts:
            attempts.append(enc)
    for enc in attempts:
        try:
            text = raw.decode(enc)
            used = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise UnicodeDecodeError(
            "rtf_merge", raw[:64], 0, len(raw[:64]),
            f"文件无法按 {attempts} 解码: {path}",
        )
    # 按记录切分：CRLF/CR/LF 均为记录分隔符；文件末尾单个换行不产生空记录
    parts = re.split(r"\r\n|\r|\n", text)
    if parts and parts[-1] == "":
        parts.pop()
    lines = [p[:LRECL].strip(" ") for p in parts]  # 先截断后 strip，与 SAS 一致
    if delline is not None:
        lines = [ln for n, ln in enumerate(lines, 1) if not delline(n, ln)]
    return lines, used


def fix_pgnrestart_line(line: str) -> str:
    """对应宏：if index(line,'\\pgnrestart\\') then
    line=compress(tranwrd(line,'\\pgnrestart',' '));
    即：含 \\pgnrestart\\ 的行，把 \\pgnrestart 替换为空格后删除该行全部空格。"""
    if "\\pgnrestart\\" in line:
        return line.replace("\\pgnrestart", " ").replace(" ", "")
    return line


def find_order_file(inpath: Path) -> Path | None:
    """查找人工排序文件：优先 file_order.xlsx（openpyxl 可读），
    其次 file_order.xls（SAS 宏生成的老版二进制格式，仅当无 xlsx 时采用）。"""
    for name in (ORDER_XLSX, ORDER_XLS):
        p = inpath / name
        if p.is_file():
            return p
    return None


def create_order_template(inpath: Path, files: list[Path]) -> Path:
    """生成排序模板（对应宏中 proc export 生成 file_order.xls 的步骤）。

    列：fname（完整路径）、order（留空待人工填写）。宏生成 .xls，本实现
    生成 .xlsx（openpyxl 不支持老版二进制 xls 写入）。
    """
    try:
        from openpyxl import Workbook
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "生成 file_order.xlsx 需要 openpyxl，请先安装：pip install openpyxl"
        ) from exc
    wb = Workbook()
    ws = wb.active
    ws.title = "order"
    ws.append(["fname", "order"])
    for f in files:
        ws.append([os.path.abspath(str(f)), None])
    out = inpath / ORDER_XLSX
    wb.save(out)
    return out


def read_order_file(order_file: Path) -> list[tuple[str, int | float | None]]:
    """读取排序文件（sheet=order，首行为列名 fname/order）。

    返回 [(fname, order), ...]，order 为 None 表示缺失（对应 SAS missing(order)，
    这些文件不参与合并）。
    """
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "读取排序文件需要 openpyxl，请先安装：pip install openpyxl"
        ) from exc
    try:
        wb = load_workbook(order_file, read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError(
            f"无法读取排序文件 {order_file}（{exc}）。"
            f"若为老版二进制 .xls，请用 Excel 另存为 .xlsx 后重试。"
        ) from exc
    ws = wb["order"] if "order" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows: list[tuple[str, int | float | None]] = []
    for idx, row in enumerate(ws.iter_rows(values_only=True)):
        if idx == 0:  # 表头（SAS dbdsopts=firstobs=1 将首行作为变量名）
            continue
        if row is None or len(row) == 0:
            continue
        fname = row[0]
        if fname is None or str(fname).strip() == "":
            continue
        order_val = row[1] if len(row) > 1 else None
        order: int | float | None = None
        if order_val is not None and str(order_val).strip() != "":
            try:
                order = int(order_val) if float(order_val).is_integer() else float(order_val)
            except (ValueError, TypeError):
                order = None  # 无法转数值视为缺失，不参与合并
        rows.append((str(fname).strip(), order))
    wb.close()
    return rows


def merge_rtf(inpath: str | os.PathLike, outfile: str | os.PathLike,
              order: int = 0, pageyn: int = 1, delline=None,
              encoding: str = "utf-8", newline: str = "\r\n",
              log: logging.Logger | None = None,
              drop_sect_gap: bool = True) -> dict:
    """合并 inpath 下所有 RTF 文件到 outfile。

    参数（对应 SAS 宏）：
      inpath  : 输入文件夹（递归查找 *.rtf）
      outfile : 输出合并后 RTF 文件
      order   : 0=按文件名自然排序（默认）；1=读取 file_order.xls/.xlsx 人工排序
      pageyn  : 1=修正页码（删除 \\pgnrestart，默认）；0=不修正
      delline : 删除指定行号，见 parse_delline
      encoding: 输入文件编码（默认 utf-8，失败回退 gbk）
      newline : 输出行尾（默认 CRLF，与 SAS Windows 端 put 一致）
      drop_sect_gap: True（默认）= 合并时"文件间分节符前的空行"（每个文件末行
         "}" 去掉后不再写成空格行），从源头消除空白页；False = 完全复现 SAS
         宏输出（会保留该空格行，Word 中可能生成空白页）。

    返回 dict：
      status  : 'merged' | 'template_created' | 'no_files' | 'too_few'
                | 'no_ordered_files' | 'empty_output' | 'error'
      outfile : 输出路径（merged 时）
      files   : 实际参与合并的文件列表
      encoding: 实际使用的输入编码
      warnings: 过程警告列表
    """
    logger = log or logging.getLogger("rtf_merge")
    warnings: list[str] = []
    inpath = Path(inpath)

    # 1) 收集并排序文件（对应 dir /b/s + proc sort）
    files = collect_rtf_files(inpath, outfile)
    if not files:
        logger.error("目录 %s 下未找到任何 *.rtf 文件，合并退出。", inpath)
        return {"status": "no_files", "files": [], "warnings": warnings}

    # 2) order=1：人工排序（对应宏 %if &order. eq 1 分支）
    if order == 1:
        order_file = find_order_file(inpath)
        if order_file is None:
            template = create_order_template(inpath, files)
            logger.warning(
                "宏参数order=1，但发现 %s 文件未存在，系统将自动生成排序模板 %s "
                "并退出当前程序的执行！请在人工完成 order 列的编辑后，再次执行本程序。",
                inpath / ORDER_XLS, template,
            )
            return {"status": "template_created", "order_template": str(template),
                    "files": files, "warnings": warnings}
        rows = read_order_file(order_file)
        excluded = [fn for fn, od in rows if od is None]
        for fn in excluded:
            logger.info("RTF文件(将不会参与RTF的合并): %s", fn)
        if len(rows) <= 1:  # 对应宏：观测数为 0 或 1 时不合并退出
            logger.warning("排序文件 %s 仅有 %d 条记录，不进行合并，退出。",
                           order_file, len(rows))
            return {"status": "too_few", "files": files, "warnings": warnings}
        ordered = [(fn, od) for fn, od in rows if od is not None]
        if not ordered:
            logger.error(
                "排序文件 %s 中 order 列全部为空（共 %d 条记录），没有任何文件"
                "可合并，已中止（未写出任何文件）。请填写 order 列后重试。",
                order_file, len(rows))
            return {"status": "no_ordered_files", "files": files,
                    "warnings": warnings, "order_file": str(order_file)}
        ordered.sort(key=lambda t: (t[1], natural_sort_key(t[0])))
        files = []
        for fn, _od in ordered:
            p = Path(fn)
            if not p.is_absolute():
                p = inpath / p
            if not p.is_file():
                raise FileNotFoundError(f"排序文件中指定的 RTF 不存在: {p}")
            files.append(p)
        logger.info("按排序文件合并，共 %d 个文件。", len(files))

    maxn = len(files)
    delline_fn = parse_delline(delline)

    # 3) 逐文件合并（对应宏 %do i=1 %to &maxn 循环）
    # 关于末行"}"：SAS 宏把末行置空后由 put 写成 1 个空格（0x20），该空格行
    # 恰好落在下一个文件 \sect 之前，是 Word 里空白页的根源。drop_sect_gap=True
    # 时直接丢弃该行（RTF 允许空行/无行），分节符本身保留，每个表格仍从新页开始。
    want: list[str] = []
    closed = False
    for i, f in enumerate(files, 1):
        logger.info("即将完成对文件：%s 的合并！", f)
        lines, used = read_rtf_lines(f, encoding, delline_fn)
        if used != encoding:
            warnings.append(f"{f}: 实际使用编码 {used}（参数为 {encoding}）")
        if not lines:
            logger.warning("文件 %s 为空（0 行），跳过其内容。", f)
            continue

        # RTF 结构 lint（只告警，不中断，兼容 SAS 可能产出的非标准行）
        if not lines[0].lstrip().startswith("{"):
            warnings.append(f"{f}: 首行不是 '{{' 开头，可能不是合法 RTF。")
        if not lines[-1].rstrip().endswith("}"):
            warnings.append(f"{f}: 末行不是 '}}' 结尾，可能不是合法 RTF。")

        if i == 1:
            if maxn > 1:
                # 首个 RTF：去掉末行“}”
                if lines[-1].strip(" ") == "}":
                    if drop_sect_gap:
                        lines.pop()
                    else:
                        lines[-1] = ""
                else:
                    lines[-1] = lines[-1].strip(" ")[:-1]
            want = list(lines)
        else:
            # 插入分节行（对应 proc sql insert into want）
            want.append(SECTION_LINE)
            # 去掉本文件末行“}”
            if lines[-1].strip(" ") == "}":
                if drop_sect_gap:
                    lines.pop()
                else:
                    lines[-1] = ""
            else:
                lines[-1] = lines[-1].strip(" ")[:-1]
            # 去掉本文件首行首个字符“{”（仅在首行仍是原始首行时执行，
            # 防 delline 删掉原首行后误删下一行首字符）
            if lines and lines[0].lstrip().startswith("{"):
                lines[0] = lines[0][1:]
            want.extend(lines)
            # 最后一个 RTF：末尾补回“}”并修正页码
            if i == maxn and want:
                want[-1] = want[-1].strip(" ") + "}"
                closed = True
                if pageyn == 1:
                    want = [fix_pgnrestart_line(ln) for ln in want]

    # 收尾兜底：最后文件为空被跳过、或只有 1 个文件时，补齐结构与页码
    if want and not closed:
        if maxn == 1:
            # 单文件：结构本已完整（末行 } 未删），只需修正页码
            if pageyn == 1:
                want = [fix_pgnrestart_line(ln) for ln in want]
        else:
            want[-1] = want[-1].strip(" ") + "}"
            if pageyn == 1:
                want = [fix_pgnrestart_line(ln) for ln in want]

    if not want:
        logger.error("合并结果为空（所有输入文件均无有效内容），已中止，未写出任何文件。")
        return {"status": "empty_output", "files": files, "warnings": warnings}

    # 4) 花式表格修正（对应宏最后的 data want 步，使用 lag(line) 语义：
    #    上一行取其原始值）
    prev: str | None = None
    for idx in range(len(want)):
        cur = want[idx]
        if cur.strip(" ") == r"\pard{\par}" and prev is not None \
                and prev.strip(" ") == r"{\row}":
            want[idx] = r"\pard"
        prev = cur

    # 5) 输出（对应 data _null_; set want; file ...; put line;）
    # SAS 的 put 对全空/全空白行写 1 个空格（实测 0x20），故空行映射为单个空格。
    out_lines = [" " if ln == "" else ln for ln in want]
    out = Path(outfile)
    if out.parent and not out.parent.exists():
        out.parent.mkdir(parents=True, exist_ok=True)
    # 原子写：先写 .tmp，校验非空后 os.replace 覆盖，避免半截文件冒充成品
    tmp = out.with_name(out.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8", newline="") as fh:
            if out_lines:
                fh.write(newline.join(out_lines))
                fh.write(newline)
        if tmp.stat().st_size == 0:
            raise OSError("写入结果为 0 字节")
        os.replace(tmp, out)
    except OSError as exc:
        try:
            tmp.unlink()
        except OSError:
            pass
        logger.error("ERROR: [rtf_merge] Unable to open: %s - check it exists! (%s)",
                     outfile, exc)
        return {"status": "error", "files": files, "warnings": warnings,
                "error": str(exc)}

    logger.info("已完成所有文件的合并！输出：%s", outfile)
    result = {"status": "merged", "outfile": str(out), "files": files,
              "encoding": encoding, "warnings": warnings}
    try:
        result["manifest"] = _write_manifest(
            out, files, encoding=encoding, pageyn=pageyn, order=order,
            drop_sect_gap=drop_sect_gap, warnings=warnings)
    except OSError as exc:
        warnings.append(f"未能写出合并清单（{exc}），不影响合并结果。")
    return result


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _write_manifest(out: Path, files: list[Path], *, encoding: str, pageyn: int,
                    order: int, drop_sect_gap: bool,
                    warnings: list[str]) -> Path:
    """写合并清单：输入清单+SHA256、顺序、参数、输出哈希、时间戳。

    注册申报资料要求可重现：能回答“合并了哪些表、按什么顺序、用了什么参数”。
    """
    manifest = {
        "tool": "rtf_merge.py",
        "version": __version__,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "outfile": str(out),
        "outfile_size": out.stat().st_size,
        "outfile_sha256": _sha256(out),
        "file_count": len(files),
        "encoding": encoding,
        "pageyn": pageyn,
        "order": order,
        "drop_sect_gap": drop_sect_gap,
        "files": [{"path": str(f), "sha256": _sha256(f)} for f in files],
        "warnings": warnings,
    }
    mpath = out.with_name(out.stem + "_manifest.json")
    with open(mpath, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    return mpath


def fix_fonts_via_word(outfile: str | os.PathLike) -> str:
    """合并后用 Word 修正字体并另存 .docx（对应宏尾部 Font.vbs 步骤）。

    实现（都不依赖 pywin32）：
      1) 用 PowerShell COM 调 Word 打开 RTF 另存为 .docx（Word 负责格式转换）；
      2) 同一次 Word 会话内，先把关键词单元格（N/BQL/Mean/Std/Median/CV(%)
         等 14 个）与其右侧空单元格水平合并（对应追加的
         MergeCellWithAdjacentEmptyRight VBA 功能）；
      3) 再按内容并铺满窗口自动调整正文所有表格（先 AutoFitBehavior
         wdAutoFitContent 再 wdAutoFitWindow，对应 AutoFitTablesContent +
         AutoFitTablesWindow VBA 功能；$doc.Tables 只含正文表格，页眉页脚
         中的表格不受影响）；
      4) 在“目录”下方插入目录（对应 InsertTableOfContents VBA 功能：
         文档无“目录”页时在开头创建；按大纲级别抓取 SAS ODS 标题，
         点状引导符、宋体 + Times New Roman、小四、蓝色、不加粗）；
      5) 第二个全新 Word 会话打开刚生成的 docx：先“温和删除”——只删除
         【空白页】页内的“仅含单个空格字符”段落中的空格字符（段落文本形态
         为[空格,分节符]；按用户要求【只删空格、保留段落标记与分节符】，
         分节符强制每表从新页开始故必须保留，且普通节边界不产生空白页的
         同类空格段落一律不动），随后仍按 FindEmptyPages 语义逐页把残余
         空白页展示给用户并询问是否删除
         （页面文本去控制字符后为空【且】页面不含 InlineShape 图片才算
         空白页，防止误删 SAS ODS 附图页）；删除改变页码，故删除后刷新
         目录字段，并把 TOC 1~9 样式本身设为蓝色（宋体+Times New
         Roman+小四+不加粗）且关闭样式自动重定义——否则用户下次在 Word
         右键“更新目录”时域重建会丢弃直接格式、按默认黑色样式渲染。
         必须在字体修正【之前】跑：实测 Word 重存会摧毁 python-docx
         写入的显式 rFonts；
      6) 用 python-docx 直接改写所有 run 的 rFonts：ascii/hAnsi/cs 设为
         Times New Roman（非中文字符）、eastAsia 设为宋体（中文字符），
         页眉/页脚字体设为 Times New Roman，并同步 Normal 样式与文档默认；
      7) 中文引号（“ ” ‘ ’）拆分为独立 run 并以宋体渲染（对应追加的
         ChangeChineseQuotesToSimSun VBA 功能）；
      8) 把源标题段落（带大纲级别的 SAS ODS 标题）的加粗从 run 直接格式
         迁移到新建的 RTF Heading 段落样式：Word 右键“更新目录”重建条目
         时会继承源标题的【直接】加粗，导致条目被加粗；样式级加粗不会被
         继承（多轮实验证实：条目/样式写 b=0 会被更新机制清除，源标题用
         样式加粗时条目无 b）。视觉效果不变。
    说明：实测 Word COM 的 Selection.Font.Name 对 RTF 导入文档不会写入
    ascii 字体（原 Font.vbs 在现代 Word 上同样存在此问题），因此字体修正
    采用 python-docx 直接改 XML，结果可验证、确定生效。
    返回另存后的 .docx 路径。
    """
    out = Path(outfile)
    if not out.is_file():
        raise RuntimeError(f"待修正字体的文件不存在: {out}")
    docx_out = out.with_suffix(".docx")

    # 第 1 趟（Word 会话 1）：RTF -> docx 转换 + 关键词合并 + AutoFit + 目录
    ps1 = r'''$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0        # wdAlertsNone：抑制一切弹窗（防隐藏实例挂死泄漏）
try {
    $doc = $word.Documents.Open($args[0])

    # ===== 第 1 步：MergeCellWithAdjacentEmptyRight（AutoFit 之前）=====
    # 对每个正文表格：单元格文本（去段落符/单元格结束符后 trim）精确等于
    # 关键词，且其右侧紧邻单元格为空时，把两者水平合并（对应追加的 VBA
    # 功能；$doc.Tables 只含正文表格，页眉/页脚表格不受影响）
    function Get-CellText([object]$c) {
        return ($c.Range.Text -replace [char]13, '' -replace [char]7, '').Trim()
    }
    $kwSet = @{}
    foreach ($kw in @('N','BQL','Mean','Std','Median','CV(%)','Min','Max',
                      'GeoMean','GeoCV(%)','Geo CV(%)','Q1','Q3','BLQ')) {
        $kwSet[$kw] = $true
    }
    $merged = 0
    foreach ($tbl in $doc.Tables) {
        $hits = New-Object System.Collections.ArrayList
        foreach ($cell in $tbl.Range.Cells) {
            if ($kwSet.ContainsKey((Get-CellText $cell))) {
                [void]$hits.Add(@{ row = [int]$cell.RowIndex; col = [int]$cell.ColumnIndex })
            }
        }
        # 从右到左处理，避免合并改变后续列索引（忠实还原 VBA 顺序）
        for ($i = $hits.Count - 1; $i -ge 0; $i--) {
            $r = $hits[$i].row
            $c = $hits[$i].col
            if ($c -ge 1 -and $c -lt $tbl.Columns.Count) {
                $right = $tbl.Cell($r, $c + 1)
                if ((Get-CellText $right) -eq '') {
                    $tbl.Cell($r, $c).Merge($right)
                    $merged++
                }
            }
        }
    }
    Write-Output ("MERGE_CELLS=" + $merged)

    # ===== 第 2 步：按内容并铺满窗口自动调整正文所有表格 =====
    # （wdAutoFitContent=1 后再 wdAutoFitWindow=2：先按内容分配列宽比例，
    # 再按页面宽度铺满，对应 AutoFitTablesContent + AutoFitTablesWindow
    # 两个 VBA 功能）；$doc.Tables 只含正文表格，页眉/页脚中的表格不在
    # 其中，天然不受影响
    $n = 0
    foreach ($tbl in $doc.Tables) {
        $tbl.AutoFitBehavior([int]1)
        $tbl.AutoFitBehavior([int]2)
        $n++
    }
    Write-Output ("AUTOFIT_TABLES=" + $n)

    # ===== 第 3 步：插入目录（对应 InsertTableOfContents VBA 功能）=====
    # 文档已有“目录”二字时在其后插入目录；否则在文档开头创建“目录”页：
    # “目录”标题（居中、宋体、16 磅加粗）+ 目录域 + 分页符（正文另起一页）。
    # 目录基于段落大纲级别抓取：包含所有非“正文文本”级别（Outline 1~9，
    # 对应 outlineLvl 0~8）的标题——不仅是 1~3 级，4 级及更深标题也进目录。
    # 目录内容格式：宋体 + Times New Roman + 小四(12 磅)
    # + 蓝色 + 不加粗 + 点状引导符 + 含页码。
    $tocRng = $null
    $created = $false
    # 定位“目录”页标题：兼容“目 录”（两字间可有任意个半角/全角空格或
    # 制表符）与紧邻的“目录”两种写法；都找不到才在文档开头新建目录页。
    $f1 = $doc.Content
    $f1.Find.Forward = $true
    $f1.Find.Wrap = 0              # wdFindStop
    $f1.Find.MatchWildcards = $true
    $f1.Find.Text = "目[ 　^t]@录"
    if ($f1.Find.Execute()) {
        $tocRng = $f1
        $tocRng.Collapse(0)        # 折叠到“目录”二字之后
    } else {
        $f2 = $doc.Content
        $f2.Find.Forward = $true
        $f2.Find.Wrap = 0
        $f2.Find.Text = "目录"
        if ($f2.Find.Execute()) {
            $tocRng = $f2
            $tocRng.Collapse(0)
        } else {
        # 无“目录”页：在文档开头创建。注意必须先 InsertParagraphBefore
        # 建新段落再写“目录”，否则文字会并入原第一段（其带大纲级别，
        # 会被目录自身抓取）
        $startRng = $doc.Range(0, 0)
        $startRng.InsertParagraphBefore()
        $titleRng = $doc.Range(0, 0)
        $titleRng.Text = "目录"
        $titleRng.ParagraphFormat.Alignment = 1   # wdAlignParagraphCenter
        $titleRng.ParagraphFormat.OutlineLevel = 10  # wdOutlineLevelBodyText
        $titleRng.Font.NameFarEast = "宋体"
        $titleRng.Font.NameAscii = "Times New Roman"
        $titleRng.Font.Size = 16
        $titleRng.Font.Bold = $true
        $titleRng.Collapse(0)
        $titleRng.InsertParagraphAfter()
        $titleRng.Collapse(0)
        $tocRng = $titleRng
        $created = $true
        }
    }
    if ($doc.TablesOfContents.Count -gt 0) {
        $doc.TablesOfContents.Item(1).Delete()
    }
    $toc = $doc.TablesOfContents.Add(
        $tocRng, $true, 1, 9, $false, "", $true, $true, "", $true, $true, $true)
    $toc.TabLeader = 1              # wdTabLeaderDots
    $toc.Update()
    $tocFont = $toc.Range.Font
    $tocFont.NameFarEast = "宋体"
    $tocFont.NameAscii = "Times New Roman"
    $tocFont.Size = 12
    $tocFont.Color = 16711680       # 蓝色：VBA RGB(0,0,255)=0xFF0000（COLORREF）
    $tocFont.Bold = $false
    if ($created) {
        $tocEnd = $toc.Range
        $tocEnd.Collapse(0)
        $tocEnd.InsertBreak(7)      # wdPageBreak：正文另起一页
    }
    Write-Output ("TOC_INSERTED=" + $doc.TablesOfContents.Count)
    $doc.SaveAs2($args[1], 16)          # wdFormatDocumentDefault (.docx)
    $doc.Close(0)
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
'''
    ps1_path = out.parent / ("_fix_fonts_" + out.stem + ".ps1")
    with open(ps1_path, "w", encoding="utf-8-sig", newline="\r\n") as fh:
        fh.write(ps1)
    shell = _find_powershell()
    try:
        proc = subprocess.run(
            [shell, "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(ps1_path), str(out.resolve()), str(docx_out.resolve())],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=1200,
            creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0),
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("Word 转换 docx 超时（>1200s），请确认 Word 未被占用。") from exc
    finally:
        try:
            ps1_path.unlink(missing_ok=True)
        except OSError:
            pass
    if proc.returncode != 0 or not docx_out.is_file():
        raise RuntimeError(
            "Word 转换 docx 失败：\n" + (proc.stderr or proc.stdout or "未知错误"))
    n_tables = None
    n_merged = None
    n_toc = None
    n_empty_del = None
    n_empty_skip = None
    for line in (proc.stdout or "").splitlines():
        if line.startswith("AUTOFIT_TABLES="):
            try:
                n_tables = int(line.split("=", 1)[1])
            except ValueError:
                pass
        elif line.startswith("MERGE_CELLS="):
            try:
                n_merged = int(line.split("=", 1)[1])
            except ValueError:
                pass
        elif line.startswith("TOC_INSERTED="):
            try:
                n_toc = int(line.split("=", 1)[1])
            except ValueError:
                pass
        elif line.startswith("EMPTY_PAGES_DELETED="):
            try:
                n_empty_del = int(line.split("=", 1)[1])
            except ValueError:
                pass
        elif line.startswith("EMPTY_PAGES_SKIPPED="):
            try:
                n_empty_skip = int(line.split("=", 1)[1])
            except ValueError:
                pass
    if n_empty_del is not None and n_empty_skip is not None:
        if n_empty_del or n_empty_skip:
            print(f"空白页处理：已删除 {n_empty_del} 个、跳过 {n_empty_skip} 个")
        else:
            print("未检测到空白页（文档中无仅含分页符/段落符/空格的页面）")
    if n_merged is not None:
        print(f"已将 {n_merged} 个关键词单元格与右侧空单元格合并（页眉页脚表格未处理）")
    if n_tables is not None:
        print(f"已按内容并铺满窗口自动调整 {n_tables} 个正文表格（页眉页脚表格未处理）")
    if n_toc is not None and n_toc > 0:
        print(f"已在“目录”下方插入目录（点状引导符、宋体+Times New Roman、小四、蓝色、"
              f"按大纲级别共抓取正文标题）")

    # ---- 第 1.5 步：python-docx 重写 docx（重新打包），修复 Word 打开挂起 ----
    # 实测：Word 从 RTF 直接 SaveAs2 生成的 docx，后续新 Word 实例 Open 会
    # 永久挂起（RPC 拒绝后 Quit 也卡死，导致“空白页处理超时”）；用
    # python-docx 读入并另存（内容原样保留、仅重新打包 zip），Word 即能
    # 正常打开（264 页文档 9 秒打开）。TOC 域/页眉页脚等 python-docx 不解析
    # 的部分按原始 XML 保留。
    try:
        from docx import Document as _RewriteDoc
        _rd = _RewriteDoc(docx_out)
        _rd.save(docx_out)
    except Exception as _rewrite_err:   # 重写失败不阻断后续（ps2 仍会尝试）
        print(f"提示：docx 重写失败（{_rewrite_err}），可能影响后续 Word 打开。")

    # ---- 第 2 趟（新 Word 会话）：查找并交互删除空白页（FindEmptyPages）----
    # 空白页判定必须基于用户实际打开文档时的分页结果，而同一文档在不同
    # Word 会话中的分页会略有差异（实测：管道内会话只检出 1 个空白页，
    # 另开会话在 AutoFit 后可见 5 个）。因此在最终 docx 上用全新会话检测，
    # 删除空白页后刷新目录字段页码（TOC Update）并重新应用目录格式。
    # 注意：此步必须在 python-docx 字体修正【之前】执行——实测 Word 重存
    # 会摧毁 python-docx 写入的显式 rFonts（docDefaults 被改写为
    # minorHAnsi/minorEastAsia 主题字体，中文会变“等线”），因此字体修正
    # 必须作为最后一次写盘。空白页判定：页面文本去掉分节符(Chr12)/段落符
    # (Chr13)/单元格符(Chr7)/空格/制表符后为空，且页面不含 InlineShape
    # 图片——SAS ODS 图形页在 Word 文本模型里为 "\r\f\r"（\f=Chr12 分节
    # 符），若按原 VBA 只去控制字符会把附图页误判为空白，必须排除。
    # 交互：逐个把空白页展示在用户屏幕前并询问是否删除；选“是”立即删除
    # 后继续展示下一个。测试可设 EMPTY_PAGE_AUTO=YES/NO 自动应答。
    ps2 = r'''$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0        # wdAlertsNone：抑制一切弹窗（防隐藏实例挂死泄漏）
Write-Output "STEP=word_started"
Start-Sleep -Seconds 3         # 等 Word COM 服务器就绪，防 RPC_E_CALL_REJECTED
try {
    # 1 参 Open（与 ps1 完全相同，PowerShell 5.1 多参 COM 绑定不稳——实测
    # 3 参/7 参 Open 会挂起或报“bool→Object”参数错误，1 参在 ps1 中稳定成功）。
    # 加 5 次重试兜底 RPC 拒绝。
    $doc = $null
    for ($i = 1; $i -le 5; $i++) {
        try {
            $doc = $word.Documents.Open($args[0])
            break
        } catch {
            Write-Output ("OPEN_RETRY_" + $i)
            if ($i -ge 5) { throw }
            Start-Sleep -Seconds 3
        }
    }
    Write-Output "STEP=doc_opened"

    # 空白页根源的空格段落（[空格, 分节符] 形态）只在【空白页】内才处理，
    # 处理逻辑见下方 Get-EmptyPages 定义之后的“温和删除”段；普通节边界的
    # 同类段落不产生空白页，一律保留（用户口径：只删空白页对应的空格）。

    function Get-EmptyPages([object]$d) {
        $total = $d.ComputeStatistics(2)        # wdStatisticPages
        $list = New-Object System.Collections.ArrayList
        for ($pn = 1; $pn -le $total; $pn++) {
            $rg = $d.GoTo(1, 1, $pn)            # wdGoToPage, wdGoToAbsolute
            $nx = $d.GoTo(1, 1, $pn + 1)
            if ($nx.Start -gt $rg.Start) {
                $rg.End = $nx.Start - 1
            } else {
                $rg.End = $d.Content.End        # 最后一页：覆盖到文档末尾
            }
            if ($rg.InlineShapes.Count -gt 0) { continue }   # 含图页非空白
            $txt = ($rg.Text -replace [char]13, '' -replace [char]7, '' `
                    -replace [char]12, '' -replace ' ', '' -replace "`t", '').Trim()
            if ($txt -eq '') { [void]$list.Add([int]$pn) }
        }
        return $list
    }

    function Get-PageRange([object]$d, [int]$pn) {
        $total = $d.ComputeStatistics(2)        # wdStatisticPages
        if ($pn -lt 1 -or $pn -gt $total) { return $null }
        $rg = $d.GoTo(1, 1, $pn)                # wdGoToPage, wdGoToAbsolute
        $nx = $d.GoTo(1, 1, $pn + 1)
        if ($nx.Start -gt $rg.Start) {
            $rg.End = $nx.Start - 1
        } else {
            $rg.End = $d.Content.End
        }
        return $rg
    }

    # ---- 温和删除：只删除【空白页】页内的“仅含空格字符”段落中的空格 ----
    # 段落文本形态为 [空格, 分节符]（SAS 空行映射的空格位于分节符边界）。
    # 只有落在空白页内的这类段落才是空白页的根源；普通节边界（上一节内容
    # 未满页）的同类段落不产生空白页，一律保留。按用户要求【只删空格字符、
    # 保留段落标记与分节符】——分节符强制每个表格从新页开始，必须保留。
    # 单轮遍历找出全部空白页（原版全页扫描一次，5-10 分钟可控；注意不要用
    # Sections 遍历定位——实测 $sec.Range.Paragraphs.Last 会卡死数十分钟），
    # 收集其中所有空格段落并按文档位置倒序删除（删除会改变后续位置，必须
    # 从后往前删）。删除后分页可能重排出新空白页，由主流程“最终复检”
    # （_recheck_empty_pages，全新 Word 实例）兜底。
    function Remove-AllSpacesInEmptyPages([object]$d) {
        $empties = @(Get-EmptyPages $d)
        if ($empties.Count -eq 0) { return 0 }
        $all = New-Object System.Collections.ArrayList
        foreach ($pn in $empties) {
            $pr = Get-PageRange $d ([int]$pn)
            if ($null -eq $pr) { continue }
            foreach ($par in $pr.Paragraphs) {
                $t = ($par.Range.Text -replace [char]13, '') -replace [char]7, '' -replace [char]12, ''
                if ($t -eq ' ') {
                    $r = $par.Range.Duplicate
                    if ($null -ne $r) { [void]$all.Add($r) }
                }
            }
        }
        if ($all.Count -eq 0) { return 0 }
        $n = 0
        $sorted = @($all | Sort-Object -Property Start -Descending)
        foreach ($r in $sorted) {
            if ($r.Text.Length -gt 0 -and [int]$r.Text[0] -eq 32) {
                $r.End = $r.Start + 1
                [void]$r.Delete()
                $n++
            }
        }
        return $n
    }

    $spaceDeleted = Remove-AllSpacesInEmptyPages $doc
    Write-Output ("STEP=spaces_deleted:" + $spaceDeleted)

    $auto = $env:EMPTY_PAGE_AUTO
    # NO 模式：不交互也不删整页——完全跳过 WinForms/全页扫描（历史超时点）。
    # YES/交互模式：全页扫描找残余空白页并（询问后）删除整页。
    if ($auto -eq 'NO') {
        $empties = @()
        $deletedPages = 0
        $skippedPages = 0
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        $empties = @(Get-EmptyPages $doc)
        Write-Output ("STEP=interactive_scan:" + $empties.Count)
        $deletedPages = 0
        $skippedPages = 0
        $idx = 0
        if (-not $auto) { $word.Visible = $true }   # 仅交互时显示，自动模式保持后台提速
        while ($idx -lt $empties.Count) {
        $pn = $empties[$idx]
        $null = $word.Selection.GoTo(1, 1, $pn)  # 把当前空白页展示在用户屏幕前
        if (-not $auto) { $word.Activate() }     # 交互模式才激活窗口（自动模式不可见）
        $title = "空白页检测（$($idx + 1)/$($empties.Count)）"
        $msg = "检测到第 $pn 页为空白页（仅含分页符/段落符/空格等非可见字符）。`r`n" +
               "是否删除该空白页？（删除后继续检查下一个空白页）"
        $yes = $false
        if ($auto -eq 'YES') { $yes = $true }
        elseif ($auto -eq 'NO') { $yes = $false }
        else {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                $msg, $title,
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            $yes = ($ans -eq [System.Windows.Forms.DialogResult]::Yes)
        }
        if ($yes) {
            $before = $empties.Count
            # 尝试 1：删除空白页整页范围（覆盖含空格/段落符/分节符的空白页）
            $rg = $doc.GoTo(1, 1, $pn)
            $nx = $doc.GoTo(1, 1, $pn + 1)
            if ($nx.Start -gt $rg.Start) {
                $rg.End = $nx.Start - 1
            } else {
                $rg.End = $doc.Content.End      # 最后一页：覆盖到文档末尾
            }
            $rg.Delete()
            $empties = @(Get-EmptyPages $doc)
            if ($empties.Count -ge $before) {
                # 尝试 2：删除页首前一字符（尾部 0 字符空白页由分页符/分节符
                # 或末尾段落符溢出产生——删除该字符让空白页并入上一页）
                $start = $doc.GoTo(1, 1, $pn).Start
                if ($start -gt 0) {
                    $pb = $doc.Range($start - 1, $start)
                    $c = [int][char]$pb.Text
                    if ($c -eq 12 -or $c -eq 13) { $pb.Delete() }
                }
                $empties = @(Get-EmptyPages $doc)
            }
            if ($empties.Count -lt $before) {
                $deletedPages++        # 删除成功，继续处理新的第一个空白页
            } else {
                $skippedPages++        # 删除无效（页数未变），跳过防死循环
                $idx++
            }
        } else {
            $skippedPages++
            $idx++
        }
    }
    }   # else（YES/交互模式）闭合
    $word.Visible = $false

    # 空白页删除会改变页码：刷新目录字段并重新应用格式
    function Update-TocBlue([object]$d) {
        if ($d.TablesOfContents.Count -gt 0) {
            $toc = $d.TablesOfContents.Item(1)
            # 先把 TOC 1~9 样式本身设为蓝色（宋体+Times New Roman+小四+不加粗）
            # 并关闭自动重定义——否则用户下次在 Word 里右键“更新目录”时，域重建
            # 会丢弃条目的直接格式并按默认（黑色）样式渲染，且 autoRedefine 会把
            # 样式改回默认。样式级设置使任何更新后的目录条目都保持蓝色（实测：
            # 按名称修改样式并关 AutomaticallyUpdate 后，更新域生成的条目无直接
            # 颜色、继承样式蓝色，TOC1/2/3 样式 XML 中 color=0000FF 持久生效）。
            foreach ($lv in 1..9) {
                try {
                    $st = $d.Styles.Item("TOC " + $lv)
                    $st.AutomaticallyUpdate = $false
                    $st.Font.NameFarEast = "宋体"
                    $st.Font.NameAscii = "Times New Roman"
                    $st.Font.Size = 12
                    $st.Font.Color = 16711680   # 蓝色：VBA RGB(0,0,255)=0xFF0000（COLORREF）
                    $st.Font.Bold = $false
                } catch {}
            }
            $toc.TabLeader = 1              # wdTabLeaderDots
            [void]$toc.Update()
            $tocFont = $toc.Range.Font
            $tocFont.NameFarEast = "宋体"
            $tocFont.NameAscii = "Times New Roman"
            $tocFont.Size = 12
            $tocFont.Color = 16711680       # 蓝色：VBA RGB(0,0,255)=0xFF0000（COLORREF）
            $tocFont.Bold = $false
        }
    }

    Update-TocBlue $doc
    Write-Output "STEP=toc_updated"
    $doc.SaveAs2($args[0], 16)          # wdFormatDocumentDefault (.docx)
    Write-Output "STEP=saved"
    $doc.Close(0)

    # ---- 复检循环：独立 Word 实例（全新分页状态，等同用户实际打开）----
    # 空白页检测对分页敏感：目录域刷新/重存后分页会变（实测同一实例内
    # 刚 SaveAs2 完重开，Word 复用内存态分页只检出 0~1 个；独立进程打开
    # 磁盘文件完整重排后可见 5 个）。因此复检每轮都启动新的 Word 实例，
    # 注：最终的“保存后复检”不再放在本脚本内——字体修正（python-docx，
    # 见主流程第 3 步）会再次改变分页，复检必须放到全部后处理完成之后、
    # 由主流程用独立 Word 实例执行（见 _recheck_empty_pages）。

    Write-Output ("SPACE_CHARS_DELETED=" + $spaceDeleted)
    Write-Output ("EMPTY_PAGES_DELETED=" + $deletedPages)
    Write-Output ("EMPTY_PAGES_SKIPPED=" + $skippedPages)
    Write-Output "STEP=closed"
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
'''
    ps2_path = out.parent / ("_fix_fonts_pass2_" + out.stem + ".ps1")
    with open(ps2_path, "w", encoding="utf-8-sig", newline="\r\n") as fh:
        fh.write(ps2)
    # 启动前检查 docx 是否被残留 Word 锁住：锁住则清理无窗口实例后重试一次
    if _docx_locked(docx_out):
        _kill_headless_word()
        time.sleep(3)
    if _docx_locked(docx_out):
        raise RuntimeError(
            "输出 docx 被占用（可能是残留 Word 实例），已尝试清理仍不可写，"
            "请关闭相关 Word 后重试。")
    try:
        proc2 = subprocess.run(
            [shell, "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(ps2_path), str(docx_out.resolve())],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=2700,
            creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0),
        )
    except subprocess.TimeoutExpired as exc:
        _kill_headless_word()
        tail = ""
        if isinstance(exc.stdout, str):
            tail = exc.stdout[-2000:]
        elif exc.stdout:
            tail = (exc.stdout or b"").decode("utf-8", errors="replace")[-2000:]
        raise RuntimeError(
            f"空白页处理超时（>2700s），请确认 Word 未被占用。\n最近输出：\n{tail}") from exc
    finally:
        try:
            ps2_path.unlink(missing_ok=True)
        except OSError:
            pass
    if proc2.returncode != 0 or not docx_out.is_file():
        _kill_headless_word()
        raise RuntimeError(
            "空白页处理失败：\n" + (proc2.stderr or proc2.stdout or "未知错误"))
    n_empty_del = None
    n_empty_skip = None
    for line in (proc2.stdout or "").splitlines():
        if line.startswith("EMPTY_PAGES_DELETED="):
            try:
                n_empty_del = int(line.split("=", 1)[1])
            except ValueError:
                pass
        elif line.startswith("EMPTY_PAGES_SKIPPED="):
            try:
                n_empty_skip = int(line.split("=", 1)[1])
            except ValueError:
                pass
    if n_empty_del is not None and n_empty_skip is not None:
        if n_empty_del or n_empty_skip:
            print(f"空白页处理：已删除 {n_empty_del} 个、跳过 {n_empty_skip} 个")
        else:
            print("未检测到空白页（文档中无仅含分页符/段落符/空格的页面）")
    n_space_del = None
    for line in (proc2.stdout or "").splitlines():
        if line.startswith("SPACE_CHARS_DELETED="):
            try:
                n_space_del = int(line.split("=", 1)[1])
            except ValueError:
                pass
    if n_space_del:
        print(f"已删除 {n_space_del} 个仅含空格的空白页根源段落中的空格字符"
              f"（保留段落标记与分节符，每个表格仍从新页开始）")

    # 第 3 步：python-docx 修正字体（确定性、可验证）
    _fix_docx_fonts(docx_out)
    # 第 4 步：中文引号（“ ” ‘ ’）拆分为独立 run 并以宋体渲染
    # （对应 ChangeChineseQuotesToSimSun VBA 功能）
    n_quotes = _fix_chinese_quotes(docx_out)
    if n_quotes is not None:
        print(f"已将 {n_quotes} 个中文引号字符设为宋体渲染（页眉页脚同样处理）")
    # 第 5 步：源标题加粗迁移到段落样式（避免用户更新目录后条目被加粗）
    n_headings = _fix_toc_source_bold(docx_out)
    if n_headings:
        print(f"已将 {n_headings} 个源标题的加粗迁移到段落样式"
              f"（更新目录后条目不再继承加粗）")
    # 第 6 步：最终复检（独立 Word 实例，按用户打开时的最终分页）
    # 空白页检测对分页敏感：字体修正等后处理会再次改变分页，ps2 内检测到
    # 的空白页状态在保存重开后会漂移（实测同一实例内刚保存完重开复用内存态
    # 分页，独立进程完整重排才看到真实空白页）。因此这里用全新 Word 实例
    # 打开最终文件，只删除仍落在空白页内的空格段落，至多 3 轮。
    n_recheck = _recheck_empty_pages(docx_out)
    if n_recheck:
        print(f"最终复检：另删 {n_recheck} 个空白页根源空格字符"
              f"（按最终分页，保留段落标记与分节符）")
    # 第 7 步：最终验收——确认最后一次写盘后字体与目录仍存在
    # （Word 重存会冲掉 python-docx 写入的显式 rFonts，此处兜底拦截）
    _verify_final_docx(docx_out)
    return str(docx_out)


def _recheck_empty_pages(docx_out: str | os.PathLike, max_rounds: int = 5,
                         timeout: int = 1200) -> int:
    """全部后处理完成后，用全新 Word 实例复检空白页，只删除空白页内的
    空格段落（保留段落标记与分节符），每轮结束后保存并重新检测，
    直到无空白页或无空格段落可删。返回累计删除的空格字符数。"""
    # PowerShell 可执行文件：优先 pwsh，回退 powershell
    import shutil
    pwsh = shutil.which("pwsh") or shutil.which("powershell")
    if not pwsh:
        return 0
    ps3 = r'''$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0        # wdAlertsNone：抑制一切弹窗（防隐藏实例挂死泄漏）
Start-Sleep -Seconds 3         # 等 Word COM 服务器就绪，防 RPC_E_CALL_REJECTED
try {
    $doc = $null
    for ($i = 1; $i -le 5; $i++) {
        try {
            $doc = $word.Documents.Open($args[0], $false, $false)
            break
        } catch {
            if ($i -ge 5) { throw }
            Start-Sleep -Seconds 3
        }
    }

    function Get-EmptyPages([object]$d) {
        $total = $d.ComputeStatistics(2)        # wdStatisticPages
        $list = New-Object System.Collections.ArrayList
        for ($pn = 1; $pn -le $total; $pn++) {
            $rg = $d.GoTo(1, 1, $pn)            # wdGoToPage, wdGoToAbsolute
            $nx = $d.GoTo(1, 1, $pn + 1)
            if ($nx.Start -gt $rg.Start) {
                $rg.End = $nx.Start - 1
            } else {
                $rg.End = $d.Content.End        # 最后一页：覆盖到文档末尾
            }
            if ($rg.InlineShapes.Count -gt 0) { continue }   # 含图页非空白
            $txt = ($rg.Text -replace [char]13, '' -replace [char]7, '' `
                    -replace [char]12, '' -replace ' ', '' -replace "`t", '').Trim()
            if ($txt -eq '') { [void]$list.Add([int]$pn) }
        }
        return $list
    }

    function Get-PageRange([object]$d, [int]$pn) {
        $total = $d.ComputeStatistics(2)
        if ($pn -lt 1 -or $pn -gt $total) { return $null }
        $rg = $d.GoTo(1, 1, $pn)
        $nx = $d.GoTo(1, 1, $pn + 1)
        if ($nx.Start -gt $rg.Start) {
            $rg.End = $nx.Start - 1
        } else {
            $rg.End = $d.Content.End
        }
        return $rg
    }

    # 只检测、不修改：合并阶段已在 RTF 源头去掉“文件间分节符前的空行”
    # （drop_sect_gap），空白页理论上不会再出现。此处仅统计并报告，绝不写盘
    # ——一旦在此 SaveAs2，Word 会重写全部 OOXML，把后续 python-docx 写入的
    # 显式字体（宋体/Times New Roman）冲成主题字体（实测：7658 个 rFonts 中
    # w:ascii="Times New Roman" 变成 0 处）。若确有残留，交由 Python 侧用
    # _recheck_empty_pages 处理，并在其后重做字体修正。
    $empties = @(Get-EmptyPages $doc)
    foreach ($pn in $empties) { Write-Output ("EMPTY_PAGE=" + $pn) }
    Write-Output ("EMPTY_PAGES_COUNT=" + $empties.Count)
    Write-Output ("SPACE_CHARS_DELETED=0")
    $doc.Close(0)
} finally {
    try { $doc.Close(0) } catch {}
    $word.Quit()
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
}
'''
    docx_path = os.fspath(docx_out)
    ps3_path = os.path.join(
        os.path.dirname(docx_path), "_recheck_empty_" + os.path.basename(docx_path)[:-5] + ".ps1")
    total = 0
    for _round in range(max_rounds):
        try:
            with open(ps3_path, "w", encoding="utf-8-sig", newline="\r\n") as fh:
                fh.write(ps3)
            proc = subprocess.run(
                [pwsh, "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-File", ps3_path, docx_path],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=timeout,
                creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0),
            )
        except subprocess.TimeoutExpired:
            _kill_headless_word()
            print(f"最终复检第 {_round + 1} 轮超时（>{timeout}s），提前结束")
            break
        finally:
            try:
                os.unlink(ps3_path)
            except OSError:
                pass
        if proc.returncode != 0:
            _kill_headless_word()
            print(f"最终复检第 {_round + 1} 轮失败：\n"
                  + (proc.stderr or proc.stdout or "未知错误"))
            break
        n = None
        for line in (proc.stdout or "").splitlines():
            if line.startswith("SPACE_CHARS_DELETED="):
                try:
                    n = int(line.split("=", 1)[1])
                except ValueError:
                    pass
        if n is None or n <= 0:
            # 只检测模式：n 恒为 0，但要把残留空白页数量暴露出来
            for line in (proc.stdout or "").splitlines():
                if line.startswith("EMPTY_PAGES_COUNT="):
                    try:
                        remaining = int(line.split("=", 1)[1])
                    except ValueError:
                        remaining = 0
                    if remaining:
                        print(f"警告：最终复检仍检测到 {remaining} 个空白页"
                              f"（源头方案未完全消除），请在 Word 中核对。")
            break
        total += n
        import time as _time
        _time.sleep(3)          # 等待上一 Word 实例完全退出、文件句柄释放
    return total


def _verify_final_docx(docx_path: str | os.PathLike, min_ratio: float = 0.95) -> None:
    """最终验收：确认最后一次写盘后，显式字体与目录域仍然存在。

    Word 的任何重存（SaveAs）都会把 python-docx 写入的显式 rFonts
    改写为主题字体（minorHAnsi/minorEastAsia），导致中文回退“等线”、
    西文非 Times New Roman，而程序此前仍报成功。此验收在全部写盘结束后
    直接解析 OOXML，字体覆盖率不达标即抛错，不再静默通过。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError("最终验收需要 python-docx：pip install python-docx") from exc

    ASC, EA = "Times New Roman", "宋体"
    doc = Document(docx_path)
    roots = [doc.part.element]
    for rel in list(doc.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            roots.append(rel.target_part.element)
    roots.append(doc.styles.element)

    total = ok_ascii = ok_ea = 0
    for root in roots:
        for rf in root.iter(qn("w:rFonts")):
            total += 1
            if rf.get(qn("w:ascii")) == ASC:
                ok_ascii += 1
            if rf.get(qn("w:eastAsia")) == EA:
                ok_ea += 1
    if total == 0:
        raise RuntimeError("最终验收失败：docx 中没有任何 rFonts 声明。")

    ratio_ascii = ok_ascii / total
    ratio_ea = ok_ea / total
    xml = doc.part.element.xml
    has_toc = "TOC \\o" in xml or "TOC \\\\o" in xml
    if ratio_ascii < min_ratio or ratio_ea < min_ratio:
        raise RuntimeError(
            f"最终验收失败：字体覆盖率不足——ascii=Times New Roman {ratio_ascii:.1%}"
            f"、eastAsia=宋体 {ratio_ea:.1%}（要求 ≥{min_ratio:.0%}，共 {total} 个 rFonts）。"
            f"最后一次写盘后字体被 Word 重存覆盖。")
    if not has_toc:
        raise RuntimeError("最终验收失败：docx 中未找到目录（TOC）域。")
    print(f"最终验收通过：rFonts {total} 个，Times New Roman 覆盖 {ratio_ascii:.1%}、"
          f"宋体覆盖 {ratio_ea:.1%}，目录域存在。")


def _fix_docx_fonts(docx_path: str | os.PathLike) -> None:
    """用 python-docx 把 docx 全文 run 的字体改为：
    中文(eastAsia)=宋体，非中文(ascii/hAnsi/cs)=Times New Roman；
    覆盖正文、表格、段落标记 pPr/rPr、页眉页脚（含各节独立页眉页脚部件）、
    styles.xml 全部样式与文档默认。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError(
            "字体修正需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    ASC = "Times New Roman"
    EA = "宋体"

    def set_rpr_fonts(rpr) -> None:
        """直接给一个 w:rPr 写入字体（用于样式/文档默认等无 run 的 rPr）。"""
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), ASC)
        rf.set(qn("w:eastAsia"), EA)

    def fix_root(root) -> int:
        """遍历部件根元素下所有字体声明并写入字体：
        1) 全部已存在的 w:rPr（覆盖 run rPr、段落标记 pPr/rPr、样式 rPr）；
        2) 无 rPr 的 run（未显式格式化）补建 rPr 并写入。
        返回处理数。"""
        n = 0
        for rpr in list(root.iter(qn("w:rPr"))):
            set_rpr_fonts(rpr)
            n += 1
        for r in root.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("w:rPr"), {})
                r.insert(0, rpr)
                set_rpr_fonts(rpr)
                n += 1
        return n

    doc = Document(docx_path)
    total = fix_root(doc.part.element)  # 正文（run + 段落标记 pPr/rPr，含全部表格）
    # 所有页眉/页脚部件（各节独立 headerN/footerN.xml 均覆盖）
    for rel in list(doc.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            total += fix_root(rel.target_part.element)
    # styles.xml：全部样式（含各段落样式的 rPr）与 docDefaults
    styles_el = doc.styles.element
    total += fix_root(styles_el)
    # 确保 Normal 样式与 docDefaults.rPrDefault.rPr 存在并写入字体
    normal = doc.styles["Normal"]
    set_rpr_fonts(normal.element.get_or_add_rPr())
    doc_defaults = styles_el.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = styles_el.makeelement(qn("w:docDefaults"), {})
        styles_el.insert(0, doc_defaults)
    rpr_def = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_def is None:
        rpr_def = doc_defaults.makeelement(qn("w:rPrDefault"), {})
        doc_defaults.insert(0, rpr_def)
    rpr = rpr_def.find(qn("w:rPr"))
    if rpr is None:
        rpr = rpr_def.makeelement(qn("w:rPr"), {})
        rpr_def.insert(0, rpr)
    set_rpr_fonts(rpr)
    doc.save(docx_path)

    # 自检：正文与页眉页脚中均应存在 ascii=Times New Roman 的 run
    check = Document(docx_path)
    ok = 0
    roots = [check.part.element]
    for rel in list(check.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            roots.append(rel.target_part.element)
    for root in roots:
        for rpr in root.iter(qn("w:rPr")):
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None and rf.get(qn("w:ascii")) == ASC \
                    and rf.get(qn("w:eastAsia")) == EA:
                ok += 1
    if ok == 0:
        raise RuntimeError("字体修正自检失败：docx 中未发现 ascii=Times New Roman 的 run。")


def _fix_chinese_quotes(docx_path: str | os.PathLike):
    """中文引号（“ ” ‘ ’）字符设为宋体渲染（对应 ChangeChineseQuotesToSimSun
    VBA 功能）。

    原理：Word 中 U+2018/2019/201C/201D 属西文标点范围，无 hint 时按
    ascii/hAnsi 字体（Times New Roman）渲染。把含引号的 run 拆分为独立
    run，并给引号 run 的 rFonts 设置 w:hint="eastAsia"（eastAsia 已由
    _fix_docx_fonts 统一设为宋体），即可让引号按宋体显示，其余字符不受
    影响。覆盖正文、表格、页眉页脚部件。

    返回处理的中文引号字符数；docx 中本就没有中文引号时返回 0。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from copy import deepcopy
    except ImportError as exc:
        raise RuntimeError(
            "中文引号处理需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    QUOTES = "“”‘’"

    def text_of(r):
        return "".join(t.text or "" for t in r.findall(qn("w:t")))

    def split_run(r, quotes=QUOTES) -> int:
        """把含引号的 run 拆成：普通文本 run + 引号 run（hint=eastAsia）。
        返回拆分出的引号字符数。含非文本子元素（w:br/w:tab/图形等）的 run
        不拆分（保持结构，这类 run 在 SAS ODS 表格中不会出现）。"""
        if any(c.tag != qn("w:t") and c.tag != qn("w:rPr")
               for c in r):
            return 0
        parts = []
        for t_el in r.findall(qn("w:t")):
            text = t_el.text or ""
            segs = re.split(r'([“‘”’])', text)
            for seg in segs:
                if seg:
                    parts.append(seg)
        if not parts:
            return 0
        # 全部是引号且 run 尚未 hint=eastAsia：直接原地设置
        rpr = r.find(qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if all(p in QUOTES for p in parts) and (
                rf is None or rf.get(qn("w:hint")) != "eastAsia"):
            if rpr is None:
                rpr = r.makeelement(qn("w:rPr"), {})
                r.insert(0, rpr)
            if rf is None:
                rf = rpr.makeelement(qn("w:rFonts"), {})
                rpr.insert(0, rf)
            rf.set(qn("w:hint"), "eastAsia")
            return len("".join(parts))
        # 混合文本：按段拆分
        n_quotes = 0
        # 清空原 run 的 w:t，随后按顺序重建（第一个片段留在原 run）
        ts = r.findall(qn("w:t"))
        for t_el in ts:
            r.remove(t_el)
        anchor = r  # 新 run 依次插到 anchor 之后
        first = True
        for seg in parts:
            if first:
                cur = r
                t_el = cur.makeelement(qn("w:t"), {})
                t_el.text = seg
                cur.append(t_el)
                first = False
            else:
                cur = deepcopy(r)
                # 去掉复制的 w:t（仅留 rPr 与结构）
                for t_el in cur.findall(qn("w:t")):
                    cur.remove(t_el)
                t_el = cur.makeelement(qn("w:t"), {})
                t_el.text = seg
                cur.append(t_el)
                anchor.addnext(cur)
            anchor = cur
            if seg in QUOTES:
                # 该 run 的 rPr/rFonts 设置 hint=eastAsia
                rpr = cur.find(qn("w:rPr"))
                if rpr is None:
                    rpr = cur.makeelement(qn("w:rPr"), {})
                    cur.insert(0, rpr)
                rf = rpr.find(qn("w:rFonts"))
                if rf is None:
                    rf = rpr.makeelement(qn("w:rFonts"), {})
                    rpr.insert(0, rf)
                rf.set(qn("w:hint"), "eastAsia")
                n_quotes += len(seg)
        return n_quotes

    def fix_root(root) -> int:
        n = 0
        for r in list(root.iter(qn("w:r"))):
            if text_of(r) and any(q in text_of(r) for q in QUOTES):
                n += split_run(r)
        return n

    doc = Document(docx_path)
    total = fix_root(doc.part.element)
    for rel in list(doc.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            total += fix_root(rel.target_part.element)
    doc.save(docx_path)

    # 自检：重开后统计 hint=eastAsia 且文本含引号的 run
    check = Document(docx_path)
    ok = 0
    roots = [check.part.element]
    for rel in list(check.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            roots.append(rel.target_part.element)
    for root in roots:
        for r in root.iter(qn("w:r")):
            t = text_of(r)
            if t and any(q in t for q in QUOTES):
                rpr = r.find(qn("w:rPr"))
                rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
                if rf is not None and rf.get(qn("w:hint")) == "eastAsia":
                    ok += 1
    if total and ok == 0:
        raise RuntimeError("中文引号处理自检失败：未发现 hint=eastAsia 的引号 run。")
    return total


def _fix_toc_source_bold(docx_path: str | os.PathLike) -> int:
    """修复“用户更新目录后条目被加粗”问题（对应 Word TOC 更新机制）。

    根因（多轮实验证实）：SAS ODS 输出的全部带大纲级别（outlineLvl）的
    源标题段落，其 run 都带直接 <w:b/><w:bCs/>（直接加粗）。Word 中右键
    “更新目录”重建条目时，会把源标题段落的【直接】加粗属性继承进条目
    run——对条目/样式写 b w:val="0" 都会被更新机制清除（v7 系列实验：
    81 条目 1197 run 写 b=0 后模拟更新，b=0 全部消失、条目重建为
    <w:b/>、TOC1/2/3 样式上的 b=0 也被清）；空白文档对照实验证明：源
    标题若用样式加粗（Heading 1），更新后条目无 b。即 Word 只继承
    “直接”字符格式，不继承样式级加粗。

    修复：把源标题的加粗从 run 直接格式迁移到段落样式（新建 RTF Heading
    段落样式，basedOn Normal，样式 rPr 带 <w:b/><w:bCs/> 与宋体/Times
    New Roman rFonts），并删除这些段落所有 run 的直接 <w:b/>/<w:bCs/>。
    视觉效果不变（样式级加粗与直接加粗渲染一致），大纲级别不受影响；
    Word 更新目录时条目不再继承直接加粗（v8 实验：迁移后模拟
    TOC.Update()，条目 run 无 <w:b/>，仅 rStyle+noProof）。

    必须在 python-docx 字体修正与中文引号之后、最后一次写盘时执行
    （Word 重存会摧毁 python-docx 写入的 XML 改动）。

    返回处理的源标题段落数。
    """
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.enum.style import WD_STYLE_TYPE
    except ImportError as exc:
        raise RuntimeError(
            "源标题加粗迁移需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    ASC = "Times New Roman"
    EA = "宋体"
    STYLE_SID = "RTFHeading"

    doc = Document(docx_path)

    # 1) 新建段落样式 RTF Heading（basedOn Normal，加粗；rFonts 与全局统一）
    sid = None
    for st in doc.styles.element.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == STYLE_SID:
            sid = STYLE_SID
            break
    if sid is None:
        st = doc.styles.add_style("RTF Heading", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        st.font.bold = True
        st.font.name = ASC
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), ASC)
        rf.set(qn("w:eastAsia"), EA)
        # 复杂文种加粗 bCs（与 b 配套，避免复杂文种渲染不一致）
        if rpr.find(qn("w:bCs")) is None:
            rpr.insert(0, rpr.makeelement(qn("w:bCs"), {}))
        sid = STYLE_SID

    # 2) 遍历所有带大纲级别的段落：设置段落样式 + 删除 run 直接 b/bCs
    n_paras = 0
    n_removed = 0
    for p in doc.part.element.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is None or ppr.find(qn("w:outlineLvl")) is None:
            continue
        ps = ppr.find(qn("w:pStyle"))
        if ps is None:
            ps = ppr.makeelement(qn("w:pStyle"), {})
            ppr.insert(0, ps)
        ps.set(qn("w:val"), sid)
        n_paras += 1
        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in ("w:b", "w:bCs"):
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
                    n_removed += 1
    doc.save(docx_path)

    # 自检：重开后 outline 段落数 > 0 且 run 无直接 b/bCs
    check = Document(docx_path)
    ol = []
    for p in check.part.element.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:outlineLvl")) is not None:
            ol.append(p)
    bad = 0
    for p in ol:
        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and (rpr.find(qn("w:b")) is not None
                                    or rpr.find(qn("w:bCs")) is not None):
                bad += 1
    if len(ol) == 0:
        raise RuntimeError("源标题加粗迁移自检失败：未找到带大纲级别的源标题段落。")
    if bad:
        raise RuntimeError(
            f"源标题加粗迁移自检失败：{len(ol)} 个源标题段落中仍有 {bad} 处直接 b/bCs。")
    if n_paras != len(ol):
        raise RuntimeError(
            f"源标题加粗迁移自检失败：处理 {n_paras} 个段落，重开后有 {len(ol)} 个。")
    return n_paras


def _find_powershell() -> str:
    """定位可用的 PowerShell 可执行文件（优先 pwsh 7，其次 Windows PowerShell）。"""
    for exe in ("pwsh", "powershell"):
        found = shutil.which(exe)
        if found:
            return found
    raise RuntimeError("未找到 PowerShell（pwsh/powershell），无法执行 Word 字体修正。")


def _kill_headless_word() -> None:
    """结束所有无窗口（MainWindowTitle 为空）的 Word 实例。

    这些实例是此前 PowerShell COM 调用因超时/异常被强杀后泄漏的，会锁住
    输出 docx 并导致后续 Documents.Open 被 RPC_E_CALL_REJECTED 拒绝。
    用户自己打开的 Word 一定有窗口标题，不受影响。
    """
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | "
             "Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=60,
            creationflags=(subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0),
        )
    except Exception:
        pass


def _docx_locked(path) -> bool:
    """以独占读写方式打开文件测试是否被占用。"""
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def _print_usage_hint() -> None:
    """双击 exe（无任何参数）时打印使用说明，避免窗口一闪而过。"""
    print("=" * 64)
    print(" rtf_merge —— RTF 合并工具（对应 SAS 宏 rtf_merge4.1_utf8.sas）")
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
    print("  2) 合并并做全部后处理（字体修正/表格调整/插入目录/删除空白页等）：")
    print("     rtf_merge.exe --inpath D:\\数据\\输出 --outfile D:\\数据\\merged.rtf "
          "--fix-fonts")
    print()
    print("提示：")
    print("  · --fix-fonts 需要本机装有 Microsoft Word（内部经 PowerShell COM 调用）")
    print("  · 也可以把命令写进 .bat 文件后双击运行（同目录已附 run_rtf_merge.bat 示例）")
    print("=" * 64)


def main(argv=None) -> int:
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
                             "交互删除空白页、中文引号以宋体渲染、在“目录”下方插入"
                             "目录，页眉页脚表格不参与单元格合并）")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s")
    logger = logging.getLogger("rtf_merge")

    try:
        result = merge_rtf(inpath=args.inpath, outfile=args.outfile, order=args.order,
                           pageyn=args.pageyn, delline=args.delline,
                           encoding=args.encoding)
    except (ValueError, FileNotFoundError, UnicodeDecodeError, RuntimeError) as exc:
        logger.error("ERROR: %s", exc)
        return 2

    for w in result.get("warnings", []):
        logger.warning("警告: %s", w)
    if result["status"] == "merged":
        logger.info("成功合并 %d 个 RTF 文件 -> %s",
                    len(result["files"]), result["outfile"])
        if args.fix_fonts:
            try:
                docx = fix_fonts_via_word(result["outfile"])
                logger.info("已用 Word 修正字体并另存为：%s", docx)
            except (RuntimeError, ImportError) as exc:
                logger.error("ERROR: 字体修正失败：%s（合并结果不受影响）", exc)
                return 3
        return 0
    # template_created / no_files / too_few / error 已由 merge_rtf 输出日志
    return 1


if __name__ == "__main__":
    sys.exit(main())
