# -*- coding: utf-8 -*-
"""合并核心：收集 / 排序 / 逐行合并 / 原子写盘 / 合并清单。

纯 Python 文本处理，不接触 Word，可独立单测。
忠实复刻 SAS 宏 rtf_merge4.1_utf8.sas 的合并语义（含 drop_sect_gap 扩展）。
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
from datetime import datetime
from pathlib import Path

from . import __version__

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


def _is_within(child: Path, parent: Path) -> bool:
    """child 是否位于 parent 目录（含子目录）之内。"""
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


# 唯一一类"使用提示"（不是产物缺陷）：输出被放在输入目录内。清单要区分
# 两类警告，用常量做前缀判定，避免文案改动后 all_ok 的口径悄悄变化。
INPUT_DIR_WARN_PREFIX = "输出文件位于输入目录内"

# warnings_detail 的分类标签：清单同时写「字符串数组 warnings」（对外旧字段，
# 保持兼容）和「对象数组 warnings_detail」（带 kind/level）。判定一律看 level，
# 不再用中文前缀做匹配——改一次措辞不会让 all_ok 的口径悄悄变化。
WARN_INPUT_DIR = "input_dir_overlap"      # 使用提示：输出落在输入目录内
WARN_ORDER_OUTPUT = "order_contains_output"   # 排序表列了输出文件本身
WARN_ORDER_PARTIAL = "order_partial_allowed"  # 排序表不全但被显式放行
WARN_RESOURCE = "resource_mismatch"       # 字体表/颜色表指纹不一致
WARN_ENCODING = "encoding_fallback"       # 实际编码与参数不同
WARN_RTF_HEAD = "rtf_header_invalid"      # 首行不是 '{
WARN_RTF_TAIL = "rtf_tail_invalid"        # 末行不是 '}'
WARN_STRUCTURE = "structure_invalid"      # 合并输出结构校验未通过
WARN_MANIFEST = "manifest_write_failed"   # 清单写失败（不影响合并结果）


def _warn(warnings: list[str], detail: list[dict], kind: str, text: str,
          level: str = "warn") -> None:
    """记一条警告：文字进 warnings（旧字段），分类进 warnings_detail。

    level 只有两档：info（使用提示，不计入 all_ok）与 warn（产物级问题）。
    目前只有「输出落在输入目录内」是 info，其余全部 warn，保证 warning_count
    与 all_ok 的判定结果和改造前完全一致。
    """
    warnings.append(text)
    detail.append({"kind": kind, "level": level, "text": text})


def normalize_outfile(outfile: str | os.PathLike) -> Path:
    """输出路径没有扩展名时补 .rtf。

    GUI 早已在 start() 里补；CLI 与 run_rtf_merge.bat 没有，导致同一句
    `--outfile "%OUTDIR%\合并结果"` 产出无扩展名文件、完成提示却写 .rtf。
    统一放到核心层，三条入口的行为才一致。
    """
    out = Path(outfile)
    if out.suffix == "":
        out = out.with_suffix(".rtf")
    return out


def merge_rtf(inpath: str | os.PathLike, outfile: str | os.PathLike,
              order: int = 0, pageyn: int = 1, delline=None,
              encoding: str = "utf-8", newline: str = "\r\n",
              log: logging.Logger | None = None,
              drop_sect_gap: bool = True,
              allow_partial_order: bool = False) -> dict:
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
      allow_partial_order: order=1 时，排序表与目录内 RTF 不一致（漏填/多填）
         默认【拒绝合并】；置 True 只告警并继续（不建议用于申报资料）。

    返回 dict：
      status  : 'merged' | 'template_created' | 'no_files' | 'too_few'
                | 'no_ordered_files' | 'order_incomplete' | 'empty_output' | 'error'
      outfile : 输出路径（merged 时）
      files   : 实际参与合并的文件列表
      encoding: 实际使用的输入编码
      warnings: 过程警告列表
    """
    logger = log or logging.getLogger("rtf_merge")
    warnings: list[str] = []
    warn_detail: list[dict] = []
    inpath = Path(inpath)
    # 输出路径不带扩展名时统一补 .rtf：CLI / GUI / bat 三条入口行为一致
    # （历史缺陷：bat 传 "%OUTDIR%\合并结果" 不带后缀，CLI 又不像 GUI 那样补，
    #   结果产出无扩展名文件，而完成提示里却写着 .rtf）
    outfile = normalize_outfile(outfile)

    # 1) 收集并排序文件（对应 dir /b/s + proc sort）
    files = collect_rtf_files(inpath, outfile)
    if not files:
        logger.error("目录 %s 下未找到任何 *.rtf 文件，合并退出。", inpath)
        return {"status": "no_files", "files": [], "warnings": warnings}
    # 目录内实际发现的文件（order=1 时用于与排序表做集合比对）
    discovered = list(files)

    # 输出落在输入目录内：本次产物已排除，但改名输出时历史产物仍会被收集
    try:
        if _is_within(Path(outfile).resolve(), inpath.resolve()):
            _warn(warnings, warn_detail, WARN_INPUT_DIR,
                  INPUT_DIR_WARN_PREFIX +
                  f"（{inpath}）。本次输出文件已自动排除，"
                  f"但如果下次改用别的文件名输出，上一次的合并结果会被当作输入"
                  f"再次合并（体积指数膨胀），建议输出到输入目录之外。",
                  level="info")
    except OSError:
        pass

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
        out_res = Path(outfile).resolve()
        resolved: list[Path] = []
        for fn, _od in ordered:
            p = Path(fn)
            if not p.is_absolute():
                p = inpath / p
            if not p.is_file():
                raise FileNotFoundError(f"排序文件中指定的 RTF 不存在: {p}")
            if p.resolve() == out_res:
                _warn(warnings, warn_detail, WARN_ORDER_OUTPUT,
                      f"排序表列出了输出文件本身，已跳过：{p}")
                continue
            resolved.append(p)

        # 与目录内实际文件做集合比对：漏填 order 的文件不会被合并，
        # 若不拦截，60 个文件只填 1 行会静默产出"只含 1 个文件"的成品。
        disc = {q.resolve() for q in discovered}
        picked = {q.resolve() for q in resolved}
        missing = sorted(disc - picked)
        extra = sorted(picked - disc)
        if missing or extra:
            names_missing = [q.name for q in missing][:10]
            names_extra = [q.name for q in extra][:10]
            msg = (f"排序表与目录内文件不一致：目录内 {len(disc)} 个 RTF，"
                   f"排序表有效条目 {len(picked)} 个。")
            if missing:
                msg += f" 未在排序表中（不会被合并）：{', '.join(names_missing)}"
                if len(missing) > 10:
                    msg += f" 等 {len(missing)} 个"
                msg += "。"
            if extra:
                msg += (f" 排序表中列出但目录内不存在：{', '.join(names_extra)}"
                        + (f" 等 {len(extra)} 个" if len(extra) > 10 else "") + "。")
            if not allow_partial_order:
                logger.error("%s 已中止（未写出任何文件）。请补全 order 列后重试，"
                             "或用 --allow-partial-order 显式放行。", msg)
                return {"status": "order_incomplete", "files": list(resolved),
                        "warnings": warnings, "order_file": str(order_file),
                        "missing": names_missing, "extra": names_extra}
            _warn(warnings, warn_detail, WARN_ORDER_PARTIAL,
                  msg + " 已按 --allow-partial-order 继续合并。")
        files = resolved
        logger.info("按排序文件合并，共 %d 个文件。", len(files))

    maxn = len(files)
    delline_fn = parse_delline(delline)

    # 2.5) 资源表一致性校验（只告警，不改变输出结构）
    # 宏的做法是每份 RTF 连头一起拼进正文，靠"字体表碰巧相同"才不出错；
    # 换模板即 \fN/\cfN 错配。此处提前把风险说出来。
    from .rtf_lint import resource_fingerprints
    _fp, lint_warnings = resource_fingerprints(files)
    for _text in lint_warnings:
        _warn(warnings, warn_detail, WARN_RESOURCE, _text)

    # 3) 逐文件合并（对应宏 %do i=1 %to &maxn 循环）
    # 关于末行"}"：SAS 宏把末行置空后由 put 写成 1 个空格（0x20），该空格行
    # 恰好落在下一个文件 \sect 之前，是 Word 里空白页的根源。drop_sect_gap=True
    # 时直接丢弃该行（RTF 允许空行/无行），分节符本身保留，每个表格仍从新页开始。
    want: list[str] = []
    used_encodings: set[str] = set()
    closed = False
    for i, f in enumerate(files, 1):
        logger.info("即将完成对文件：%s 的合并！", f)
        lines, used = read_rtf_lines(f, encoding, delline_fn)
        used_encodings.add(used)
        if used != encoding:
            _warn(warnings, warn_detail, WARN_ENCODING,
                  f"{f}: 实际使用编码 {used}（参数为 {encoding}）")
        if not lines:
            logger.warning("文件 %s 为空（0 行），跳过其内容。", f)
            continue

        # RTF 结构 lint（只告警，不中断，兼容 SAS 可能产出的非标准行）
        if not lines[0].lstrip().startswith("{"):
            _warn(warnings, warn_detail, WARN_RTF_HEAD,
                  f"{f}: 首行不是 '{{' 开头，可能不是合法 RTF。")
        if not lines[-1].rstrip().endswith("}"):
            _warn(warnings, warn_detail, WARN_RTF_TAIL,
                  f"{f}: 末行不是 '}}' 结尾，可能不是合法 RTF。")

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

    # 结构校验（花括号配平 / 首末行）：合并输出是交付物的起点，一旦配平错，
    # Word 可能打开失败或吞内容。结果写进清单的 rtf_validation，不只是打日志。
    # 校验不通过时状态直接改为 structure_invalid（文件与清单仍保留，便于排查），
    # 三条入口（CLI/GUI/bat）据此统一停在后处理之前——若只在此处记 warnings、
    # 状态仍返回 merged，GUI 会因只判断 status 而继续跑 Word 并弹"完成"。
    from .rtf_lint import check_structure_file
    rtf_validation = check_structure_file(out)
    if rtf_validation["issues"]:
        for issue in rtf_validation["issues"]:
            logger.error("结构校验失败：%s", issue)
        _warn(warnings, warn_detail, WARN_STRUCTURE,
              "合并输出的 RTF 结构校验未通过，详见清单 rtf_validation。")
        status = "structure_invalid"
    else:
        status = "merged"
        logger.info("结构校验通过：%d 行、花括号 %d/%d 配平、\\rtf1 %d 处。",
                    rtf_validation["lines"], rtf_validation["braces"][0],
                    rtf_validation["braces"][1], rtf_validation["rtf1_count"])

    result = {"status": status, "outfile": str(out), "files": files,
              "encoding": encoding, "encodings_used": sorted(used_encodings),
              "warnings": warnings, "warnings_detail": warn_detail,
              "rtf_validation": rtf_validation}
    try:
        result["manifest"] = _write_manifest(
            out, files, encoding=encoding, pageyn=pageyn, order=order,
            drop_sect_gap=drop_sect_gap, warnings=warnings,
            warnings_detail=warn_detail,
            resource_groups=_fp, rtf_validation=rtf_validation,
            encodings_used=sorted(used_encodings))
    except OSError as exc:
        _warn(warnings, warn_detail, WARN_MANIFEST,
              f"未能写出合并清单（{exc}），不影响合并结果。")
    return result


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _write_manifest(out: Path, files: list[Path], *, encoding: str, pageyn: int,
                    order: int, drop_sect_gap: bool,
                    warnings: list[str],
                    warnings_detail: list[dict] | None = None,
                    resource_groups: dict[str, list[str]] | None = None,
                    rtf_validation: dict | None = None,
                    encodings_used: list[str] | None = None) -> Path:
    """写合并清单：输入清单+SHA256、顺序、参数、输出哈希、时间戳。

    注册申报资料要求可重现：能回答“合并了哪些表、按什么顺序、用了什么参数”。
    后处理（docx 哈希/页数/字体/Word 版本/命令行）由 update_manifest_postprocess
    在本清单写好后追加，两段合起来才是完整留痕。
    """
    manifest = {
        "tool": "rtf_merge",
        "version": __version__,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "outfile": str(out),
        "outfile_size": out.stat().st_size,
        "outfile_sha256": _sha256(out),
        "file_count": len(files),
        "encoding": encoding,
        "encodings_used": encodings_used or [encoding],
        "pageyn": pageyn,
        "order": order,
        "drop_sect_gap": drop_sect_gap,
        "resource_table_groups": [
            {"fingerprint": h, "count": len(v), "files": v}
            for h, v in (resource_groups or {}).items()
        ],
        "rtf_validation": rtf_validation,
        "validation": {"rtf_structure_ok": not (rtf_validation or {}).get("issues")}
        if rtf_validation else None,
        "files": [{"path": str(f), "sha256": _sha256(f)} for f in files],
        "warnings": warnings,
        # 结构化副本：kind/level 决定口径（level != info 才计入 warning_count），
        # warnings 仍保留纯字符串，旧读取方不受影响。
        "warnings_detail": warnings_detail or [],
    }
    mpath = out.with_name(out.stem + "_manifest.json")
    # 原子写：与 update_manifest_postprocess 一致，避免半截清单冒充成品
    tmp = mpath.with_name(mpath.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, mpath)
    return mpath


def update_manifest_postprocess(manifest_path: str | os.PathLike,
                                post: dict) -> dict:
    """把后处理结果并入合并清单（原子写）。

    post 由 word_post.fix_fonts_via_word 返回，含 docx 路径/哈希/页数/
    空白页/字体覆盖率/各环节计数/Word 版本/耗时；调用方还可补
    command_line 与 mode。合并清单里的 validation 段在此收敛为总验收结论。
    """
    mp = Path(manifest_path)
    manifest = json.loads(mp.read_text(encoding="utf-8"))
    manifest["postprocess"] = post
    fonts = post.get("fonts") or {}
    validation = manifest.get("validation") or {}
    # 产物级警告才影响总验收：「输出在输入目录内」属使用提示（level=info），
    # 不计入。优先按 warnings_detail 的 level 判定；老清单（无该字段）回退到
    # 中文前缀匹配，保证读旧产物时口径不变。
    detail = manifest.get("warnings_detail")
    if detail:
        prod_warnings = [d for d in detail if d.get("level") != "info"]
    else:
        prod_warnings = [{"text": w, "level": "warn"}
                         for w in (manifest.get("warnings") or [])
                         if not w.startswith(INPUT_DIR_WARN_PREFIX)]
    merge_skipped = (post.get("step1") or {}).get("merge_skipped") or 0
    validation.update({
        "docx_sha256": post.get("docx_sha256"),
        "docx_size": post.get("docx_size"),
        "docx_pages": post.get("pages"),
        "docx_empty_pages": post.get("remaining_empty_pages"),
        "docx_rfonts": fonts.get("rfonts"),
        "docx_times_new_roman_ratio": fonts.get("tnr_ratio"),
        "docx_song_ratio": fonts.get("song_ratio"),
        "toc_domain": fonts.get("has_toc"),
        "toc_styles_fixed": post.get("toc_styles"),
        "word_version": post.get("word_version"),
        "warning_count": len(prod_warnings),
        "merge_skipped": merge_skipped,
        "all_ok": bool(
            validation.get("rtf_structure_ok")
            and (fonts.get("tnr_ratio") or 0) >= 0.95
            and (fonts.get("song_ratio") or 0) >= 0.95
            and fonts.get("has_toc")
            and not (post.get("remaining_empty_pages") or 0)
            and not prod_warnings
            and not merge_skipped
        ),
    })
    manifest["validation"] = validation
    tmp = mp.with_name(mp.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, mp)
    return manifest


