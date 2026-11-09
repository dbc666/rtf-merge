# -*- coding: utf-8 -*-
"""RTF 结构 lint 与资源表一致性校验。

SAS 宏的合并方式是"删首字符 {、删末字符 }、直接拼接"，因此每个文件都带着
完整的 \\rtf1 / \\fonttbl / \\colortbl 头。RTF 规范只允许一个 \\rtf1，当前
批次因 60 个文件字体表完全相同才没有出错；换模板（字体表/颜色表不同）就会
出现字体或颜色错乱，且宏本身不检测。此模块只做告警，不改变输出结构
（用户已确认保留 SAS 原语义）。
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

# 资源表所在的头部区域：前 400 行足够覆盖 SAS ODS 生成的 RTF 头
HEAD_LINES = 400


def _balanced_block(text: str, start: int) -> str:
    """从 start（指向 '{'）取出花括号配平的整块（跳过转义 \\{ \\}）。"""
    depth = 0
    i = start
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "\\" and i + 1 < n:
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
        i += 1
    return text[start:]


def extract_resource_tables(path: str | Path) -> dict[str, str]:
    """提取输入 RTF 的 \\fonttbl 与 \\colortbl 原始文本。"""
    out = {"fonttbl": "", "colortbl": ""}
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            head = "\n".join(fh.readline() for _ in range(HEAD_LINES))
    except OSError:
        return out
    for key, token in (("fonttbl", "\\fonttbl"), ("colortbl", "\\colortbl")):
        idx = head.find(token)
        if idx < 0:
            continue
        # 向前找最近的 '{'
        brace = head.rfind("{", 0, idx)
        if brace < 0:
            continue
        out[key] = _balanced_block(head, brace)
    return out


def resource_fingerprints(paths: list[Path]) -> tuple[dict[str, str], list[str]]:
    """计算各文件资源表指纹；返回 (指纹->文件列表, 告警列表)。"""
    groups: dict[str, list[str]] = {}
    no_font: list[str] = []
    for p in paths:
        tabs = extract_resource_tables(p)
        if not tabs["fonttbl"]:
            no_font.append(Path(p).name)
            continue
        body = tabs["fonttbl"] + "\x00" + tabs["colortbl"]
        h = hashlib.sha1(body.encode("utf-8", "replace")).hexdigest()[:12]
        groups.setdefault(h, []).append(Path(p).name)

    warnings: list[str] = []
    if len(groups) > 1:
        detail = "；".join(f"{h}: {len(v)} 个文件" for h, v in groups.items())
        warnings.append(
            f"输入文件的字体/颜色资源表不一致（{len(groups)} 组：{detail}）。"
            f"合并后 \\fN/\\cfN 编号可能错配，导致字体或颜色异常，"
            f"建议统一模板后重新导出。")
    if no_font:
        warnings.append(
            f"{len(no_font)} 个文件未找到 \\fonttbl（{', '.join(no_font[:5])}"
            f"{' 等' if len(no_font) > 5 else ''}），可能不是 ODS 生成的 RTF。")
    return groups, warnings


def check_structure(lines: list[str], name: str = "") -> list[str]:
    """结构 lint：首行 '{'、末行 '}'、花括号配平、\\rtf1 出现次数。"""
    issues: list[str] = []
    if not lines:
        return [f"{name}: 文件为空"]
    if not lines[0].lstrip().startswith("{"):
        issues.append(f"{name}: 首行不是 '{{' 开头，可能不是合法 RTF")
    if not lines[-1].rstrip().endswith("}"):
        issues.append(f"{name}: 末行不是 '}}' 结尾，可能不是合法 RTF")
    text = "\n".join(lines)
    opens = len(re.findall(r"(?<!\\)\{", text))
    closes = len(re.findall(r"(?<!\\)\}", text))
    if opens != closes:
        issues.append(f"{name}: 花括号不配平（{opens} 开 / {closes} 闭）")
    n_rtf1 = text.count("\\rtf1")
    if n_rtf1 != 1:
        issues.append(f"{name}: \\rtf1 出现 {n_rtf1} 次（单文件应恰好 1 次）")
    return issues


def check_structure_file(path: str | Path) -> dict:
    """对合并后的 RTF 文件做流式结构校验（不把 36MB 全文读成字符串）。

    返回 {"name", "lines", "braces": [开, 闭], "rtf1_count", "issues": [...]}。
    合并输出会有 N 个 \\rtf1（SAS 原语义，用户已确认保留），因此这里只把
    \\rtf1 次数记录为数据，不判为问题；真正要盯的是花括号配平与首末行。
    """
    p = Path(path)
    name = p.name
    issues: list[str] = []
    opens = closes = 0
    n_rtf1 = 0
    n_lines = 0
    first = last = None
    try:
        with open(p, "r", encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if first is None:
                    first = line
                last = line
                n_lines += 1
                # 转义 \{ \} 不计入配平
                opens += line.count("{") - line.count("\\{")
                closes += line.count("}") - line.count("\\}")
                n_rtf1 += line.count("\\rtf1")
    except OSError as exc:
        return {"name": name, "lines": 0, "braces": [0, 0],
                "rtf1_count": 0, "issues": [f"{name}: 无法读取（{exc}）"]}

    if n_lines == 0:
        issues.append(f"{name}: 文件为空")
    else:
        if not (first or "").lstrip().startswith("{"):
            issues.append(f"{name}: 首行不是 '{{' 开头")
        if not (last or "").rstrip().endswith("}"):
            issues.append(f"{name}: 末行不是 '}}' 结尾")
    if opens != closes:
        issues.append(f"{name}: 花括号不配平（{opens} 开 / {closes} 闭）")
    return {"name": name, "lines": n_lines, "braces": [opens, closes],
            "rtf1_count": n_rtf1, "issues": issues}
