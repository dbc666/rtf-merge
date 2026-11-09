# -*- coding: utf-8 -*-
"""docx OOXML 层修正：字体 / 中文引号 / 标题加粗迁移 / 最终验收。

全部通过 python-docx + lxml 直接改 XML，是最后一次写盘的执行者。
"""
from __future__ import annotations

import gc
import os
import re
from pathlib import Path

# 内存约定：15MB / 278 页的 docx 展开成 lxml DOM 约 470MB（实测），而每个
# 修正函数都是「save 后立刻重新打开自检」。若不在 save 后先释放主 DOM，两个
# DOM 会同时存活、峰值翻倍到 950MB——这是 GUI 完成任务后仍占 2GB 的主因。
# 因此每个 _fix_* 在 save 之后都 `del doc; gc.collect()` 再复检。

def _verify_final_docx(docx_path: str | os.PathLike, min_ratio: float = 0.95) -> dict:
    """最终验收：确认最后一次写盘后，显式字体与目录域仍然存在。

    Word 的任何重存（SaveAs）都会把 python-docx 写入的显式 rFonts
    改写为主题字体（minorHAnsi/minorEastAsia），导致中文回退“等线”、
    西文非 Times New Roman，而程序此前仍报成功。此验收在全部写盘结束后
    直接解析 OOXML，字体覆盖率不达标即抛错，不再静默通过。

    返回统计（供合并清单留痕）：
      {"rfonts": 总数, "tnr_ratio": ..., "song_ratio": ..., "has_toc": bool}
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
    del doc, xml, roots
    gc.collect()
    return {"rfonts": total, "tnr_ratio": round(ratio_ascii, 4),
            "song_ratio": round(ratio_ea, 4), "has_toc": bool(has_toc)}


def _apply_fonts(doc) -> int:
    """在【已打开】的 doc 上写入中西文双字体，不落盘。

    中文(eastAsia)=宋体，非中文(ascii/hAnsi/cs)=Times New Roman；
    覆盖正文、表格、段落标记 pPr/rPr、页眉页脚（含各节独立页眉页脚部件）、
    styles.xml 全部样式与文档默认。返回处理的 rPr 数。

    单独拆出纯修改部分，是为了支持"一次 load → 多次修改 → 一次 save"
    （见 apply_all_fixes）；`_fix_docx_fonts` 仍保留为单步入口。
    """
    from docx.oxml.ns import qn

    ASC = "Times New Roman"
    EA = "宋体"
    # 符号字体白名单：这些 rFonts 承载的是符号（项目符号、勾选、箭头等），
    # 覆盖成 Times New Roman 会让符号变成乱码，必须原样保留。
    SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3",
                    "webdings", "marlett"}

    def is_symbol_fonts(rf) -> bool:
        if rf is None:
            return False
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            v = rf.get(qn(attr))
            if v and v.strip().lower() in SYMBOL_FONTS:
                return True
        return False

    def set_rpr_fonts(rpr) -> bool:
        """直接给一个 w:rPr 写入字体（用于样式/文档默认等无 run 的 rPr）。
        命中符号字体白名单时不改写，返回 False。"""
        rf = rpr.find(qn("w:rFonts"))
        if is_symbol_fonts(rf):
            return False
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), ASC)
        rf.set(qn("w:eastAsia"), EA)
        return True

    def fix_root(root) -> int:
        """遍历部件根元素下所有字体声明并写入字体：
        1) 全部已存在的 w:rPr（覆盖 run rPr、段落标记 pPr/rPr、样式 rPr）；
        2) 无 rPr 的 run（未显式格式化）补建 rPr 并写入。
        符号字体（Symbol/Wingdings/Webdings 等）原样保留。
        返回处理数。"""
        n = 0
        for rpr in list(root.iter(qn("w:rPr"))):
            if set_rpr_fonts(rpr):
                n += 1
        for r in root.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("w:rPr"), {})
                r.insert(0, rpr)
                if set_rpr_fonts(rpr):
                    n += 1
        return n

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
    return total


def _fix_docx_fonts(docx_path: str | os.PathLike) -> None:
    """单步入口（兼容旧调用与单测）：load → 改字体 → save → 自检。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError(
            "字体修正需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    ASC, EA = "Times New Roman", "宋体"
    doc = Document(docx_path)
    _apply_fonts(doc)
    doc.save(docx_path)
    del doc                      # 先释放主 DOM，避免与自检 DOM 并存（峰值翻倍）
    gc.collect()

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
    del check
    gc.collect()


def _apply_quotes(doc) -> int:
    """在【已打开】的 doc 上把中文引号拆成 hint=eastAsia 的独立 run，不落盘。

    原理：Word 中 U+2018/2019/201C/201D 属西文标点范围，无 hint 时按
    ascii/hAnsi 字体（Times New Roman）渲染。把含引号的 run 拆分为独立
    run，并给引号 run 的 rFonts 设置 w:hint="eastAsia"（eastAsia 已由
    _apply_fonts 统一设为宋体），即可让引号按宋体显示，其余字符不受
    影响。覆盖正文、表格、页眉页脚部件。

    返回处理的中文引号字符数；docx 中本就没有中文引号时返回 0。
    """
    from copy import deepcopy
    from docx.oxml.ns import qn

    QUOTES = "“”‘’"
    ATOM_TAGS = (qn("w:br"), qn("w:tab"))   # 允许的结构性原子子元素

    def text_of(r):
        return "".join(t.text or "" for t in r.findall(qn("w:t")))

    def set_quote_hint(cur) -> None:
        """给一个 run 的 rPr/rFonts 设 w:hint=eastAsia（引号按宋体渲染）。"""
        rpr = cur.find(qn("w:rPr"))
        if rpr is None:
            rpr = cur.makeelement(qn("w:rPr"), {})
            cur.insert(0, rpr)
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        rf.set(qn("w:hint"), "eastAsia")

    def split_run(r, quotes=QUOTES) -> int:
        """把含引号的 run 拆成：普通文本 run + 引号 run（hint=eastAsia）。
        返回拆分出的引号字符数。

        w:br / w:tab 作为“原子段”保留（各自独立成一个 run，不参与文本拆分）；
        含图形/域等其它非文本子元素的 run 不拆分（避免破坏结构）。"""
        children = [c for c in r if c.tag != qn("w:rPr")]
        if any(c.tag != qn("w:t") and c.tag not in ATOM_TAGS for c in children):
            return 0
        parts = []          # [("text", 文本) | ("atom", 元素)]
        for c in children:
            if c.tag in ATOM_TAGS:
                parts.append(("atom", c))
                continue
            for seg in re.split(r'([“‘”’])', c.text or ""):
                if seg:
                    parts.append(("text", seg))
        if not parts:
            return 0
        rpr = r.find(qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        # 纯引号文本（无原子段）且尚未 hint：原地设置，不拆分
        if (all(k == "text" for k, _ in parts)
                and all(v in QUOTES for _, v in parts)
                and (rf is None or rf.get(qn("w:hint")) != "eastAsia")):
            set_quote_hint(r)
            return len("".join(v for _, v in parts))

        n_quotes = 0
        # 清空原 run 的全部非 rPr 子元素，随后按 parts 顺序重建
        for c in children:
            r.remove(c)
        anchor = r
        first = True
        for kind, val in parts:
            if first:
                cur = r
                first = False
            else:
                cur = deepcopy(r)
                for c in list(cur):
                    if c.tag != qn("w:rPr"):
                        cur.remove(c)
                anchor.addnext(cur)
            anchor = cur
            if kind == "text":
                t_el = cur.makeelement(qn("w:t"), {})
                t_el.text = val
                if val != val.strip():
                    t_el.set(qn("xml:space"), "preserve")
                cur.append(t_el)
                if val in QUOTES:
                    set_quote_hint(cur)
                    n_quotes += len(val)
            else:
                cur.append(deepcopy(val))
        return n_quotes

    def fix_root(root) -> int:
        n = 0
        for r in list(root.iter(qn("w:r"))):
            if text_of(r) and any(q in text_of(r) for q in QUOTES):
                n += split_run(r)
        return n

    total = fix_root(doc.part.element)
    for rel in list(doc.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            total += fix_root(rel.target_part.element)
    return total


def _fix_chinese_quotes(docx_path: str | os.PathLike):
    """单步入口（兼容旧调用与单测）：load → 改引号 → save → 自检。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError(
            "中文引号处理需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    QUOTES = "“”‘’"

    def text_of(r):
        return "".join(t.text or "" for t in r.findall(qn("w:t")))

    doc = Document(docx_path)
    total = _apply_quotes(doc)
    doc.save(docx_path)
    del doc                      # 先释放主 DOM，避免与自检 DOM 并存（峰值翻倍）
    gc.collect()

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
    del check
    gc.collect()
    return total


def _apply_toc_styles(doc) -> int:
    """在【已打开】的 doc 上把目录样式 TOC 1~9 设成：宋体 + Times New Roman、
    小四、蓝色、不加粗（不落盘）。

    为什么必须在 XML 层做（而不是再让 Word 跑一遍）：
      Word 的“更新目录”会用样式重建条目，条目上的**直接格式**会被丢掉，
      只有样式级设置能存活。此前这段由 step2_find_empty.ps1 的 Update-TocBlue
      在 Word 里改样式并 SaveAs2；但那次 SaveAs2 会重写整个 OOXML，把
      python-docx 写入的显式字体冲成主题字体（实测 7658 个 rFonts 中
      w:ascii="Times New Roman" 变成 0 处）。移到 XML 层后既不丢蓝色，
      也不再有“Word 重存摧毁字体”的顺序耦合。

    同时关闭 w:autoRedefine（Word 的“自动更新样式”），否则用户改一次目录
    外观，样式就被反向覆盖回默认。

    返回处理的样式个数。
    """
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    ASC, EA = "Times New Roman", "宋体"
    BLUE = RGBColor(0x00, 0x00, 0xFF)          # 目录条目蓝色

    n = 0
    for i in range(1, 10):
        sid = f"TOC{i}"
        st = next((s for s in doc.styles
                   if (s.element.get(qn("w:styleId")) or "") == sid), None)
        if st is None:
            # 文档里没有这一级目录样式（正文目前只用到前几级标题）。这里补建占位
            # 样式：以后正文出现 H4~H9 标题时，目录条目直接继承蓝色/小四/不加粗，
            # 而不是回落到 Word 默认的黑色。
            st = doc.styles.add_style(sid, WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = doc.styles["Normal"]
        el = st.element
        # 关掉“自动更新样式”，否则用户更新目录后样式会被反向覆盖
        for sub in el.findall(qn("w:autoRedefine")):
            el.remove(sub)
        st.font.name = ASC                     # ascii + hAnsi
        st.font.size = Pt(12)                  # 小四
        st.font.bold = False                   # 显式 w:b val="0"
        st.font.color.rgb = BLUE               # w:color val="0000FF"
        rpr = el.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        rf.set(qn("w:eastAsia"), EA)
        rf.set(qn("w:cs"), ASC)
        n += 1
    return n


def _fix_toc_styles(docx_path: str | os.PathLike) -> int:
    """单步入口（兼容旧调用与单测）：load → 改目录样式 → save → 自检。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError(
            "目录样式修正需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    doc = Document(docx_path)
    n = _apply_toc_styles(doc)
    doc.save(docx_path)
    del doc                      # 先释放主 DOM，避免与自检 DOM 并存（峰值翻倍）
    gc.collect()

    # 自检：重开后确认目录样式确实带上了蓝色/字号/不加粗
    check = Document(docx_path)
    bad = []
    for st in check.styles:
        sid = st.element.get(qn("w:styleId")) or ""
        if not re.fullmatch(r"TOC[1-9]", sid):
            continue
        rpr = st.element.find(qn("w:rPr"))
        color = rpr.find(qn("w:color")) if rpr is not None else None
        sz = rpr.find(qn("w:sz")) if rpr is not None else None
        b = rpr.find(qn("w:b")) if rpr is not None else None
        if (color is None or (color.get(qn("w:val")) or "").upper() != "0000FF"
                or sz is None or sz.get(qn("w:val")) != "24"
                or b is None or b.get(qn("w:val")) in (None, "1", "true")):
            bad.append(sid)
    if n == 0:
        raise RuntimeError("目录样式修正失败：docx 中未找到 TOC 1~9 样式"
                           "（可能是目录尚未插入）。")
    if bad:
        raise RuntimeError(f"目录样式修正自检失败：{', '.join(bad)} 未写入蓝色/小四/不加粗。")
    del check
    gc.collect()
    return n


def _apply_toc_source_bold(doc) -> int:
    """在【已打开】的 doc 上修复“用户更新目录后条目被加粗”问题（不落盘）。

    根因：SAS ODS 输出的全部带大纲级别（outlineLvl）的源标题段落，其 run
    都带直接 <w:b/><w:bCs/>；Word 右键“更新目录”重建条目时会把源标题的
    【直接】字符格式继承进条目 run（写 b=0 会被清），而样式级加粗不会被
    继承。修复即：把加粗从 run 直接格式迁移到段落样式，并删除 run 的直接
    b/bCs，视觉效果不变。

    本版改进（v1.1.2）：
      · 只处理“run 本来带直接加粗”的段落，不带直接加粗的段落不动；
      · 段落已有非 Normal 的 pStyle 时，直接给该样式补 b/bCs（保留原
        pStyle，不把 Heading1/2 等样式覆盖成 RTF Heading）；
      · 只有无 pStyle 或 Normal 样式的标题才改用 RTF Heading 样式；
      · 文档里没有带大纲级别的标题时不抛错，返回 0。

    必须在 python-docx 字体修正与中文引号之后、最后一次写盘时执行
    （Word 重存会摧毁 python-docx 写入的 XML 改动）。

    返回处理的源标题段落数。
    """
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn

    ASC = "Times New Roman"
    EA = "宋体"
    STYLE_SID = "RTFHeading"
    NORMAL_SIDS = {"Normal", "a", "a0", "Standard"}

    def ensure_style_bold(style_el) -> None:
        """给样式 rPr 补 <w:b/>/<w:bCs/> 与统一字体（幂等）。"""
        rpr = style_el.find(qn("w:rPr"))
        if rpr is None:
            rpr = style_el.makeelement(qn("w:rPr"), {})
            style_el.append(rpr)
        if rpr.find(qn("w:b")) is None:
            rpr.insert(0, rpr.makeelement(qn("w:b"), {}))
        if rpr.find(qn("w:bCs")) is None:
            rpr.insert(0, rpr.makeelement(qn("w:bCs"), {}))
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), ASC)
        rf.set(qn("w:eastAsia"), EA)

    def get_style_by_id(style_id):
        for st in doc.styles.element.findall(qn("w:style")):
            if st.get(qn("w:styleId")) == style_id:
                return st
        return None

    def get_or_add_rtf_heading() -> str:
        if get_style_by_id(STYLE_SID) is not None:
            return STYLE_SID
        st = doc.styles.add_style("RTF Heading", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        st.font.bold = True
        st.font.name = ASC
        ensure_style_bold(st.element)
        return STYLE_SID

    # 1) 收集带大纲级别的标题段落
    targets = []
    for p in doc.part.element.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:outlineLvl")) is not None:
            targets.append(p)
    if not targets:
        print("未找到带大纲级别的标题段落，跳过源标题加粗迁移。")
        return 0

    # 2) 逐段：仅迁移“run 带直接加粗”的段落
    n_paras = 0
    n_removed = 0
    rtf_heading_id = None
    for p in targets:
        ppr = p.find(qn("w:pPr"))
        has_direct_b = False
        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and (rpr.find(qn("w:b")) is not None
                                    or rpr.find(qn("w:bCs")) is not None):
                has_direct_b = True
                break
        if not has_direct_b:
            continue    # 本来就不是直接加粗，无需迁移

        ps = ppr.find(qn("w:pStyle"))
        style_id = ps.get(qn("w:val")) if ps is not None else None
        if style_id and style_id not in NORMAL_SIDS:
            style_el = get_style_by_id(style_id)
            if style_el is not None:
                ensure_style_bold(style_el)      # 保留原 pStyle，样式级加粗
            else:
                style_id = None
        if not style_id or style_id in NORMAL_SIDS:
            if rtf_heading_id is None:
                rtf_heading_id = get_or_add_rtf_heading()
            if ps is None:
                ps = ppr.makeelement(qn("w:pStyle"), {})
                ppr.insert(0, ps)
            ps.set(qn("w:val"), rtf_heading_id)

        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            for tag in ("w:b", "w:bCs"):
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
                    n_removed += 1
        n_paras += 1
    return n_paras


def _fix_toc_source_bold(docx_path: str | os.PathLike) -> int:
    """单步入口（兼容旧调用与单测）：load → 迁移加粗 → save → 自检。"""
    try:
        from docx import Document
        from docx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError(
            "源标题加粗迁移需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    doc = Document(docx_path)
    n_paras = _apply_toc_source_bold(doc)
    doc.save(docx_path)
    del doc                      # 先释放主 DOM，避免与自检 DOM 并存（峰值翻倍）
    gc.collect()

    # 自检：重开后确认已处理的段落不再有直接 b/bCs
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
    if n_paras and bad:
        raise RuntimeError(
            f"源标题加粗迁移自检失败：{len(ol)} 个源标题段落中仍有 {bad} 处直接 b/bCs。")
    del check
    gc.collect()
    return n_paras


def _self_check(docx_path: str | os.PathLike, stats: dict) -> None:
    """写盘后一次性复检，替代四个修正函数各自「重开一次」的自检。

    检查项与原各函数自检等价：
      1) 全文存在 ascii=Times New Roman + eastAsia=宋体 的 rPr；
      2) 本次处理过中文引号时，存在 hint=eastAsia 的引号 run；
      3) 本次设置过目录样式时，TOC 1~9 均为 蓝色 0000FF / 小四 sz=24 / 不加粗；
      4) 本次迁移过源标题加粗时，带大纲级别的段落不再有直接 b/bCs。
    """
    from docx import Document
    from docx.oxml.ns import qn

    ASC, EA = "Times New Roman", "宋体"
    QUOTES = "“”‘’"
    BLUE, SZ = "0000FF", "24"

    doc = Document(docx_path)
    roots = [doc.part.element]
    for rel in list(doc.part.rels.values()):
        reltype = getattr(rel, "reltype", "") or ""
        if "header" in reltype or "footer" in reltype:
            roots.append(rel.target_part.element)

    # 1) 字体
    font_ok = 0
    rpr = rf = None
    for root in roots:
        for rpr in root.iter(qn("w:rPr")):
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None and rf.get(qn("w:ascii")) == ASC \
                    and rf.get(qn("w:eastAsia")) == EA:
                font_ok += 1
    # 循环变量会一直绑着最后一个 lxml 元素（→ 整棵树），必须显式断开，
    # 否则下面的 gc.collect() 收不掉：实测不清时函数返回后工作集仍 499MB。
    rpr = rf = None
    if font_ok == 0:
        raise RuntimeError("字体修正自检失败：docx 中未发现 ascii=Times New Roman 的 run。")

    # 2) 中文引号
    if stats.get("chinese_quotes"):
        quote_ok = 0
        r = rpr = rf = None
        for root in roots:
            for r in root.iter(qn("w:r")):
                t = "".join(x.text or "" for x in r.findall(qn("w:t")))
                if t and any(q in t for q in QUOTES):
                    rpr = r.find(qn("w:rPr"))
                    rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
                    if rf is not None and rf.get(qn("w:hint")) == "eastAsia":
                        quote_ok += 1
        r = rpr = rf = None
        if quote_ok == 0:
            raise RuntimeError("中文引号处理自检失败：未发现 hint=eastAsia 的引号 run。")

    # 3) 目录样式
    bad_styles = []
    st = sz = color = b = None
    for st in doc.styles:
        sid = st.element.get(qn("w:styleId")) or ""
        if not re.fullmatch(r"TOC[1-9]", sid):
            continue
        rpr = st.element.find(qn("w:rPr"))
        color = rpr.find(qn("w:color")) if rpr is not None else None
        sz = rpr.find(qn("w:sz")) if rpr is not None else None
        b = rpr.find(qn("w:b")) if rpr is not None else None
        if (color is None or (color.get(qn("w:val")) or "").upper() != BLUE
                or sz is None or sz.get(qn("w:val")) != SZ
                or b is None or b.get(qn("w:val")) in (None, "1", "true")):
            bad_styles.append(sid)
    st = rpr = color = sz = b = None
    if stats.get("toc_styles") and bad_styles:
        raise RuntimeError(
            f"目录样式修正自检失败：{', '.join(bad_styles)} 未写入蓝色/小四/不加粗。")

    # 4) 源标题加粗
    if stats.get("headings_bolded"):
        bad_b = 0
        p = ppr = r = rpr = None
        for p in doc.part.element.iter(qn("w:p")):
            ppr = p.find(qn("w:pPr"))
            if ppr is None or ppr.find(qn("w:outlineLvl")) is None:
                continue
            for r in p.iter(qn("w:r")):
                rpr = r.find(qn("w:rPr"))
                if rpr is not None and (rpr.find(qn("w:b")) is not None
                                        or rpr.find(qn("w:bCs")) is not None):
                    bad_b += 1
        p = ppr = r = rpr = None
        if bad_b:
            raise RuntimeError(
                f"源标题加粗迁移自检失败：源标题段落中仍有 {bad_b} 处直接 b/bCs。")

    del doc, roots
    gc.collect()


def apply_all_fixes(docx_path: str | os.PathLike) -> dict:
    """一次 load → 四次修改 → 一次 save → 一次复检（合并写盘）。

    改动前：四个修正函数各自「load → 改 → save → 重开自检」，串行时每次都
    要重新解析整棵约 470MB 的 lxml 树、并各自序列化一遍 document.xml，
    实测峰值 1043MB，`xml_fix` 段耗时也被四次序列化拖长。

    改动后：解析一次、序列化一次，修改顺序与原实现完全一致
    （字体 → 中文引号 → 源标题加粗 → 目录样式），写盘后只重开一次做全部
    自检，峰值降到「单棵树 + 一次序列化」的水平。

    返回各步统计（键名沿用合并清单里的字段）：
      {"rfonts": int, "chinese_quotes": int, "headings_bolded": int,
       "toc_styles": int}
    """
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "docx 修正需要 python-docx，请先安装：pip install python-docx"
        ) from exc

    stats: dict = {}
    doc = Document(docx_path)
    stats["rfonts"] = _apply_fonts(doc)
    stats["chinese_quotes"] = _apply_quotes(doc)
    stats["headings_bolded"] = _apply_toc_source_bold(doc)
    stats["toc_styles"] = _apply_toc_styles(doc)
    doc.save(docx_path)          # 唯一一次序列化
    del doc                      # 先释放主 DOM，复检树才不致与它并存
    gc.collect()
    _self_check(docx_path, stats)
    # 自检函数的栈帧退出后再收一次：在自检内部调 gc 时，它的局部变量还活着，
    # 实测那一跳收不掉（返回后仍 498MB，补一次 gc 立刻 34MB）。
    gc.collect()
    return stats


