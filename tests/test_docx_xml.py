# -*- coding: utf-8 -*-
"""docx OOXML 层单测（M7 待补项）。

覆盖两类此前只靠真实产物间接验证、从未被测试触达的路径：
  1) 符号字体白名单：Symbol/Wingdings 等 rFonts 不能被覆盖成 Times New Roman；
  2) 源标题加粗迁移：带非 Normal pStyle（如 Heading 1）的标题必须保留原
     pStyle 并改走样式级加粗，而不是被覆盖成 RTFHeading。

另外补两条验收函数的边界：字体覆盖率不足时 `_verify_final_docx` 必须抛错；
中文引号必须被拆成 hint=eastAsia 的独立 run。

全部为合成 docx，不依赖真实样本、不启动 Word。
"""
from __future__ import annotations

import importlib.util
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rtfmerge import docx_xml                        # noqa: E402

HAS_DOCX = importlib.util.find_spec("docx") is not None

TOC_TEXT = 'TOC \\o "1-9" \\h \\z \\u'


def _make_doc(path: Path):
    from docx import Document
    return Document()


def _add_run(par, text, font=None):
    """加一个 run；font 不为 None 时写入显式 w:rFonts（ascii/hAnsi）。"""
    r = par.add_run(text)
    if font:
        from docx.oxml.ns import qn
        rpr = r._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = rpr.makeelement(qn("w:rFonts"), {})
            rpr.insert(0, rf)
        rf.set(qn("w:ascii"), font)
        rf.set(qn("w:hAnsi"), font)
    return r


@unittest.skipUnless(HAS_DOCX, "缺少 python-docx，跳过 docx_xml 用例")
class DocxXmlCase(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="rtfmerge_docxxml_"))

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _path(self, name="t.docx"):
        return self.dir / name

    def _rfonts_of(self, run_el):
        from docx.oxml.ns import qn
        rpr = run_el.find(qn("w:rPr"))
        if rpr is None:
            return None
        return rpr.find(qn("w:rFonts"))


class TestSymbolFonts(DocxXmlCase):
    def test_symbol_run_kept_after_fix(self):
        """Symbol 字体承载项目符号/勾选，修字体时必须原样保留。"""
        from docx.oxml.ns import qn

        p = self._path()
        doc = _make_doc(p)
        par = doc.add_paragraph()
        sym = _add_run(par, "\u2022", font="Symbol")
        plain = _add_run(par, "正文", font=None)
        doc.add_paragraph(TOC_TEXT)          # 验收要求文档里存在 TOC 域
        doc.save(p)

        docx_xml._fix_docx_fonts(p)

        from docx import Document
        doc2 = Document(p)
        runs = ran = doc2.paragraphs[0].runs
        rf_sym = self._rfonts_of(ran[0]._element)
        rf_plain = self._rfonts_of(ran[1]._element)
        self.assertEqual(rf_sym.get(qn("w:ascii")), "Symbol")
        self.assertEqual(rf_plain.get(qn("w:ascii")), "Times New Roman")
        self.assertEqual(rf_plain.get(qn("w:eastAsia")), "宋体")
        # 白名单 run 计入验收分母，因此在混合文档里要给出相称的阈值
        res = docx_xml._verify_final_docx(p, min_ratio=0.5)
        self.assertGreaterEqual(res["rfonts"], 2)

    def test_verification_raises_when_fonts_not_fixed(self):
        """未修字体的 docx（默认模板 Calibri）必须被验收拦下，而不是静默通过。"""
        p = self._path("raw.docx")
        doc = _make_doc(p)
        doc.add_paragraph(TOC_TEXT)
        doc.add_paragraph("正文")
        doc.save(p)
        with self.assertRaises(RuntimeError) as ctx:
            docx_xml._verify_final_docx(p, min_ratio=0.99)
        self.assertIn("字体覆盖率不足", str(ctx.exception))


class TestTocSourceBold(DocxXmlCase):
    def _add_heading_par(self, doc, text, style="Heading 1", bold=True):
        from docx.oxml.ns import qn
        par = doc.add_paragraph(style=style)
        ppr = par._p.get_or_add_pPr()
        ol = ppr.find(qn("w:outlineLvl"))       # 大纲级别是 TOC 的收集依据
        if ol is None:
            ol = ppr.makeelement(qn("w:outlineLvl"), {})
            ppr.append(ol)
        ol.set(qn("w:val"), "0")
        r = par.add_run(text)
        if bold:
            r.bold = True                        # 产生 run 级直接 w:b
        return par

    def test_heading_style_preserved_and_bold_moved_to_style(self):
        from docx import Document
        from docx.oxml.ns import qn

        p = self._path()
        doc = _make_doc(p)
        self._add_heading_par(doc, "第一章 概述")
        self._add_heading_par(doc, "正文也用了同一标题样式", bold=False)
        doc.add_paragraph("普通正文，没有大纲级别")
        doc.save(p)

        n = docx_xml._fix_toc_source_bold(p)
        self.assertEqual(n, 1)                   # 只有带直接加粗的那段被处理

        doc2 = Document(p)
        p1 = doc2.paragraphs[0]
        p2 = doc2.paragraphs[1]
        # 1) 原 pStyle 必须保留（不能被覆盖成 RTFHeading）
        ps = p1._p.find(qn("w:pPr")).find(qn("w:pStyle"))
        self.assertEqual(ps.get(qn("w:val")), "Heading1")
        # 2) 加粗迁移到样式层
        st = None
        for s in doc2.styles.element.findall(qn("w:style")):
            if s.get(qn("w:styleId")) == "Heading1":
                st = s
        self.assertIsNotNone(st)
        rpr = st.find(qn("w:rPr"))
        self.assertIsNotNone(rpr)
        self.assertIsNotNone(rpr.find(qn("w:b")))
        # 3) 被处理段落的 run 不再有直接 b
        for r in p1.runs:
            rf = r._element.find(qn("w:rPr"))
            if rf is not None:
                self.assertIsNone(rf.find(qn("w:b")))
        # 4) 同样式但原本不带直接加粗的段落不受影响
        for r in p2.runs:
            rf = r._element.find(qn("w:rPr"))
            if rf is not None:
                self.assertIsNone(rf.find(qn("w:b")))

    def test_no_outline_paragraphs_returns_zero(self):
        p = self._path("plain.docx")
        doc = _make_doc(p)
        doc.add_paragraph("只有正文，没有任何大纲级别")
        doc.save(p)
        self.assertEqual(docx_xml._fix_toc_source_bold(p), 0)


class TestChineseQuotes(DocxXmlCase):
    def test_quotes_are_split_with_eastasia_hint(self):
        from docx import Document
        from docx.oxml.ns import qn

        p = self._path()
        doc = _make_doc(p)
        doc.add_paragraph("他说“你好”，我说‘好’")
        doc.save(p)

        total = docx_xml._fix_chinese_quotes(p)
        self.assertEqual(total, 4)

        doc2 = Document(p)
        ok = 0
        for r in doc2.paragraphs[0].runs:
            t = "".join(x.text or "" for x in r._element.findall(qn("w:t")))
            if t and any(q in t for q in "“”‘’"):
                rf = self._rfonts_of(r._element)
                if rf is not None and rf.get(qn("w:hint")) == "eastAsia":
                    ok += 1
        self.assertEqual(ok, 4)

    def test_no_quotes_returns_zero(self):
        p = self._path("noq.docx")
        doc = _make_doc(p)
        doc.add_paragraph("没有中文引号的一段话")
        doc.save(p)
        self.assertEqual(docx_xml._fix_chinese_quotes(p), 0)


class TestApplyAllFixes(DocxXmlCase):
    """C 级重构后的新主路径（`apply_all_fixes` / `_self_check`）回归门。

    这批用例的由来：第八批把四个修正合并成「一次 load → 四改 → 一次 save →
    一次自检」，但当时只做过一次性手工 A/B，`tests/` 里没有一条用例触达新
    主路径。若不固化，后续任何一次重构都可能静默改坏主路径而 24 项仍全绿。
    """

    def _rich_doc(self):
        """带目录域、大纲加粗标题、中文引号的合成 docx（四项检查都会真的跑）。"""
        p = self._path("rich.docx")
        doc = _make_doc(p)
        doc.add_paragraph(TOC_TEXT)
        par = doc.add_paragraph(style="Heading 1")
        from docx.oxml.ns import qn
        ppr = par._p.get_or_add_pPr()
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is None:
            ol = ppr.makeelement(qn("w:outlineLvl"), {})
            ppr.append(ol)
        ol.set(qn("w:val"), "0")
        r = par.add_run("第一章 概述")
        r.bold = True
        doc.add_paragraph("他说“你好”，我说‘好’")
        doc.save(p)
        return p

    def test_stats_complete_and_self_check_passes(self):
        """四个统计键必须齐全，且四项自检不得误抛。"""
        p = self._rich_doc()
        stats = docx_xml.apply_all_fixes(p)

        self.assertEqual(
            set(stats), {"rfonts", "chinese_quotes", "headings_bolded", "toc_styles"})
        self.assertGreater(stats["rfonts"], 0)
        self.assertEqual(stats["chinese_quotes"], 4)
        self.assertEqual(stats["headings_bolded"], 1)
        self.assertIsInstance(stats["toc_styles"], int)
        # 自检已在函数内跑过；此处再从产物侧确认两项硬指标
        from docx import Document
        from docx.oxml.ns import qn

        doc2 = Document(p)
        blue = 0
        for s in doc2.styles.element.findall(qn("w:style")):
            sid = s.get(qn("w:styleId")) or ""
            if sid.startswith("TOC") and s.find(qn("w:rPr")) is not None:
                col = s.find(qn("w:rPr")).find(qn("w:color"))
                if col is not None and (col.get(qn("w:val")) or "").upper() == "0000FF":
                    blue += 1
        self.assertEqual(blue, stats["toc_styles"])

    def test_no_quotes_no_outline_does_not_raise(self):
        """无中文引号 / 无大纲标题时，自检的两个条件分支必须跳过而非误抛。"""
        p = self._path("plain.docx")
        doc = _make_doc(p)
        doc.add_paragraph(TOC_TEXT)
        doc.add_paragraph("只有正文，没有中文引号也没有大纲级别")
        doc.save(p)

        stats = docx_xml.apply_all_fixes(p)
        self.assertEqual(stats["chinese_quotes"], 0)
        self.assertEqual(stats["headings_bolded"], 0)
        self.assertGreater(stats["rfonts"], 0)

    def test_equals_stepwise_implementation(self):
        """同一输入分别走「四个单步入口」与 `apply_all_fixes`，产物必须逐成员相同。

        把第八批那次一次性手工 A/B 变成可重复的回归门：只排除含时间戳的
        `docProps/`，其余 zip 成员按 SHA256 全等。
        """
        import hashlib
        import zipfile

        p_step = self._rich_doc()
        p_all = self.dir / "rich_all.docx"
        shutil.copyfile(p_step, p_all)
        p_step.rename(self.dir / "rich_step.docx")
        p_step = self.dir / "rich_step.docx"

        docx_xml._fix_docx_fonts(p_step)
        docx_xml._fix_chinese_quotes(p_step)
        docx_xml._fix_toc_source_bold(p_step)
        docx_xml._fix_toc_styles(p_step)
        docx_xml.apply_all_fixes(p_all)

        def parts(path):
            out = {}
            with zipfile.ZipFile(path) as z:
                for n in z.namelist():
                    if n.startswith("docProps/"):
                        continue
                    out[n] = hashlib.sha256(z.read(n)).hexdigest()
            return out

        a, b = parts(p_step), parts(p_all)
        self.assertEqual(sorted(a), sorted(b))
        diff = [n for n in a if a[n] != b[n]]
        self.assertEqual(diff, [], "单步与合并两条路径产出的 XML 成员不一致：%s" % diff)


if __name__ == "__main__":
    unittest.main(verbosity=2)
