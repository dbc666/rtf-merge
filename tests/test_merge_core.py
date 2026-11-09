# -*- coding: utf-8 -*-
"""合并核心回归测试（unittest）。

除 order 相关用例外均为纯标准库，无需第三方依赖；
order 用例需要 openpyxl，缺失时自动跳过。
真实样本目录默认取 F:\\rtf_merge\\output_rtf_空腹_TF，可用环境变量
RTF_MERGE_TEST_INPUT 覆盖；目录不存在时相关用例跳过。

运行（建议指定带依赖的解释器）：
    python -m unittest discover -s F:\\rtf_merge\\tests -v
"""
from __future__ import annotations

import importlib.util
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rtfmerge.merge_core import merge_rtf          # noqa: E402
from rtfmerge.rtf_lint import (                    # noqa: E402
    check_structure, extract_resource_tables, resource_fingerprints,
)

HAS_OPENPYXL = importlib.util.find_spec("openpyxl") is not None
REAL_INPUT = Path(os.environ.get("RTF_MERGE_TEST_INPUT",
                                 r"F:\rtf_merge\output_rtf_空腹_TF"))

# 合成样例：末行独立为 "}"（SAS ODS 的典型结构），用于验证 drop_sect_gap
SYNTH_RTF = (
    b"{\\rtf1\\ansi\\ansicpg936\\uc1\r\n"
    b"{\\fonttbl{\\f0\\fnil Arial;}}\r\n"
    b"{\\colortbl;\\red0\\green0\\blue0;}\r\n"
    b"\\pard hello world\\par\r\n"
    b"}\r\n"
)


def _brace_balanced(text: str) -> bool:
    return text.count("{") == text.count("}")


class TempDirCase(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="rtfmerge_test_"))
        self.samples = sorted(REAL_INPUT.glob("*.rtf"))[:3] if REAL_INPUT.is_dir() else []

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def require_samples(self, n=1):
        if len(self.samples) < n:
            self.skipTest(f"缺少真实输入样本（{REAL_INPUT}），跳过")

    def copy_samples(self, n=3, dest=None):
        self.require_samples(n)
        dest = dest or self.dir
        copied = []
        for p in self.samples[:n]:
            q = dest / p.name
            shutil.copy2(p, q)
            copied.append(q)
        return copied


class TestCollectAndExclude(TempDirCase):
    def test_skip_lock_hidden_and_outfile(self):
        self.copy_samples(2)
        (self.dir / "~$l_merged.rtf").write_bytes(b"\xd0\xcf\x11\xe0fake")
        (self.dir / ".hidden.rtf").write_text("{\\rtf1 x}", encoding="utf-8")
        out = self.dir / "all_merged.rtf"
        res = merge_rtf(self.dir, out, order=0, pageyn=1)
        names = [Path(f).name for f in res["files"]]
        self.assertNotIn("~$l_merged.rtf", names)
        self.assertNotIn(".hidden.rtf", names)
        self.assertNotIn("all_merged.rtf", names)
        self.assertEqual(len(names), 2)

    def test_second_run_does_not_grow(self):
        self.copy_samples(2)
        out = self.dir / "all_merged.rtf"
        merge_rtf(self.dir, out, order=0, pageyn=1)
        res2 = merge_rtf(self.dir, out, order=0, pageyn=1)
        self.assertEqual(len(res2["files"]), 2)


class TestBoundaries(TempDirCase):
    def test_single_file_closed(self):
        self.copy_samples(1)
        out = self.dir / "out.rtf"
        res = merge_rtf(self.dir, out, order=0, pageyn=1)
        self.assertEqual(res["status"], "merged")
        txt = out.read_text(encoding="utf-8")
        self.assertTrue(_brace_balanced(txt))
        self.assertNotIn("\\pgnrestart", txt)

    def test_empty_input_refused(self):
        for i in range(2):
            (self.dir / f"e{i}.rtf").write_text("", encoding="utf-8")
        out = self.dir / "out.rtf"
        res = merge_rtf(self.dir, out, order=0, pageyn=1)
        self.assertEqual(res["status"], "empty_output")
        self.assertFalse(out.exists())

    def test_last_file_empty_keeps_structure(self):
        self.copy_samples(1)
        (self.dir / "z.rtf").write_text("", encoding="utf-8")
        out = self.dir / "out.rtf"
        res = merge_rtf(self.dir, out, order=0, pageyn=1)
        self.assertEqual(res["status"], "merged")
        txt = out.read_text(encoding="utf-8")
        self.assertTrue(_brace_balanced(txt))

    @unittest.skipUnless(HAS_OPENPYXL, "未安装 openpyxl，跳过 order 相关用例")
    def test_all_order_empty_refused(self):
        import openpyxl
        self.copy_samples(2)
        out = self.dir / "out.rtf"
        merge_rtf(self.dir, out, order=1, pageyn=1)      # 生成模板并退出
        wb = openpyxl.load_workbook(self.dir / "file_order.xlsx")
        ws = wb.active
        for row in ws.iter_rows(min_row=2):
            row[1].value = None
        wb.save(self.dir / "file_order.xlsx")
        res = merge_rtf(self.dir, out, order=1, pageyn=1)
        self.assertEqual(res["status"], "no_ordered_files")
        self.assertFalse(out.exists())

    def test_manifest_written(self):
        self.copy_samples(2)
        out = self.dir / "out.rtf"
        res = merge_rtf(self.dir, out, order=0, pageyn=1)
        self.assertIn("manifest", res)
        self.assertTrue(Path(res["manifest"]).exists())

    @unittest.skipUnless(HAS_OPENPYXL, "未安装 openpyxl，跳过 order 相关用例")
    def test_partial_order_refused(self):
        """排序表只填一部分时必须拒绝合并，不能静默漏合并。"""
        import openpyxl
        self.copy_samples(3)
        out = self.dir / "out.rtf"
        merge_rtf(self.dir, out, order=1, pageyn=1)      # 生成模板并退出
        wb = openpyxl.load_workbook(self.dir / "file_order.xlsx")
        ws = wb.active
        ws.cell(row=2, column=2).value = 1               # 只给第 1 行填 order
        wb.save(self.dir / "file_order.xlsx")

        res = merge_rtf(self.dir, out, order=1, pageyn=1)
        self.assertEqual(res["status"], "order_incomplete")
        self.assertFalse(out.exists())
        self.assertTrue(res.get("missing"))

        # 显式放行时只告警、继续合并
        res2 = merge_rtf(self.dir, out, order=1, pageyn=1, allow_partial_order=True)
        self.assertEqual(res2["status"], "merged")
        self.assertTrue(any("allow-partial-order" in w for w in res2["warnings"]))


class TestDropSectGap(TempDirCase):
    def test_drop_removes_one_line_per_file(self):
        # 无真实样本时直接 skip：否则 self.samples 为空，后面读不到 out.rtf
        # 会抛 FileNotFoundError，让回归门在别的机器上变成 ERROR 而不是 skip
        # 至少要 2 个样本：单文件时 maxn==1，merge_core 的首文件分支根本不删
        # 末行 "}"（差 0），而 expected 仍按"末行为 } 的文件数"算 1 → 必然失败。
        self.require_samples(2)
        # 两版分别输出到独立目录，避免前一版产物被当作输入（同目录会自包含）
        d1, d2 = self.dir / "gap", self.dir / "nogap"
        for d in (d1, d2):
            d.mkdir()
        files = []
        for p in self.samples:
            files.append(shutil.copy2(p, d1 / p.name))
            shutil.copy2(p, d2 / p.name)
        a = d1 / "out.rtf"
        b = d2 / "out.rtf"
        merge_rtf(d1, a, order=0, pageyn=1, drop_sect_gap=False)
        merge_rtf(d2, b, order=0, pageyn=1, drop_sect_gap=True)
        la = a.read_text(encoding="utf-8").splitlines()
        lb = b.read_text(encoding="utf-8").splitlines()
        # 实现里每个「末行独立为 }」的文件都会丢掉该行（含最后一个，
        # 最后由 want[-1] + "}" 补回闭合），所以按全量 files 计数。
        expected = sum(
            1 for p in files
            if p.read_text(encoding="utf-8", errors="replace").rstrip("\r\n").splitlines()[-1].strip() == "}"
        )
        self.assertEqual(len(la) - len(lb), expected)
        # 分节符数量不受影响（用户要求：不能删分节符）
        self.assertEqual(a.read_text(encoding="utf-8").count("\\sect"),
                         b.read_text(encoding="utf-8").count("\\sect"))

    def test_drop_on_synthetic_bare_braces(self):
        """合成 3 个末行均为独立 } 的文件：drop 应比 keep 恰好少 3 行。

        真实样本里常混有末行为 \\pard} 的文件，会把断言的偏差掩盖掉，
        因此这里用纯合成数据做黄金基线（不依赖真实目录）。
        """
        d1, d2 = self.dir / "g1", self.dir / "g2"
        for d in (d1, d2):
            d.mkdir()
        for i in range(3):
            (d1 / f"s{i}.rtf").write_bytes(SYNTH_RTF)
            (d2 / f"s{i}.rtf").write_bytes(SYNTH_RTF)
        merge_rtf(d1, d1 / "out.rtf", order=0, pageyn=1, drop_sect_gap=False)
        merge_rtf(d2, d2 / "out.rtf", order=0, pageyn=1, drop_sect_gap=True)
        la = (d1 / "out.rtf").read_bytes().split(b"\r\n")
        lb = (d2 / "out.rtf").read_bytes().split(b"\r\n")
        self.assertEqual(len(la) - len(lb), 3)
        self.assertEqual((d1 / "out.rtf").read_text(encoding="utf-8").count("\\sect"),
                         (d2 / "out.rtf").read_text(encoding="utf-8").count("\\sect"))


class TestStructureGate(TempDirCase):
    """结构校验不通过时，状态必须是 structure_invalid（三条入口共用同一判据）。

    回归点：此前 merge_rtf 只把 issues 记进 warnings、status 仍为 merged，
    CLI 另有一道 issues 判断而 GUI 只看 status → GUI 会继续跑 Word 并弹"完成"。
    """

    BAD = (
        b"{\\rtf1\\ansi\\ansicpg936\\uc1\r\n"
        b"{\\fonttbl{\\f0\\fnil Arial;}}\r\n"
        b"{\\colortbl;\\red0\\green0\\blue0;}\r\n"
        b"\\pard hello\\par\r\n"        # 故意不写末行 }：4 开 3 闭
    )

    def test_bad_structure_returns_structure_invalid(self):
        d = self.dir / "bad"
        d.mkdir()
        (d / "a.rtf").write_bytes(self.BAD)
        out = self.dir / "out.rtf"
        res = merge_rtf(d, out, order=0, pageyn=1)
        self.assertEqual(res["status"], "structure_invalid")
        self.assertTrue(res["rtf_validation"]["issues"])
        # 文件与清单仍要写出（便于排查），只是不进入后处理
        self.assertTrue(out.exists())
        self.assertTrue(res.get("manifest") and Path(res["manifest"]).exists())

    def test_good_structure_stays_merged(self):
        d = self.dir / "good"
        d.mkdir()
        (d / "a.rtf").write_bytes(SYNTH_RTF)
        res = merge_rtf(d, self.dir / "ok.rtf", order=0, pageyn=1)
        self.assertEqual(res["status"], "merged")


class TestWarningDetail(TempDirCase):
    """告警的双形态：warnings（旧字符串数组）+ warnings_detail（kind/level/text）。

    回归点：改造前「输出落在输入目录内」这类使用提示靠中文前缀过滤，改一个字
    就会把提示当成产物缺陷、把 all_ok 拉黑。改造后判定只看 level。
    """

    def _merged(self, out_in_input_dir: bool):
        indir = self.dir / "in"
        indir.mkdir()
        for name in ("a.rtf", "b.rtf"):
            (indir / name).write_bytes(SYNTH_RTF)
        out = (indir / "merged.rtf") if out_in_input_dir else (self.dir / "merged.rtf")
        return merge_rtf(indir, out, order=0, pageyn=1)

    def test_input_dir_warning_is_info_and_mirrored(self):
        res = self._merged(out_in_input_dir=True)
        self.assertEqual(res["status"], "merged")
        detail = res["warnings_detail"]
        self.assertEqual(len(detail), 1)
        self.assertEqual(detail[0]["kind"], "input_dir_overlap")
        self.assertEqual(detail[0]["level"], "info")
        # 两个字段必须同源：文本逐条一致，旧读取方拿到的仍是纯字符串
        self.assertEqual(res["warnings"], [d["text"] for d in detail])
        import json
        m = json.loads(Path(res["manifest"]).read_text(encoding="utf-8"))
        self.assertEqual(m["warnings"], res["warnings"])
        self.assertEqual(m["warnings_detail"], detail)

    def test_info_warning_does_not_block_all_ok(self):
        """使用提示（level=info）不得计入 warning_count，也不得让 all_ok 变 False。"""
        from rtfmerge.merge_core import update_manifest_postprocess

        res = self._merged(out_in_input_dir=True)
        post = {
            "docx_sha256": "x", "docx_size": 1, "pages": 1,
            "remaining_empty_pages": 0,
            "fonts": {"rfonts": 100, "tnr_ratio": 1.0, "song_ratio": 1.0, "has_toc": True},
            "step1": {"merge_skipped": 0},
        }
        m = update_manifest_postprocess(res["manifest"], post)
        self.assertEqual(m["validation"]["warning_count"], 0)
        self.assertTrue(m["validation"]["all_ok"])

    def test_no_warning_outside_input_dir(self):
        res = self._merged(out_in_input_dir=False)
        self.assertEqual(res["warnings"], [])
        self.assertEqual(res["warnings_detail"], [])


class TestLint(unittest.TestCase):
    def test_structure_ok(self):
        lines = ["{\\rtf1\\ansi", "{\\fonttbl{\\f0 Arial;}}", "\\pard text\\par", "}"]
        self.assertEqual(check_structure(lines, "t.rtf"), [])

    def test_structure_bad(self):
        # 首行不是 { 、花括号不配平、\rtf1 出现两次
        lines = ["\\rtf1 {\\rtf1", "text", "}"]
        issues = check_structure(lines, "t.rtf")
        self.assertTrue(any("首行" in i for i in issues))
        self.assertTrue(any("\\rtf1" in i for i in issues))

    @unittest.skipUnless(REAL_INPUT.is_dir(), "缺少真实输入目录")
    def test_resource_tables_extracted(self):
        files = sorted(REAL_INPUT.glob("*.rtf"))[:5]
        tabs = extract_resource_tables(files[0])
        self.assertTrue(tabs["fonttbl"].startswith("{"))
        self.assertIn("\\fonttbl", tabs["fonttbl"])
        groups, warns = resource_fingerprints(files)
        self.assertEqual(len(groups), 1)      # 同批模板应一致
        self.assertEqual(warns, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
