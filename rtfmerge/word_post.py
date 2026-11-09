# -*- coding: utf-8 -*-
"""Word 后处理编排层：PowerShell / Word COM 调用，脚本外置为 ps/*.ps1。

零 pywin32：全部通过 PowerShell COM 驱动 Word。
三段脚本分别负责：
  step1_convert.ps1    RTF → docx，关键词单元格合并、正文表格 AutoFit、插入/刷新目录
  step2_find_empty.ps1 空白页【只读】检测（不写盘、不弹窗，已不能交互删页）
  step3_recheck.ps1    最终复检（独立实例，只检测不写盘）

写盘顺序（关键）：ps1 写 docx（Word 唯一一次写盘）→ python-docx 重写（第 1.5 步）
→ ps2 只读检测 → docx_xml 修正字体/引号/加粗/目录样式（最后一次写盘）
→ ps3 只读检测 → 最终验收。

ps2 自 v1.1.2 起已降级为纯只读：既不再 SaveAs，也不再弹窗询问，
因此不存在“Word 重存把显式字体冲成主题字体”的顺序耦合风险
（历史上该风险由 ps2 置于 docx_xml 之前来规避，现已从根上消除）。

清理残留 Word 进程时只终止【本次任务新增、且无窗口】的进程，绝不碰用户
自己打开的 Word（按 MainWindowTitle 批量强杀会丢失未保存文档）。
"""
from __future__ import annotations

import gc
import hashlib
import os
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from . import docx_xml
from . import meminfo

PS_SUBDIR = "ps"

# 取消信号（GUI 的“停止”按钮 / 关窗时设置；CLI 从不设置，故行为不变）
CANCEL_EVENT = threading.Event()


class TaskCancelled(RuntimeError):
    """任务被用户取消。"""


def resource_path(rel: str) -> str:
    """定位打包内/源码目录中的资源文件（ps 脚本、图标）。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        # PyInstaller 打包：add-data 到 rtfmerge/ps 下
        return os.path.join(base, "rtfmerge", rel)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel)


def _find_powershell() -> str:
    """定位可用的 PowerShell（优先 pwsh 7，其次 Windows PowerShell）。"""
    for exe in ("pwsh", "powershell"):
        p = shutil.which(exe)
        if p:
            return p
    raise RuntimeError("未找到 PowerShell（pwsh/powershell），无法执行 Word 字体修正。")


def _list_word_pids() -> set[int] | None:
    """列出当前 WINWORD 进程 PID；失败返回 None（表示无法判断）。"""
    if os.name != "nt":
        return set()
    try:
        proc = subprocess.run(
            [_find_powershell(), "-NoProfile", "-Command",
             # 末尾 exit 0：Windows PowerShell 5.1 在 Get-Process 无匹配时
             # 退出码为 1（即使 -ErrorAction SilentlyContinue），不加会把
             # "没有 Word 进程"误判为"枚举失败"。
             "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | "
             "ForEach-Object { $_.Id }; exit 0"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if proc.returncode != 0:
            return None
        return {int(t) for t in (proc.stdout or "").split() if t.strip().isdigit()}
    except Exception:
        return None


# 进程内的 Word 基线：只拍一次，之后所有任务共用同一个起点。
# 因果要看清：基线里的进程永远不杀，基线之外的才可能清。因而
#   · 同一进程内，本工具自己起的 Word 始终在基线之外 → 后续任务仍能清掉
#     （上一任务若被强杀留下僵尸，它比基线晚，仍可清）；
#   · 若每次任务都重拍，上一任务的残留会被并进基线 → 反而永远清不掉，
#     连 pre-flight 也清不掉（同一个判据），只能报错让用户手动结束。
# 代价：本进程存活期间用户自己新开的无窗口 Word 可能被当成"新增"——
# 但用户打开的文档窗口必有标题，判据第 2 条已把它排除。
# 跨进程（进程被强杀）遗留的僵尸无法定位，属已知限制，由 pre-flight 文案引导。
_WORD_PID_BASELINE: set[int] | None = None
_BASELINE_TAKEN = False

# 本次任务里由本工具创建的 Word 实例 PID。三段 ps1 各自在创建实例前后枚举
# WINWORD 差集并输出 WORD_PID=，这里累加起来；任务结束/取消/异常时按 PID
# 精确回收。它不依赖窗口标题、不依赖进程基线，也不会波及用户自己的 Word
# （ps1 复用已有实例时差集为空，正好不记录）。
_TASK_WORD_PIDS: set[int] = set()


def _snapshot_word_baseline() -> None:
    """记录进程启动以来第一次看到的 Word 进程（这些进程绝不能被清理）。

    只在首次调用时拍，之后复用（见上方常量注释里的因果说明）。
    """
    global _WORD_PID_BASELINE, _BASELINE_TAKEN
    if _BASELINE_TAKEN:
        return
    pids = _list_word_pids()
    _WORD_PID_BASELINE = pids
    _BASELINE_TAKEN = True
    if pids is None:
        print("提示：无法枚举现有 Word 进程，本次不做残留进程清理（不会影响结果）。")


def _kill_headless_word() -> None:
    """清理残留的 Word 实例。

    第一步按「本次任务 PID」精确回收（ps1 上报的、本工具实打实创建的那些），
    第二步再用基线差集兜底（覆盖 ps1 没报 PID 的老情况）。
    精确回收不依赖窗口标题、不依赖基线，也不会波及用户自己打开的 Word。
    """
    _kill_task_word_pids()
    if os.name != "nt":
        return
    baseline = _WORD_PID_BASELINE
    if baseline is None:
        return           # 没有可信基线：安全优先，不做清理
    try:
        base_lit = ",".join(str(p) for p in sorted(baseline)) if baseline else ""
        script = (
            "$base = @(" + base_lit + "); "
            "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | "
            "Where-Object { ($base -notcontains $_.Id) -and "
            "($_.MainWindowTitle -eq '') } | Stop-Process -Force"
        )
        subprocess.run(
            [_find_powershell(), "-NoProfile", "-Command", script],
            capture_output=True, timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        pass


def _collect_word_pids(stdout: str | None) -> None:
    """从 ps1 输出里取 WORD_PID=1,2（创建实例前后的 WINWORD 差集）。"""
    v = _grab_str(stdout, "WORD_PID=")
    if not v:
        return
    for part in v.split(","):
        part = part.strip()
        if part.isdigit():
            _TASK_WORD_PIDS.add(int(part))


def _kill_task_word_pids() -> None:
    """按 PID 精确回收本次任务创建的 Word 实例（多一道进程名校验防 PID 复用）。"""
    if os.name != "nt" or not _TASK_WORD_PIDS:
        return
    try:
        ids = ",".join(str(p) for p in sorted(_TASK_WORD_PIDS))
        script = (
            "$ids = @(" + ids + "); "
            "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | "
            "Where-Object { $ids -contains $_.Id } | Stop-Process -Force"
        )
        subprocess.run(
            [_find_powershell(), "-NoProfile", "-Command", script],
            capture_output=True, timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        pass
    finally:
        _TASK_WORD_PIDS.clear()


def _docx_locked(path: str | os.PathLike) -> bool:
    """docx 是否被独占（Word 残留实例会持有句柄）。"""
    try:
        fh = open(path, "rb+")
        fh.close()
        return False
    except OSError:
        return True


def _run_ps(script: str, args: list[str], timeout: int) -> subprocess.CompletedProcess:
    """执行 ps/*.ps1（直接 -File 指向资源文件，不落地临时副本）。

    与 subprocess.run 的区别：可被 CANCEL_EVENT 中断（GUI 停止/关窗时杀掉
    当前脚本子进程），并继续遵守 timeout。"""
    path = resource_path(os.path.join(PS_SUBDIR, script))
    if not os.path.isfile(path):
        raise RuntimeError(f"缺少后处理脚本：{path}")
    shell = _find_powershell()
    cmd = [shell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path] + args
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace", creationflags=flags)
    deadline = time.time() + timeout
    while True:
        if CANCEL_EVENT.is_set():
            _kill(proc)
            raise TaskCancelled(f"任务已被用户取消（{script}）")
        try:
            out, err = proc.communicate(timeout=1)
            return subprocess.CompletedProcess(cmd, proc.returncode, out, err)
        except subprocess.TimeoutExpired:
            if time.time() > deadline:
                _kill(proc)
                raise subprocess.TimeoutExpired(cmd, timeout)


def _kill(proc: subprocess.Popen) -> None:
    """尽力杀掉子进程及其子进程树（PowerShell 会再拉起 Word）。"""
    try:
        proc.kill()
    except Exception:
        pass
    try:
        proc.communicate(timeout=5)
    except Exception:
        pass


def check_cancel() -> None:
    """在长流程的检查点调用：已请求取消则抛出 TaskCancelled。"""
    if CANCEL_EVENT.is_set():
        raise TaskCancelled("任务已被用户取消")


def _grab(stdout: str | None, prefix: str) -> int | None:
    for line in (stdout or "").splitlines():
        if line.startswith(prefix):
            try:
                return int(line.split("=", 1)[1])
            except ValueError:
                return None
    return None


def _grab_str(stdout: str | None, prefix: str) -> str | None:
    """取 key=value 形式的字符串值（用于 Word 版本号等）。"""
    for line in (stdout or "").splitlines():
        if line.startswith(prefix):
            return line.split("=", 1)[1].strip()
    return None


def _sha256(path: str | os.PathLike) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class PostResult(dict):
    """后处理结果：既可当字典取各项统计，也保留旧的“当路径字符串用”的用法。"""

    def __str__(self) -> str:
        return str(self.get("docx", ""))


def fix_fonts_via_word(outfile: str | os.PathLike) -> PostResult:
    """合并后用 Word 修正字体并另存 .docx（对应宏尾部 Font.vbs 步骤）。

    全流程零 pywin32，仅通过 PowerShell COM 驱动 Word：
      1) ps1：RTF → docx（关键词单元格合并 + 正文表格 AutoFit + 插入目录）；
      2) python-docx 重写 docx（Word 直接 SaveAs2 的 docx 新实例打开会挂起）；
      3) ps2：空白页检测（只读，不写盘、不弹窗）；
      4) docx_xml：字体 / 中文引号 / 标题加粗迁移 / 目录样式（唯一写盘者）；
      5) ps3：独立实例最终复检（只读）；
      6) 最终验收：字体覆盖率 + 目录域。

    返回 PostResult（dict）：含 docx 路径、页数、字体覆盖率、各环节计数，
    可直接交给 merge_core.update_manifest_postprocess 写入合并清单；
    str(返回值) 仍是 docx 路径，兼容旧的“当字符串用”的调用方式。
    """
    out = Path(outfile)
    if not out.is_file():
        raise RuntimeError(f"待修正字体的文件不存在: {out}")
    docx_out = out.with_suffix(".docx")
    # 先给现存 Word 进程拍快照：后续清理只针对本次新增的无窗口实例
    _snapshot_word_baseline()
    _TASK_WORD_PIDS.clear()      # 本次任务从零开始收集 ps1 上报的 Word PID
    t0 = time.time()
    t_lap = t0
    durations: dict = {}

    def _lap(name: str) -> None:
        """记录上一段耗时（写入 durations，单位秒，保留 1 位）。"""
        nonlocal t_lap
        now = time.time()
        durations[name] = round(now - t_lap, 1)
        t_lap = now

    info: dict = {"outfile": str(out), "docx": str(docx_out)}
    # 后处理起点的内存读数（清单里要给"这次占多少"，不能只在 GUI 日志里出现）
    mem_start, _ = meminfo.current()

    # 第 1 趟之前先确认目标 docx 没被占用。上一轮跑崩或别的工具留下的无窗口
    # Word 会一直持有旧 docx 的句柄，导致 ps1 的 SaveAs2 报“文件已在别处打开”，
    # 报错信息完全指不到真正原因。这里按基线差集清一次，清不掉再明确报错。
    if docx_out.exists() and _docx_locked(docx_out):
        _kill_headless_word()
        time.sleep(2)
        if _docx_locked(docx_out):
            raise RuntimeError(
                f"输出 docx 被占用，无法覆盖：{docx_out}\n"
                f"可能仍有 Word 实例打开着它（用户自己打开的 Word 不会被自动关闭）。\n"
                f"如果看不到任何 Word 窗口，多半是上次任务残留的无窗口实例："
                f"请在任务管理器里结束 WINWORD.EXE 后重试。")

    # ---- 第 1 趟（Word 会话 1）：RTF -> docx + 单元格合并 + AutoFit + 目录 ----
    check_cancel()
    try:
        proc = _run_ps("step1_convert.ps1",
                       [str(out.resolve()), str(docx_out.resolve())], timeout=1200)
    except subprocess.TimeoutExpired as exc:
        _kill_headless_word()
        raise RuntimeError("Word 转换 docx 超时（>1200s），请确认 Word 未被占用。") from exc
    if proc.returncode != 0 or not docx_out.is_file():
        raise RuntimeError(
            "Word 转换 docx 失败：\n" + (proc.stderr or proc.stdout or "未知错误"))

    _collect_word_pids(proc.stdout)
    _lap("step1_convert")

    n_tables = _grab(proc.stdout, "AUTOFIT_TABLES=")
    n_merged = _grab(proc.stdout, "MERGE_CELLS=")
    n_skipped = _grab(proc.stdout, "MERGE_SKIPPED=")
    n_toc = _grab(proc.stdout, "TOC_INSERTED=")
    info["step1"] = {"merged_cells": n_merged, "autofit_tables": n_tables,
                     "merge_skipped": n_skipped, "toc_inserted": n_toc}
    if n_merged is not None:
        print(f"已将 {n_merged} 个关键词单元格与右侧空单元格合并（页眉页脚表格未处理）")
    if n_skipped:
        print(f"注意：{n_skipped} 个单元格因向右无单元格（多为垂直合并的表头）跳过未合并")
    if n_tables is not None:
        print(f"已按内容并铺满窗口自动调整 {n_tables} 个正文表格（页眉页脚表格未处理）")
    if n_toc:
        print("已在“目录”下方插入目录（点状引导符、宋体+Times New Roman、小四、蓝色、"
              "按大纲级别共抓取正文标题）")

    # ---- 第 1.5 步：python-docx 重写（重新打包），修复 Word 打开挂起 ----
    try:
        from docx import Document as _RewriteDoc
        _rd = _RewriteDoc(docx_out)
        _rd.save(docx_out)
    except Exception as _rewrite_err:
        print(f"提示：docx 重写失败（{_rewrite_err}），可能影响后续 Word 打开。")
    finally:
        # 这棵树约 470MB（15MB docx 展开约 30 倍）。若留到函数返回，它会成为
        # 后面 4 个 XML 修正步骤的常驻底座，把峰值从 ~1.0GB 抬到 ~1.5GB；
        # 而且 lxml 元素树有循环引用，光靠引用计数释放不掉，不 gc 的话任务
        # 结束后工作集停在 1GB 以上（实测 del 后仍 966MB，gc 一次立刻回到 37MB）。
        # 故这里必须 del + gc.collect()，且重写失败时也要走这一步。
        try:
            del _rd
        except UnboundLocalError:
            pass
        gc.collect()
    _lap("docx_rewrite")

    # ---- 第 2 趟（新 Word 会话）：空白页检测（只检测，不写盘）----
    check_cancel()
    if _docx_locked(docx_out):
        _kill_headless_word()
        time.sleep(3)
    if _docx_locked(docx_out):
        raise RuntimeError(
            "输出 docx 被占用（可能是残留 Word 实例），已尝试清理仍不可写，"
            "请关闭相关 Word 后重试。")
    try:
        proc2 = _run_ps("step2_find_empty.ps1", [str(docx_out.resolve())], timeout=2700)
    except subprocess.TimeoutExpired as exc:
        _kill_headless_word()
        tail = ""
        if isinstance(exc.stdout, str):
            tail = exc.stdout[-2000:]
        elif exc.stdout:
            tail = (exc.stdout or b"").decode("utf-8", errors="replace")[-2000:]
        raise RuntimeError(
            f"空白页检测超时（>2700s），请确认 Word 未被占用。\n最近输出：\n{tail}") from exc
    if proc2.returncode != 0 or not docx_out.is_file():
        _kill_headless_word()
        raise RuntimeError(
            "空白页检测失败：\n" + (proc2.stderr or proc2.stdout or "未知错误"))
    _collect_word_pids(proc2.stdout)
    _lap("step2_detect")
    n_empty2 = _grab(proc2.stdout, "EMPTY_PAGES_COUNT=")
    pages2 = _grab(proc2.stdout, "DOC_PAGES=")
    word_ver = _grab_str(proc2.stdout, "WORD_VERSION=")
    info["step2_detect"] = {"pages": pages2, "empty_pages": n_empty2,
                            "stage": "pre_xml_fix"}
    if n_empty2:
        print(f"提示：转换阶段检测到 {n_empty2} 个空白页"
              f"（本环节只读、不修改，详见最终复检）")
    else:
        print("未检测到空白页（文档中无仅含分页符/段落符/空格的页面）")

    # ---- 第 3~5 步：python-docx 修正（确定性、可验证）----
    # 四个修正合并成「一次 load → 四次改 → 一次 save → 一次自检」：原来每步
    # 都要重新解析整棵约 470MB 的树并各自序列化一遍，峰值 1043MB、xml_fix
    # 也被四次序列化拖长。修改顺序与改动前逐字一致：字体→引号→标题加粗→目录样式。
    check_cancel()
    fixes = docx_xml.apply_all_fixes(docx_out)
    n_quotes = fixes.get("chinese_quotes")
    if n_quotes is not None:
        print(f"已将 {n_quotes} 个中文引号字符设为宋体渲染（页眉页脚同样处理）")
    n_headings = fixes.get("headings_bolded")
    if n_headings:
        print(f"已将 {n_headings} 个源标题的加粗迁移到段落样式"
              f"（更新目录后条目不再继承加粗）")
    n_toc_styles = fixes.get("toc_styles")
    print(f"已将目录样式 TOC 1~9 设为宋体+Times New Roman、小四、蓝色、不加粗"
          f"（共 {n_toc_styles} 个样式；用户在 Word 里更新目录后仍保持）")
    _lap("xml_fix")

    # ---- 第 6 步：最终复检（独立 Word 实例，只检测）----
    check_cancel()
    remaining, pages3 = _recheck_empty_pages(docx_out)
    _lap("step3_recheck")
    info["step3_recheck"] = {"pages": pages3,
                             "remaining_empty_pages": remaining,
                             "stage": "final"}
    # ---- 第 7 步：最终验收（字体覆盖率 + 目录域）----
    fonts = docx_xml._verify_final_docx(docx_out)
    _lap("verify")

    # 先 gc 再读内存：否则 end_mb 反映的是尚未回收的文档树残留
    gc.collect()

    info.update({
        "docx_size": docx_out.stat().st_size,
        "docx_sha256": _sha256(docx_out),
        "durations": durations,
        "pages": pages3 or pages2,
        "remaining_empty_pages": remaining,
        "fonts": fonts,
        "toc_styles": n_toc_styles,
        "chinese_quotes": n_quotes,
        "headings_bolded": n_headings,
        "word_version": word_ver,
        "duration_sec": round(time.time() - t0, 1),
        "finished_at": datetime.now().isoformat(timespec="seconds"),
        "memory": meminfo.snapshot(mem_start),
    })
    # 正常结束也清理一次：ps1 应已 Quit，若因异常没退干净，按本次任务 PID 精确回收
    _kill_headless_word()
    return PostResult(info)


def _recheck_empty_pages(docx_out: str | os.PathLike, max_rounds: int = 5,
                         timeout: int = 900) -> tuple[int, int | None]:
    """独立 Word 实例按最终分页复检空白页（只检测，不写盘）。

    返回 (残留空白页数, 文档总页数)。只检测模式下第一轮即结束；
    max_rounds 保留以兼容旧调用。
    """
    remaining_final = 0
    pages_final: int | None = None
    for _round in range(max(1, max_rounds)):
        try:
            proc = _run_ps("step3_recheck.ps1", [os.fspath(docx_out)], timeout=timeout)
        except subprocess.TimeoutExpired:
            _kill_headless_word()
            print(f"最终复检第 {_round + 1} 轮超时（>{timeout}s），提前结束")
            break
        if proc.returncode != 0:
            _kill_headless_word()
            print(f"最终复检第 {_round + 1} 轮失败：\n"
                  + (proc.stderr or proc.stdout or "未知错误"))
            break
        remaining = _grab(proc.stdout, "EMPTY_PAGES_COUNT=")
        pages = _grab(proc.stdout, "DOC_PAGES=")
        _collect_word_pids(proc.stdout)
        if remaining is not None:
            remaining_final = remaining
        if pages is not None:
            pages_final = pages
        if remaining:
            print(f"警告：最终复检仍检测到 {remaining} 个空白页"
                  f"（源头方案未完全消除），请在 Word 中核对。")
        n = _grab(proc.stdout, "SPACE_CHARS_DELETED=")
        if not n:
            break
        time.sleep(3)          # 等待上一 Word 实例完全退出、文件句柄释放
    return remaining_final, pages_final
