# -*- coding: utf-8 -*-
"""rtf_merge GUI —— 常规软件式操作界面（tkinter 标准库实现，零额外依赖）。

用法：
    python rtf_merge_gui.py                 # 直接运行
    或打包：pyinstaller --onefile --windowed --name rtf_merge_gui rtf_merge_gui.py

界面功能：
    · 选择输入文件夹（递归查找 *.rtf）与输出文件（.rtf，不写后缀也自动补）
    · 勾选：完整后处理（--fix-fonts）、人工排序（--order 1）、修正页码（--pageyn）
    · 可选：删除指定行（--delline）、输入编码
    · 空白页为自动检测（只读统计，不删页、不弹窗）；空白页已从合并阶段
      源头消除，检测结果只作留痕与兜底提示
    · 运行中可点“停止”中断；关窗时确认后中断并清理本次产生的无窗口 Word
    · 运行日志实时显示，完成/出错/已取消弹窗提示

后台实际调用 rtf_merge 模块（merge_rtf / fix_fonts_via_word），
与命令行版完全同源，不依赖 pywin32。
"""
import os
import sys
import gc
import queue
import threading
import logging
import time
import traceback

# 高 DPI 适配（高分屏不模糊）
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from . import merge_core as rtf_merge   # 合并/后处理核心（包内模块）
from . import meminfo                   # 进程内存读数（与清单共用口径）
from .word_post import (CANCEL_EVENT, TaskCancelled, fix_fonts_via_word,
                        _kill_headless_word)

LOG_QUEUE = queue.Queue()
RESULT_QUEUE = queue.Queue()


class QueueWriter:
    """承接 print() 输出，写入日志队列（GUI 无控制台窗口）。"""

    def __init__(self, q):
        self._q = q

    def write(self, s):
        s = (s or "").rstrip()
        if s:
            self._q.put(s)

    def flush(self):
        pass


class QueueLogHandler(logging.Handler):
    """承接 logging 输出（rtf_merge 模块的 logger），写入日志队列。"""

    def emit(self, record):
        try:
            msg = self.format(record)
            if msg:
                LOG_QUEUE.put(msg)
        except Exception:
            pass


def _mem_mb():
    """当前进程工作集与峰值（MB）。零依赖（ctypes），仅 Windows 有效。

    实现在 `rtfmerge.meminfo`（与写进合并清单的那份共用同一口径）。
    失败时返回 (None, None)，调用方据此跳过内存那行日志。
    """
    return meminfo.current()


def install_logging():
    """给 rtf_merge 的 logger 挂队列 handler（不重复、不依赖 basicConfig）。"""
    lg = logging.getLogger("rtf_merge")
    lg.setLevel(logging.INFO)
    h = QueueLogHandler()
    h.setFormatter(logging.Formatter("%(message)s"))
    lg.handlers.clear()
    lg.addHandler(h)


def run_task(opts):
    """后台线程执行体。opts: dict —— inpath/outfile/order/pageyn/delline/
    encoding/fix_fonts。完成后向 RESULT_QUEUE 放结果。"""
    t0 = time.time()
    ws0, _ = _mem_mb()
    try:
        # 每次任务开始重置取消信号（“停止”按钮 / 关窗会置位）
        CANCEL_EVENT.clear()
        # GUI（--windowed）下 stdout/stderr 为 None，先替换为队列写入器
        sys.stdout = QueueWriter(LOG_QUEUE)
        sys.stderr = QueueWriter(LOG_QUEUE)

        LOG_QUEUE.put(">>> 开始合并（输入目录：%s）" % opts["inpath"])
        result = rtf_merge.merge_rtf(
            inpath=opts["inpath"],
            outfile=opts["outfile"],
            order=1 if opts.get("order") else 0,
            pageyn=1 if opts.get("pageyn") else 0,
            delline=opts.get("delline") or None,
            encoding=opts.get("encoding") or "utf-8",
        )
        for w in result.get("warnings", []):
            LOG_QUEUE.put("警告: %s" % w)

        if result["status"] != "merged":
            hints = {
                "no_files": "输入目录里没有找到 RTF 文件。",
                "no_ordered_files": "排序表 file_order.xlsx 的 order 列全为空，请填写后再运行"
                                    "（或把“按文件名排序”设为 0）。",
                "order_incomplete": "排序表与目录内文件不一致（有文件没填 order），"
                                    "为避免漏合并已中止。请补全 order 列后重试。",
                "empty_output": "合并结果为空（输入文件可能都是空文件），已拒绝写出。",
                "structure_invalid": "合并文件已写出，但 RTF 结构校验未通过"
                                     "（花括号不配平或首末行异常），已停止 Word 后处理。"
                                     "请检查输入文件，结构详情见清单 rtf_validation。",
            }
            LOG_QUEUE.put(">>> 未进入后处理（状态：%s）" % result["status"])
            RESULT_QUEUE.put(("error", hints.get(result["status"],
                                                 "合并未完成，详见日志。")))
            return

        LOG_QUEUE.put(">>> 成功合并 %d 个 RTF 文件 -> %s"
                      % (len(result["files"]), result["outfile"]))
        LOG_QUEUE.put(">>> 合并阶段耗时：%.1f 秒" % (time.time() - t0))
        if opts.get("fix_fonts"):
            LOG_QUEUE.put(">>> 正在用 Word 执行字体修正与后处理"
                          "（合并单元格/表格调整/目录/空白页/引号等，"
                          "视文档大小约 10~20 分钟，请耐心等待）...")
            post = fix_fonts_via_word(result["outfile"])
            LOG_QUEUE.put(">>> 后处理完成，另存为：%s" % post)
            durs = post.get("durations") or {}
            if durs:
                LOG_QUEUE.put(">>> 后处理分段耗时：%s（合计 %.1f 秒）" % (
                    " | ".join("%s %.1fs" % (k, v) for k, v in durs.items()),
                    post.get("duration_sec") or sum(durs.values())))
            if result.get("manifest"):
                # 清单要能回答"当时选了哪个目录、勾了什么"，故记快照而非 "gui"
                post["command_line"] = "gui: " + " ".join(
                    "%s=%s" % (k, opts.get(k)) for k in
                    ("inpath", "outfile", "order", "pageyn", "fix_fonts",
                     "delline", "encoding"))
                post["mode"] = "gui"
                post["options"] = dict(opts)
                rtf_merge.update_manifest_postprocess(result["manifest"], post)
                LOG_QUEUE.put(">>> 合并清单（含后处理留痕）：%s" % result["manifest"])
        RESULT_QUEUE.put(("done", "任务完成。"))
    except TaskCancelled:
        LOG_QUEUE.put(">>> 已按用户请求停止任务。")
        _kill_headless_word()
        RESULT_QUEUE.put(("cancel", "任务已取消。"))
    except Exception as exc:
        LOG_QUEUE.put(">>> 错误：%s" % exc)
        LOG_QUEUE.put(traceback.format_exc())
        RESULT_QUEUE.put(("error", str(exc)))
    finally:
        # 无论完成/取消/出错都报总耗时，便于事后核对"这次为什么慢"。
        # 先 gc 再测内存：lxml 的文档树有循环引用，不 gc 时这里读到的仍是
        # 任务期间的残留（实测同一进程 del 后 966MB、gc 一次后 37MB），
        # 会把"当前占用"报得虚高。
        gc.collect()
        elapsed = time.time() - t0
        LOG_QUEUE.put(">>> 总耗时：%.1f 秒（约 %.1f 分钟）"
                      % (elapsed, elapsed / 60.0))
        ws, peak = _mem_mb()
        if ws is not None:
            LOG_QUEUE.put(">>> 进程内存：全程峰值 %.0f MB，当前 %.0f MB%s"
                          % (peak, ws,
                             ("（起始 %.0f MB）" % ws0) if ws0 else ""))


def _set_window_icon(root):
    """设置窗口/任务栏图标（打包时随 --add-data 携带 rtf_merge.ico）。"""
    try:
        base = getattr(sys, "_MEIPASS",
                       os.path.dirname(os.path.abspath(__file__)))
        ico = os.path.join(base, "rtf_merge.ico")
        if os.path.isfile(ico):
            root.iconbitmap(ico)
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("rtf_merge —— RTF 合并工具")
        root.geometry("820x680")
        root.minsize(720, 560)
        _set_window_icon(root)
        self.thread = None

        pad = {"padx": 8, "pady": 4}

        # ===== 路径区 =====
        frm_path = ttk.LabelFrame(root, text=" 1. 文件路径 ")
        frm_path.pack(fill="x", padx=10, pady=6)
        self.var_in = tk.StringVar()
        self.var_out = tk.StringVar()

        row1 = ttk.Frame(frm_path)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="输入文件夹（递归查找 *.rtf）：", width=28).pack(side="left")
        ttk.Entry(row1, textvariable=self.var_in).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row1, text="浏览...", command=self.browse_in).pack(side="left")

        row2 = ttk.Frame(frm_path)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="输出文件（.rtf，可不写后缀）：", width=28).pack(side="left")
        self.ent_out = ttk.Entry(row2, textvariable=self.var_out)
        self.ent_out.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row2, text="浏览...", command=self.browse_out).pack(side="left")

        # 单行输入框放不下长路径（只会显示尾部），下面这行灰色小字始终
        # 完整回显一次，避免用户以为"只填了文件名"而重复手打目录。
        self.lbl_out_full = ttk.Label(frm_path, text="", foreground="#555555",
                                      justify="left", wraplength=760)
        self.lbl_out_full.pack(fill="x", padx=8, pady=(0, 4))
        self.var_out.trace_add("write", lambda *_: self.sync_out_label())
        self.sync_out_label()
        # 输入框里贴路径（而不是点"浏览..."）同样要带出默认输出名。
        # 逐字符键入时大部分中间态不是有效目录，isdir 一挡就不会乱填；
        # 只有贴完整路径（或敲完整路径）那一次才会推导。
        self.var_in.trace_add("write", lambda *_: self.on_in_typed())

        # ===== 选项区 =====
        frm_opt = ttk.LabelFrame(root, text=" 2. 选项 ")
        frm_opt.pack(fill="x", padx=10, pady=6)
        self.var_fix = tk.BooleanVar(value=True)
        self.var_order = tk.BooleanVar(value=False)
        self.var_pageyn = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm_opt, text="完整后处理（--fix-fonts：字体修正、关键词单元格合并、"
                                      "表格自动调整、插入目录、空白页检测、中文引号宋体、"
                                      "目录更新不加粗；需要本机装有 Word）",
                        variable=self.var_fix).pack(anchor="w", **pad)
        ttk.Checkbutton(frm_opt, text="人工排序（--order 1：读取 file_order.xlsx/.xls）",
                        variable=self.var_order).pack(anchor="w", **pad)
        ttk.Checkbutton(frm_opt, text="修正页码（--pageyn，默认勾选）",
                        variable=self.var_pageyn).pack(anchor="w", **pad)
        # 空白页环节自 v1.1.2 起为纯只读检测，没有可选项，因此只作说明而不放死控件
        ttk.Label(frm_opt, text="空白页：自动检测（只读统计，不删页、不弹窗；"
                               "空白页已在合并阶段从源头消除）",
                  foreground="#555555").pack(anchor="w", **pad)

        row3 = ttk.Frame(frm_opt)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="删除指定行（--delline，如 85-87 或 85 86 87）：",
                  width=40).pack(side="left")
        self.var_delline = tk.StringVar()
        ttk.Entry(row3, textvariable=self.var_delline, width=20).pack(side="left")

        row4 = ttk.Frame(frm_opt)
        row4.pack(fill="x", **pad)
        ttk.Label(row4, text="输入编码（--encoding，默认 utf-8）：", width=40).pack(side="left")
        self.var_encoding = tk.StringVar(value="utf-8")
        ttk.Entry(row4, textvariable=self.var_encoding, width=20).pack(side="left")

        # ===== 按钮区 =====
        frm_btn = ttk.Frame(root)
        frm_btn.pack(fill="x", padx=10, pady=6)
        self.btn_start = ttk.Button(frm_btn, text="开始合并", command=self.start)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(frm_btn, text="停止", command=self.stop,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.btn_quit = ttk.Button(frm_btn, text="退出", command=self.on_close)
        self.btn_quit.pack(side="left", padx=4)
        self.lbl_status = ttk.Label(frm_btn, text="就绪")
        self.lbl_status.pack(side="right", padx=4)

        # ===== 日志区 =====
        frm_log = ttk.LabelFrame(root, text=" 3. 运行日志 ")
        frm_log.pack(fill="both", expand=True, padx=10, pady=6)
        self.txt = tk.Text(frm_log, wrap="word", state="disabled",
                           font=("Consolas", 9))
        sb = ttk.Scrollbar(frm_log, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(120, self.poll_log)

    # ---- 交互 ----
    def sync_out_label(self):
        """把输出框内容完整回显到下方灰色小字（单行框会截断长路径）。"""
        v = self.var_out.get().strip()
        txt = v if v else "（尚未设置，选中输入文件夹后自动填入）"
        self.lbl_out_full.config(text="完整输出路径：" + txt)

    def apply_default_out(self, inpath):
        """按输入目录推导默认输出名；只覆盖空值或上次的自动默认值。

        输出默认放在输入目录的【上一级】：避免产物落进输入目录、被下次
        扫描当成输入重复合并（README 明确警告，早期默认值恰好违反）。
        文件名默认取输入文件夹名（如 …\data\空腹_TF → …\data\空腹_TF.rtf）。
        用户手改过的路径不会被下一次推导冲掉。
        """
        norm = os.path.abspath(inpath)      # 已规范化尾部反斜杠（盘根除外）
        parent = os.path.dirname(norm) or norm
        name = os.path.basename(norm) or "合并结果"
        dflt = os.path.join(parent, name + ".rtf")
        cur = self.var_out.get().strip()
        changed = False
        if not cur or cur == getattr(self, "_last_default", None):
            self.var_out.set(dflt)
            # 让单行框从路径开头显示，而不是只露出尾部
            self.ent_out.icursor(0)
            self.ent_out.xview_moveto(0)
            changed = True
        self._last_default = dflt
        return changed

    def on_in_typed(self):
        """输入框内容变化（粘贴/键入）时，只对有效目录补默认输出名。"""
        raw = self.var_in.get().strip().strip('"').strip()
        if not raw or not os.path.isdir(raw):
            return
        # 只接受"本身就是规范化绝对路径"的输入。逐字符键入的中间态会被
        # abspath 改写成别的目录（"F:" 是盘符当前目录的简写，isdir 为真，
        # 但 abspath("F:") = "F:\<该盘当前目录>"，据此推导会填出错误的名字），
        # 相对路径同理不取样。这类输入一律不推导，等路径完整再触发。
        norm = os.path.abspath(raw)
        if os.path.normcase(norm.rstrip("\\/")) != os.path.normcase(raw.rstrip("\\/")):
            return
        self.apply_default_out(raw)

    def browse_in(self):
        d = filedialog.askdirectory(title="选择包含 RTF 的文件夹")
        if d:
            self.var_in.set(d)      # 交给 on_in_typed 统一推导，避免两处逻辑漂移

    def browse_out(self):
        f = filedialog.asksaveasfilename(
            title="选择输出文件", defaultextension=".rtf",
            filetypes=[("RTF 文件", "*.rtf"), ("所有文件", "*.*")])
        if f:
            self.var_out.set(f)
            self.ent_out.icursor(0)
            self.ent_out.xview_moveto(0)

    def start(self):
        inpath = self.var_in.get().strip().strip('"')
        outfile = self.var_out.get().strip().strip('"')
        if not inpath or not os.path.isdir(inpath):
            messagebox.showwarning("提示", "请先选择有效的输入文件夹。")
            return
        if not outfile:
            messagebox.showwarning("提示", "请先设置输出文件路径。")
            return
        # 自动补 .rtf 后缀：用户无需在路径中写后缀
        if not outfile.lower().endswith(".rtf"):
            outfile += ".rtf"
            self.var_out.set(outfile)
            self.sync_out_label()
        if self.thread is not None and self.thread.is_alive():
            messagebox.showwarning("提示", "已有任务在运行，请等待完成。")
            return

        opts = {
            "inpath": inpath,
            "outfile": outfile,
            "order": self.var_order.get(),
            "pageyn": self.var_pageyn.get(),
            "delline": self.var_delline.get().strip() or None,
            "encoding": self.var_encoding.get().strip() or "utf-8",
            "fix_fonts": self.var_fix.get(),
        }
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.lbl_status.config(text="运行中...")
        self.thread = threading.Thread(target=run_task, args=(opts,), daemon=True)
        self.thread.start()
        self.root.after(150, self.check_done)

    def check_done(self):
        try:
            kind, msg = RESULT_QUEUE.get_nowait()
        except queue.Empty:
            self.root.after(200, self.check_done)
            return
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        if kind == "done":
            self.lbl_status.config(text="完成")
            messagebox.showinfo("完成", msg + "\n请查看输出目录中的文件。")
        elif kind == "cancel":
            self.lbl_status.config(text="已取消")
            messagebox.showinfo("已取消", msg)
        else:
            self.lbl_status.config(text="出错")
            messagebox.showerror("错误", msg + "\n详情见运行日志。")

    def poll_log(self):
        try:
            while True:
                line = LOG_QUEUE.get_nowait()
                self.txt.config(state="normal")
                self.txt.insert("end", line + "\n")
                self.txt.see("end")
                self.txt.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(120, self.poll_log)

    def stop(self):
        """请求停止当前任务：置取消信号，当前 PowerShell 步骤会被杀进程中断。"""
        if self.thread is not None and self.thread.is_alive():
            CANCEL_EVENT.set()
            self.lbl_status.config(text="正在停止...")
            self.btn_stop.config(state="disabled")
            LOG_QUEUE.put(">>> 已请求停止，等待当前步骤中断...")

    def on_close(self):
        if self.thread is not None and self.thread.is_alive():
            if not messagebox.askyesno("确认", "任务仍在运行，退出将中断任务。确定退出？"):
                return
            CANCEL_EVENT.set()
            self.thread.join(timeout=10)   # 给 _run_ps 的取消检查一个周期
            _kill_headless_word()          # 清理本工具留下的无窗口 Word
        self.root.destroy()


def main():
    install_logging()
    try:
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception:
        # 打包版无控制台：把崩溃现场写盘便于定位
        try:
            crash = os.path.join(os.environ.get("TEMP", r"C:\Windows\Temp"),
                                 "rtf_merge_gui_crash.log")
            with open(crash, "w", encoding="utf-8") as fh:
                fh.write(traceback.format_exc())
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
