# -*- coding: utf-8 -*-
"""rtf_merge GUI —— 常规软件式操作界面（tkinter 标准库实现，零额外依赖）。

用法：
    python rtf_merge_gui.py                 # 直接运行
    或打包：pyinstaller --onefile --windowed --name rtf_merge_gui rtf_merge_gui.py

界面功能：
    · 选择输入文件夹（递归查找 *.rtf）与输出文件（.rtf）
    · 勾选：完整后处理（--fix-fonts）、人工排序（--order 1）、修正页码（--pageyn）
    · 可选：删除指定行（--delline）、输入编码
    · 空白页处理默认逐页弹窗询问（与原 VBA FindEmptyPages 一致），
      可勾选“自动跳过空白页”改为不询问
    · 运行日志实时显示，完成/出错弹窗提示

后台实际调用 rtf_merge 模块（merge_rtf / fix_fonts_via_word），
与命令行版完全同源，不依赖 pywin32。
"""
import os
import sys
import queue
import threading
import logging
import traceback

# 高 DPI 适配（高分屏不模糊）
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import rtf_merge  # 与 rtf_merge.py 同目录；打包时 PyInstaller 自动收录

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
    encoding/fix_fonts/empty_auto_no。完成后向 RESULT_QUEUE 放结果。"""
    try:
        # GUI（--windowed）下 stdout/stderr 为 None，先替换为队列写入器
        sys.stdout = QueueWriter(LOG_QUEUE)
        sys.stderr = QueueWriter(LOG_QUEUE)
        if opts.get("empty_auto_no"):
            os.environ["EMPTY_PAGE_AUTO"] = "NO"
        else:
            os.environ.pop("EMPTY_PAGE_AUTO", None)

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
                "empty_output": "合并结果为空（输入文件可能都是空文件），已拒绝写出。",
            }
            LOG_QUEUE.put(">>> 未生成合并结果（状态：%s）" % result["status"])
            RESULT_QUEUE.put(("error", hints.get(result["status"],
                                                 "合并未完成，详见日志。")))
            return

        LOG_QUEUE.put(">>> 成功合并 %d 个 RTF 文件 -> %s"
                      % (len(result["files"]), result["outfile"]))
        if opts.get("fix_fonts"):
            LOG_QUEUE.put(">>> 正在用 Word 执行字体修正与后处理"
                          "（合并单元格/表格调整/目录/空白页/引号等，"
                          "视文档大小约 10~20 分钟，请耐心等待）...")
            docx = rtf_merge.fix_fonts_via_word(result["outfile"])
            LOG_QUEUE.put(">>> 后处理完成，另存为：%s" % docx)
        RESULT_QUEUE.put(("done", "任务完成。"))
    except Exception as exc:
        LOG_QUEUE.put(">>> 错误：%s" % exc)
        LOG_QUEUE.put(traceback.format_exc())
        RESULT_QUEUE.put(("error", str(exc)))


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
        ttk.Entry(row2, textvariable=self.var_out).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row2, text="浏览...", command=self.browse_out).pack(side="left")

        # ===== 选项区 =====
        frm_opt = ttk.LabelFrame(root, text=" 2. 选项 ")
        frm_opt.pack(fill="x", padx=10, pady=6)
        self.var_fix = tk.BooleanVar(value=True)
        self.var_order = tk.BooleanVar(value=False)
        self.var_pageyn = tk.BooleanVar(value=True)
        self.var_empty = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm_opt, text="完整后处理（--fix-fonts：字体修正、关键词单元格合并、"
                                      "表格自动调整、插入目录、删除空白页、中文引号宋体、"
                                      "目录更新不加粗；需要本机装有 Word）",
                        variable=self.var_fix).pack(anchor="w", **pad)
        ttk.Checkbutton(frm_opt, text="人工排序（--order 1：读取 file_order.xlsx/.xls）",
                        variable=self.var_order).pack(anchor="w", **pad)
        ttk.Checkbutton(frm_opt, text="修正页码（--pageyn，默认勾选）",
                        variable=self.var_pageyn).pack(anchor="w", **pad)
        ttk.Checkbutton(frm_opt, text="空白页自动处理（默认勾选：不逐页弹窗；"
                                      "空白页已在合并阶段从源头消除，此处仅兜底检测）",
                        variable=self.var_empty).pack(anchor="w", **pad)

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
    def browse_in(self):
        d = filedialog.askdirectory(title="选择包含 RTF 的文件夹")
        if d:
            self.var_in.set(d)
            # 输出默认放在输入目录的【上一级】：避免产物落进输入目录被下次
            # 扫描重复合并（README 第 55 行明确警告，原默认值恰好违反）
            if not self.var_out.get().strip():
                parent = os.path.dirname(os.path.abspath(d.rstrip("\\/"))) or d
                self.var_out.set(os.path.join(parent, "all_merged.rtf"))

    def browse_out(self):
        f = filedialog.asksaveasfilename(
            title="选择输出文件", defaultextension=".rtf",
            filetypes=[("RTF 文件", "*.rtf"), ("所有文件", "*.*")])
        if f:
            self.var_out.set(f)

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
            "empty_auto_no": self.var_empty.get(),
        }
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")
        self.btn_start.config(state="disabled")
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
        if kind == "done":
            self.lbl_status.config(text="完成")
            messagebox.showinfo("完成", msg + "\n请查看输出目录中的文件。")
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

    def on_close(self):
        if self.thread is not None and self.thread.is_alive():
            if not messagebox.askyesno("确认", "任务仍在运行，退出将中断任务。确定退出？"):
                return
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
