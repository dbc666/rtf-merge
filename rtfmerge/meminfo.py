# -*- coding: utf-8 -*-
"""进程内存读数（零依赖，仅 Windows 有效；其它平台返回 (None, None)）。

后处理峰值与收尾占用要写进合并清单，GUI 日志也要用同一套口径，所以抽成
独立模块，避免两处各写一份 ctypes 结构体、读数口径漂移。

调用方一律先 `gc.collect()` 再读，否则读到的是尚未回收的文档树残留
（lxml 元素树有循环引用，`del` 不等于释放）。
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt

_MB = 1048576.0


class _PMC(ctypes.Structure):
    _fields_ = [("cb", wt.DWORD), ("PageFaultCount", wt.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t)]


_BOUND = None


def _bind():
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.GetCurrentProcess.restype = wt.HANDLE
    psapi.GetProcessMemoryInfo.argtypes = [wt.HANDLE, ctypes.POINTER(_PMC),
                                           wt.DWORD]
    psapi.GetProcessMemoryInfo.restype = wt.BOOL
    return psapi, k32


def current():
    """返回 (工作集 MB, 峰值工作集 MB)；取不到时返回 (None, None)。"""
    global _BOUND
    try:
        if _BOUND is None:
            _BOUND = _bind()
        psapi, k32 = _BOUND
        pmc = _PMC()
        pmc.cb = ctypes.sizeof(pmc)
        if not psapi.GetProcessMemoryInfo(k32.GetCurrentProcess(),
                                          ctypes.byref(pmc), pmc.cb):
            return (None, None)
        return (pmc.WorkingSetSize / _MB, pmc.PeakWorkingSetSize / _MB)
    except Exception:
        return (None, None)


def snapshot(start_mb=None):
    """给合并清单用的三元组（start 由调用方在任务开始时采样）。"""
    ws, peak = current()
    return {
        "start_mb": None if start_mb is None else round(start_mb, 1),
        "peak_mb": None if peak is None else round(peak, 1),
        "end_mb": None if ws is None else round(ws, 1),
    }
