#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""rtf_merge_gui —— 兼容入口（v1.1.0 起实现位于 rtfmerge/gui.py）。"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from rtfmerge.gui import main   # noqa: E402

if __name__ == "__main__":
    main()
