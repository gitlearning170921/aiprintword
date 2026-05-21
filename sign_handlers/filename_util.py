# -*- coding: utf-8 -*-
"""展示用文件名规范化（保留中文；修复 UTF-8 被误读为 Latin-1 的乱码）。"""
from __future__ import annotations

import os
import re


def is_internal_cache_filename(name: str) -> bool:
    base = os.path.basename((name or "").replace("\\", "/"))
    low = base.lower()
    if low.startswith(("_ftptpl_", "_dbtpl_", "link_")):
        return True
    if re.match(
        r"^_ftptpl_[0-9a-f\-]{36}(?:_\d{14})?\.(docx|doc|xlsx|xls|pdf)$",
        low,
    ):
        return True
    return False


def normalize_display_filename(name: str, *, default: str = "document.docx") -> str:
    s = (name or "").strip().replace("\\", "/")
    s = os.path.basename(s) or s
    if not s or is_internal_cache_filename(s):
        return default
    if _looks_like_utf8_mojibake(s):
        repaired = _repair_utf8_mojibake(s)
        if repaired:
            s = repaired
    return s[:512]


def _looks_like_utf8_mojibake(s: str) -> bool:
    if not s or not any(ord(c) > 127 for c in s):
        return False
    if re.search(r"[\u4e00-\u9fff]", s):
        return False
    high = sum(1 for c in s if 0xC0 <= ord(c) <= 0xFF or c in "ÃÂÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ")
    return high >= max(2, len(s) // 4)


def _repair_utf8_mojibake(s: str) -> str:
    for enc in ("latin-1", "cp1252"):
        try:
            out = s.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if out and out != s and re.search(r"[\u4e00-\u9fff]", out):
            return out
        if out and out != s and re.search(r"[\u0400-\u04FF\w.\-()（）\s]", out):
            bad = sum(1 for c in out if ord(c) in range(0xC0, 0x100))
            if bad < len(out) // 3:
                return out
    return ""
