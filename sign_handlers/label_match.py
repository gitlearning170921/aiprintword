# -*- coding: utf-8 -*-
"""签名占位：单元格/短标签与关键词等价匹配（中英、冒号、大小写）。"""
from __future__ import annotations

import re


def cell_text_matches_keyword(cell_text, keyword: str) -> bool:
    """
    单元格整格为「标签」时成立，例如：Author、Author:、author：、作者、作者：
    不匹配长句中的子串，也不匹配 Authors 等更长词。
    """
    if cell_text is None:
        return False
    s = str(cell_text).strip()
    kw = str(keyword).strip()
    if not s or not kw:
        return False
    esc = re.escape(kw)
    if all(ord(c) < 128 for c in kw):
        pat = re.compile(r"^" + esc + r"\s*[:：]?\s*$", re.IGNORECASE)
    else:
        pat = re.compile(r"^" + esc + r"\s*[:：]?\s*$")
    return bool(pat.match(s))


def paragraph_text_keyword_end_offset(full_text: str, keyword: str) -> int:
    """
    段落中「关键词 + 可选冒号」的结束位置，用于在其后插入图片。
    英文关键词使用词边界，避免 author 匹配到 authors。
    """
    if not full_text or not keyword:
        return -1
    esc = re.escape(keyword)
    if all(ord(c) < 128 for c in keyword):
        m = re.search(r"(?i)\b" + esc + r"\b\s*[:：]?", full_text)
    else:
        m = re.search(re.escape(keyword) + r"\s*[:：]?", full_text)
    return m.end() if m else -1
