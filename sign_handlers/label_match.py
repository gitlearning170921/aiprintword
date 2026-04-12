# -*- coding: utf-8 -*-
"""签名占位：单元格/短标签与关键词等价匹配（中英、冒号、大小写）。"""
from __future__ import annotations

import re

# 关键词后无冒号时，仅空白/下划线/横线/方框等视为「留空签字位」（与 sign_docx 占位语义一致）
_PLACEHOLDER_TAIL_RE = re.compile(
    r"^[\s\u00a0\u1680\u2000-\u200a\u202f\u205f\u3000\uFEFF"
    r"_\-—–\u2014\u2015\u2500~～·…．.□☐▢○◇◯]+$",
    re.UNICODE,
)


def _rest_is_blank_or_placeholder(rest: str) -> bool:
    if not rest:
        return True
    return bool(_PLACEHOLDER_TAIL_RE.match(rest))


def cell_text_matches_keyword(cell_text, keyword: str) -> bool:
    """
    单元格整格为「标签」时成立，例如：Author、Author:、作者、编制人：、编制人____（无冒号仅下划线/空格留位）。
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
        m_full = re.match(r"^" + esc + r"(.*)$", s, re.IGNORECASE)
    else:
        pat = re.compile(r"^" + esc + r"\s*[:：]?\s*$")
        m_full = re.match(r"^" + esc + r"(.*)$", s)
    if pat.match(s):
        return True
    # 「编制人____」「Author ___」等同格标签+留空，无冒号
    if m_full and _rest_is_blank_or_placeholder(m_full.group(1)):
        return True
    return False


def xlsx_cell_has_leading_role_keyword(cell_text, keyword: str) -> bool:
    """
    Word 表格单元格 / Excel 单元格：格首同义词 + 可选冒号，或关键词后仅空白/下划线等签字留空（无冒号亦可）。
    整格仅关键词见 cell_text_matches_keyword（已含「关键词+占位」）。
    """
    if cell_text is None:
        return False
    if cell_text_matches_keyword(cell_text, keyword):
        return True
    s = str(cell_text).strip()
    kw = str(keyword).strip()
    if not s or not kw or len(s) > 96:
        return False
    has_colon = "：" in s or ":" in s
    if all(ord(c) < 128 for c in kw):
        esc = re.escape(kw)
        if re.match(r"^\s*" + esc + r"\s*[:：]", s, re.IGNORECASE):
            return True
        m = re.match(r"^\s*" + esc + r"(.*)$", s, re.IGNORECASE)
        if m and _rest_is_blank_or_placeholder(m.group(1)):
            return True
        return bool(re.match(r"^\s*" + esc + r"\s*$", s, re.IGNORECASE))
    if not s.startswith(kw):
        return False
    rest = s[len(kw) :].lstrip()
    if not rest:
        return True
    if _rest_is_blank_or_placeholder(rest):
        return True
    if rest[0] in (":", "："):
        return True
    if has_colon:
        return True
    if rest and "\u4e00" <= rest[0] <= "\u9fff":
        if len(rest) >= 2:
            return False
        # 避免「审核人员签」类：以「…人员」为关键词时，单字「签名章印」不作标签格
        if len(rest) == 1 and not has_colon and "人员" in kw and rest[0] in "签名章印":
            return False
        return len(s) <= len(kw) + 2
    return True


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
