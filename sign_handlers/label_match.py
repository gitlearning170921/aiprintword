# -*- coding: utf-8 -*-
"""签名占位：单元格/短标签与关键词等价匹配（中英、冒号、大小写）。"""
from __future__ import annotations

import re
from typing import List

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


def _rest_is_signoff_date_phrase(rest: str) -> bool:
    """识别「及日期/和日期/与日期//日期」这类签批标签尾巴。"""
    s = str(rest or "").strip()
    if not s:
        return False
    return bool(
        re.match(
            r"^(?:[/／]|及|和|与|and)\s*(?:测试日期|签署日期|日期|date)\s*[:：]?$",
            s,
            re.IGNORECASE,
        )
    )


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
    # 「编制及日期」「批准人/日期」等签批标签
    if m_full and _rest_is_signoff_date_phrase(m_full.group(1)):
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
        if m and _rest_is_signoff_date_phrase(m.group(1)):
            return True
        return bool(re.match(r"^\s*" + esc + r"\s*$", s, re.IGNORECASE))
    if not s.startswith(kw):
        return False
    rest = s[len(kw) :].lstrip()
    if not rest:
        return True
    if _rest_is_blank_or_placeholder(rest):
        return True
    if _rest_is_signoff_date_phrase(rest):
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


def iter_role_label_lines(cell_text) -> List[str]:
    """单元格/段落拆行：签批栏常在同一格内纵向排列 Author/Reviewer/Approver。"""
    raw = str(cell_text or "").strip()
    if not raw:
        return []
    lines: List[str] = [raw]
    if re.search(r"[\r\n]", raw):
        lines = [x.strip() for x in re.split(r"[\r\n]+", raw) if x and x.strip()]
    return lines


# 签批栏「测试人/日期」「编制及日期」「批准人和日期」等；区别于用例表列头仅「测试人」
_SIGNOFF_DATE_IN_LABEL_RE = re.compile(
    r"(?:[/／]|及|和|与)\s*(?:测试日期|签署日期|日期|date)|(?:测试日期|签署日期|日期|date)\s*[:：]",
    re.IGNORECASE,
)


def _signoff_role_date_label_end_offset(text: str, keyword: str) -> int:
    """
    识别「角色 + 连接词 + 日期[:：]」标签结束位置。
    例如：测试人/日期：____、编制及日期：____、Reviewer and Date: ____。
    """
    s = str(text or "")
    kw = str(keyword or "").strip()
    if not s or not kw:
        return -1
    esc = re.escape(kw)
    ascii_only = all(ord(c) < 128 for c in kw)
    use_word_boundary = ascii_only and bool(re.match(r"^[A-Za-z0-9_]+$", kw))
    joiner = r"(?:[/／]|及|和|与|and)"
    date_tag = r"(?:测试日期|签署日期|日期|date)"
    if use_word_boundary:
        m = re.search(r"(?i)\b" + esc + r"\b\s*" + joiner + r"\s*" + date_tag + r"\s*[:：]?", s)
    else:
        m = re.search(r"(?i)" + esc + r"\s*" + joiner + r"\s*" + date_tag + r"\s*[:：]?", s)
    return m.end() if m else -1


def cell_looks_like_signoff_date_label(cell_text) -> bool:
    """是否为「测试人/日期」「复核人/日期」类签批栏标签格（非独立 Date 列表头）。"""
    return bool(_SIGNOFF_DATE_IN_LABEL_RE.search(str(cell_text or "")))


def cell_has_label_inline_reservation(cell_text, keyword: str) -> bool:
    """
    标签格内留位：关键词后仅空格/下划线/点线（无冒号亦可），签字应紧挨标签后插入。
    适用于作者/审核/批准/测试人等所有角色。
    """
    if cell_text is None or not str(keyword or "").strip():
        return False
    kw = str(keyword).strip()

    def _line_has_tail(line: str) -> bool:
        line = str(line or "").strip()
        if not line:
            return False
        if cell_text_matches_keyword(line, kw):
            # 「测试人/日期：」等签批栏整格仅标签：签字在右侧/下方空白格，非标签格内
            if _SIGNOFF_DATE_IN_LABEL_RE.search(line):
                return False
            return True
        esc = re.escape(kw)
        if all(ord(c) < 128 for c in kw):
            m = re.match(r"^\s*" + esc + r"(.*)$", line, re.IGNORECASE)
        else:
            m = re.match(r"^\s*" + esc + r"(.*)$", line)
        if m and _rest_is_blank_or_placeholder(m.group(1).lstrip(" ：:\t")):
            return True
        return False

    for line in iter_role_label_lines(cell_text):
        if _line_has_tail(line):
            return True
    return False


def cell_is_bare_role_column_header(cell_text, keyword: str) -> bool:
    """
    宽表列头：整格仅「测试人」「审核」等，无 /日期、无同格下划线留位。
    不是签字锚点（签在这里会误报成功且看不见）。
    """
    if not cell_has_role_keyword(cell_text, keyword):
        return False
    raw = str(cell_text or "")
    if _SIGNOFF_DATE_IN_LABEL_RE.search(raw):
        return False
    kw_n = re.sub(r"\s+", "", str(keyword or "").strip())
    if not kw_n:
        return False
    for line in iter_role_label_lines(cell_text):
        line_n = re.sub(r"\s+", "", line)
        if re.fullmatch(re.escape(kw_n) + r"[:：]?", line_n, flags=re.IGNORECASE):
            return True
    return False


def cell_is_role_signoff_label_slot(cell_text, keyword: str) -> bool:
    """签批栏：标签含「角色/日期」（如测试人/日期、复核人/日期）。"""
    if not cell_has_role_keyword(cell_text, keyword):
        return False
    if _SIGNOFF_DATE_IN_LABEL_RE.search(str(cell_text or "")):
        return True
    esc = re.escape(str(keyword or "").strip())
    if esc and re.search(esc + r"\s*[/／]\s*日期", str(cell_text or ""), re.IGNORECASE):
        return True
    return False


def cell_has_signoff_inline_reservation(cell_text, keyword: str) -> bool:
    """
    签批栏同格留位：「执行人/日期：____」等，签名与日期应紧挨标签后拼接（非右侧分列）。
    """
    if not cell_is_role_signoff_label_slot(cell_text, keyword):
        return False
    for line in iter_role_label_lines(cell_text):
        if not cell_has_role_keyword(line, keyword):
            continue
        off = _signoff_role_date_label_end_offset(line, keyword)
        if off >= 0:
            tail = line[off:].lstrip(" \t")
            if not tail or _rest_is_blank_or_placeholder(tail):
                return True
        off = paragraph_text_keyword_end_offset(line, keyword)
        if off < 0:
            continue
        tail = line[off:].lstrip(" ：:\t")
        if not tail or _rest_is_blank_or_placeholder(tail):
            return True
    raw = str(cell_text or "")
    off = _signoff_role_date_label_end_offset(raw, keyword)
    if off >= 0:
        tail = raw[off:].lstrip(" \t")
        return (not tail) or _rest_is_blank_or_placeholder(tail)
    off = paragraph_text_keyword_end_offset(raw, keyword)
    if off < 0:
        return False
    tail = raw[off:].lstrip(" ：:\t")
    return not tail or _rest_is_blank_or_placeholder(tail)


def cell_inline_insert_offset_px(cell_text, keyword: str, *, max_px: int = 300) -> int | None:
    """
    同格标签+占位：返回应紧跟标签后插图的 x 偏移（px）；签批栏「角色/日期：____」亦适用。
    """
    if cell_text is None:
        return None
    txt = str(cell_text)
    off = paragraph_text_keyword_end_offset(txt, keyword)
    if off < 0:
        return None
    tail = txt[off:].lstrip(" ：:\t")
    if cell_is_role_signoff_label_slot(txt, keyword):
        signoff_off = _signoff_role_date_label_end_offset(txt, keyword)
        if signoff_off >= 0:
            signoff_tail = txt[signoff_off:].lstrip(" \t")
            if signoff_tail and not _rest_is_blank_or_placeholder(signoff_tail):
                return None
            off = signoff_off
        else:
            if not tail or not _rest_is_blank_or_placeholder(tail):
                return None
    elif tail and not _rest_is_blank_or_placeholder(tail):
        return None
    tail = txt[off:].lstrip(" ：:\t")
    pre = txt[:off]
    px = 4
    for ch in pre:
        if ch in ("\t",):
            px += 12
        elif ch == " ":
            px += 4
        elif ord(ch) > 127:
            px += 10
        else:
            px += 7
    return min(max(px, 4), max_px)


def cell_has_role_keyword(cell_text, keyword: str) -> bool:
    """整格或任一行格首/整格匹配角色标签（中英、冒号、留空位）。"""
    if cell_text is None or not str(keyword or "").strip():
        return False
    for line in iter_role_label_lines(cell_text):
        if cell_text_matches_keyword(line, keyword):
            return True
        if xlsx_cell_has_leading_role_keyword(line, keyword):
            return True
    return cell_text_matches_keyword(cell_text, keyword) or xlsx_cell_has_leading_role_keyword(
        cell_text, keyword
    )


def paragraph_text_keyword_end_offset(full_text: str, keyword: str) -> int:
    """
    段落中「关键词 + 可选冒号」的结束位置，用于在其后插入图片。
    英文关键词使用词边界，避免 author 匹配到 authors。
    """
    if not full_text or not keyword:
        return -1
    esc = re.escape(keyword)
    # 英文关键词里若包含 ':' '/' 等非“单词字符”，使用 \b...\b 词边界会匹配失败。
    # 仅当关键词本身是纯 [A-Za-z0-9_] 时才使用词边界，避免 Author 匹配到 Authors。
    ascii_only = all(ord(c) < 128 for c in keyword)
    use_word_boundary = ascii_only and bool(re.match(r"^[A-Za-z0-9_]+$", str(keyword or '')))
    if ascii_only and use_word_boundary:
        m = re.search(r"(?i)\b" + esc + r"\b\s*[:：]?", full_text)
    else:
        m = re.search(r"(?i)" + esc + r"\s*[:：]?", full_text)
    return m.end() if m else -1
