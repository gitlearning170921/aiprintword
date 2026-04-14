# -*- coding: utf-8 -*-
"""
Word .docx：在模板已预留的空白处插入签名/日期 PNG（不新增签名字段）。
定位顺序：
1) 表格内「角色/姓名格 + 邻列 Date/日期 表头」→ 签名只插角色侧（姓名后或尾部空白），日期只插日期列；
2) 三列表「角色+姓名 | 空签字列 | 日期列」→ 签名插中间格，日期插日期列；
3) 原逻辑：关键词单元格 → 右侧空单元格；段落内关键词后的占位符/下划线 run。
"""
from __future__ import annotations

import io
import os
import re
from typing import Optional

from docx import Document
from docx.document import Document as DocumentObject
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

from sign_handlers.config import ROLE_ID_TO_KEYWORD, role_keywords
from sign_handlers.label_match import (
    cell_text_matches_keyword,
    paragraph_text_keyword_end_offset,
    xlsx_cell_has_leading_role_keyword,
)

# 可见占位字符（作者：____ 等）
_PLACEHOLDER_CHARS = re.compile(r"^[\s_\-—–\u2014\u2015\u2500\u3000\.·]+$")
# 右侧「日期」列表头（整格以 Date/日期 开头）
_DATE_HEADER_CELL = re.compile(r"^\s*(日期|Date)\s*[:：]?\s*", re.IGNORECASE)

_PIC_WIDTH = Cm(2.8)
_PIC_WIDTH_SMALL = Cm(2.0)


def _is_emptyish_text(s: str) -> bool:
    if s is None:
        return True
    t = str(s).strip()
    if not t:
        return True
    if _PLACEHOLDER_CHARS.match(t):
        return True
    return len(t) < 2


def _cell_text(cell) -> str:
    return (cell.text or "").strip()


def _paragraph_full_text(p: Paragraph) -> str:
    return p.text or ""


def _run_has_underline_or_border(r) -> bool:
    try:
        f = r.font
        if f is not None and f.underline:
            return True
    except Exception:
        pass
    try:
        if r._element.rPr is not None:
            ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            rpr = r._element.rPr
            if rpr.find(f"{ns}u") is not None:
                return True
    except Exception:
        pass
    return False


def _find_keyword_in_paragraph(p: Paragraph, keyword: str) -> int:
    return paragraph_text_keyword_end_offset(_paragraph_full_text(p), keyword)


def _clear_runs_after_offset(p: Paragraph, start_char_offset: int) -> None:
    """从段落文本的 start_char_offset 起删除文本（用于去掉 ____ 占位）。"""
    full = _paragraph_full_text(p)
    if start_char_offset <= 0 or start_char_offset >= len(full):
        return
    if not full[start_char_offset:].strip():
        return
    acc = 0
    truncate_after = None
    for i, r in enumerate(list(p.runs)):
        rt = r.text or ""
        ln = len(rt)
        if acc + ln <= start_char_offset:
            acc += ln
            continue
        if acc >= start_char_offset:
            r.text = ""
            acc += ln
            continue
        cut = start_char_offset - acc
        r.text = rt[:cut]
        truncate_after = i
        break
    if truncate_after is not None:
        for r in list(p.runs)[truncate_after + 1 :]:
            r.text = ""


def _insert_pictures_in_paragraph(p: Paragraph, sig_png: bytes, date_png: bytes) -> None:
    p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
    p.add_run(" ")
    p.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_SMALL)


def _cell_looks_like_date_header_cell(text: str) -> bool:
    t = (text or "").strip()
    return bool(_DATE_HEADER_CELL.match(t))


def _char_is_signature_blank_placeholder(c: str) -> bool:
    if not c:
        return False
    if c in " \t\n\r_\u00a0":
        return True
    if c in "-—–\u2014\u2015\u2500\u3000.·…．~～":
        return True
    return False


def _sig_image_insert_char_offset(full: str, kw_end: int) -> int:
    """
    签名图插入点（字符下标）：关键词结束后，若有「姓名等正文 + 尾部下划线空白」则插在空白起点；
    若关键词后整段为占位则插在关键词后；否则插在段落末尾（姓名后无下划线时）。
    """
    if kw_end < 0:
        return len(full)
    if kw_end >= len(full):
        return len(full)
    tail = full[kw_end:]
    if _is_emptyish_text(tail):
        return kw_end
    j = len(full)
    while j > kw_end and _char_is_signature_blank_placeholder(full[j - 1]):
        j -= 1
    if j < len(full):
        return j
    return len(full)


def _find_date_label_end_in_paragraph(p: Paragraph) -> int:
    full = _paragraph_full_text(p)
    if not full:
        return -1
    m = re.search(r"(?i)(?:\bDate\b|日期)\s*[:：]?", full)
    return m.end() if m else -1


def _insert_sig_only_at_char_offset(p: Paragraph, char_offset: int, sig_png: bytes) -> None:
    # 不能把插入点挪到段落末尾：
    # Word 模板中常用 tab leader/制表符/空格作为“占位留白”，若插入点被挪到末尾，
    # 图片会被推到占位末端（看起来离“Author:/Date:”很远）。
    # 因此保持插入点在“正文末尾/冒号后”，并清除其后的占位字符即可。
    _clear_runs_after_offset(p, char_offset)
    p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)


def _insert_date_only_in_date_cell(date_cell, date_png: bytes) -> bool:
    """在「日期/Date」标签后的空白处只插入日期手写图。"""
    for p in date_cell.paragraphs:
        offd = _find_date_label_end_in_paragraph(p)
        if offd < 0:
            continue
        _clear_runs_after_offset(p, offd)
        p.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_SMALL)
        return True
    if date_cell.paragraphs:
        p0 = date_cell.paragraphs[0]
        for r in list(p0.runs):
            r.text = ""
        p0.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_SMALL)
        return True
    p = date_cell.add_paragraph()
    p.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_SMALL)
    return True


def _insert_sig_in_role_cell_for_adjacent_date_column(
    role_cell, keyword: str, sig_png: bytes
) -> bool:
    for p in role_cell.paragraphs:
        off = _find_keyword_in_paragraph(p, keyword)
        if off < 0:
            continue
        full = _paragraph_full_text(p)
        ins = _sig_image_insert_char_offset(full, off)
        _insert_sig_only_at_char_offset(p, ins, sig_png)
        return True
    return False


def _insert_sig_only_in_empty_cell(cell, sig_png: bytes) -> None:
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    for r in list(p.runs):
        r.text = ""
    p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)


def _try_table_adjacent_date_column(
    table: Table,
    keyword: str,
    sig_png: bytes,
    date_png: bytes,
) -> bool:
    """
    常见审批表：
    - 两列：左格「角色/姓名 + 签字空白」+ 右邻格「Date:/日期: + 空白」→ 签名只插左格，日期只插右格。
    - 三列：角色+姓名 | 空签字列 | 日期列 → 签名优先插左格关键词/冒号后（更贴近字段标记），日期插右格。
    避免两段图都落在左侧姓名后。
    """
    for row in table.rows:
        cells = row.cells
        for ci in range(len(cells) - 1):
            left = cells[ci]
            lt = _cell_text(left)
            if not xlsx_cell_has_leading_role_keyword(lt, keyword):
                continue
            date_idx = None
            if ci + 1 < len(cells) and _cell_looks_like_date_header_cell(_cell_text(cells[ci + 1])):
                date_idx = ci + 1
            elif ci + 2 < len(cells) and _cell_looks_like_date_header_cell(_cell_text(cells[ci + 2])):
                date_idx = ci + 2
            if date_idx is None:
                continue
            date_cell = cells[date_idx]
            if date_idx == ci + 1:
                if not _insert_sig_in_role_cell_for_adjacent_date_column(left, keyword, sig_png):
                    continue
            else:
                mid = cells[ci + 1]
                if not _is_emptyish_text(_cell_text(mid)):
                    continue
                # 优先贴近“角色/姓名”字段标记插入，避免落到 tab leader/占位末端导致距离过远
                if not _insert_sig_in_role_cell_for_adjacent_date_column(left, keyword, sig_png):
                    _insert_sig_only_in_empty_cell(mid, sig_png)
            if not _insert_date_only_in_date_cell(date_cell, date_png):
                continue
            return True
    return False


def _try_paragraph_inline(
    p: Paragraph,
    keyword: str,
    sig_png: bytes,
    date_png: bytes,
) -> bool:
    """段落内：关键词后占位符清除并插图，或下划线 run 后插图，否则关键词后追加。"""
    off = _find_keyword_in_paragraph(p, keyword)
    if off < 0:
        return False
    full = _paragraph_full_text(p)
    rest = full[off:]
    rest_stripped = rest.lstrip(" ：:\t")
    # 同段有「日期」且要先签后日期：用户按角色提供 sig+date，都插在同一角色关联区域
    head = (rest_stripped[:20] if len(rest_stripped) > 20 else rest_stripped) if rest_stripped else ""
    if _is_emptyish_text(rest_stripped) or (head and _PLACEHOLDER_CHARS.match(head)):
        _clear_runs_after_offset(p, off)
        _insert_pictures_in_paragraph(p, sig_png, date_png)
        return True
    # 尝试：从第一个 run 开始找关键词后的下划线 run
    acc = 0
    found_kw_end = False
    insert_after_run_idx = None
    for i, r in enumerate(p.runs):
        rt = r.text or ""
        seg = acc + len(rt)
        if not found_kw_end:
            # off 为关键词匹配结束位置（见 paragraph_text_keyword_end_offset）
            if acc < off <= seg:
                found_kw_end = True
        if found_kw_end and acc >= off and _run_has_underline_or_border(r):
            insert_after_run_idx = i
            break
        acc = seg
    if insert_after_run_idx is not None:
        # 在指定 run 后插入新 run（通过往段落添加；python-docx 只能 add_run 在末尾）
        # 故退化为：清除关键词后的占位并末尾追加图（避免破坏顺序时）
        _clear_runs_after_offset(p, off)
        _insert_pictures_in_paragraph(p, sig_png, date_png)
        return True
    # 关键词后已有非占位正文：避免误删，交其它段落/表格再匹配
    if rest_stripped and not _PLACEHOLDER_CHARS.match(rest_stripped):
        return False
    _clear_runs_after_offset(p, off)
    _insert_pictures_in_paragraph(p, sig_png, date_png)
    return True


def _try_table_role(
    table: Table,
    keyword: str,
    sig_png: bytes,
    date_png: bytes,
) -> bool:
    if _try_table_adjacent_date_column(table, keyword, sig_png, date_png):
        return True
    rows_list = list(table.rows)
    for ri, row in enumerate(rows_list):
        cells = row.cells
        for ci, cell in enumerate(cells):
            if not cell_text_matches_keyword(_cell_text(cell), keyword):
                continue
            if ci + 1 >= len(cells):
                continue
            sig_cell = cells[ci + 1]
            if not _is_emptyish_text(_cell_text(sig_cell)):
                if ci + 2 < len(cells) and _is_emptyish_text(_cell_text(cells[ci + 2])):
                    sig_cell = cells[ci + 2]
                else:
                    continue
            si = None
            for j in range(len(cells)):
                if cells[j]._tc == sig_cell._tc:
                    si = j
                    break
            date_cell = None
            if si is not None:
                if si + 1 < len(cells) and _is_emptyish_text(_cell_text(cells[si + 1])):
                    date_cell = cells[si + 1]
                elif si + 2 < len(cells) and _is_emptyish_text(_cell_text(cells[si + 2])):
                    date_cell = cells[si + 2]
            if date_cell is None and ri + 1 < len(rows_list):
                nrow = rows_list[ri + 1]
                nc = nrow.cells
                if ci + 1 < len(nc) and _is_emptyish_text(_cell_text(nc[ci + 1])):
                    date_cell = nc[ci + 1]
                elif ci < len(nc) and _is_emptyish_text(_cell_text(nc[ci])):
                    date_cell = nc[ci]
            sig_cell.text = ""
            p0 = sig_cell.paragraphs[0] if sig_cell.paragraphs else sig_cell.add_paragraph()
            for r in list(p0.runs):
                r.text = ""
            if date_cell is None:
                _insert_pictures_in_paragraph(p0, sig_png, date_png)
                return True
            p0.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
            date_cell.text = ""
            p2 = date_cell.paragraphs[0] if date_cell.paragraphs else date_cell.add_paragraph()
            for r in list(p2.runs):
                r.text = ""
            p2.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_SMALL)
            return True
    return False


def _iter_tables(doc: DocumentObject) -> list[Table]:
    out = []
    try:
        for t in doc.tables:
            out.append(t)
    except Exception:
        pass
    return out


def _iter_body_paragraphs(doc: DocumentObject) -> list[Paragraph]:
    return list(doc.paragraphs)


def sign_docx(
    path: str,
    role_to_signature_png: dict,
    role_to_date_png: dict,
    out_path: Optional[str] = None,
) -> str:
    path = os.path.abspath(path)
    if out_path is None:
        base, ext = os.path.splitext(path)
        out_path = f"{base}_signed{ext}"
    doc = Document(path)

    for role_id in ROLE_ID_TO_KEYWORD:
        sig = role_to_signature_png.get(role_id)
        dt = role_to_date_png.get(role_id)
        if not sig or not dt:
            continue
        done = False
        for kw in role_keywords(role_id):
            if done:
                break
            for table in _iter_tables(doc):
                if _try_table_role(table, kw, sig, dt):
                    done = True
                    break
            if done:
                break
            for p in _iter_body_paragraphs(doc):
                if _find_keyword_in_paragraph(p, kw) < 0:
                    continue
                if _try_paragraph_inline(p, kw, sig, dt):
                    done = True
                    break
            if done:
                break
            for table in _iter_tables(doc):
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if _find_keyword_in_paragraph(p, kw) < 0:
                                continue
                            if _try_paragraph_inline(p, kw, sig, dt):
                                done = True
                                break
                        if done:
                            break
                    if done:
                        break
                if done:
                    break

    doc.save(out_path)
    return out_path
