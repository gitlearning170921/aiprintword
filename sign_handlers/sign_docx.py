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
from sign_handlers.png_word_compat import (
    prepare_png_for_word,
    prepare_signature_date_pair_for_word,
)

# 可见占位字符（作者：____ 等）
_PLACEHOLDER_CHARS = re.compile(r"^[\s_\-—–\u2014\u2015\u2500\u3000\.·]+$")
# 右侧「日期」列表头（整格以 Date/日期 开头）
_DATE_HEADER_CELL = re.compile(r"^\s*(日期|Date)\s*[:：]?\s*", re.IGNORECASE)

_PIC_WIDTH = Cm(2.8)
# 日期与签名使用同一显示宽度，避免缩放后笔画变细、观感不一致
_PIC_WIDTH_DATE = _PIC_WIDTH


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
    # 勿用 strip() 判断「无内容」：关键词后常见仅 Tab/全角空格/制表符前导点线，
    # strip 后为空会误判并跳过删除，导致后续 add_run 把图片追加到段末（离标签很远、易换行）。
    if not full[start_char_offset:]:
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


def _truncate_after_offset_and_get_anchor_run_idx(
    p: Paragraph, start_char_offset: int
) -> int:
    """
    从 start_char_offset 起清空文本，并返回“插入锚点 run”的索引：
    - 若 offset 落在某个 run 内：该 run 被截断到 offset，返回该 run idx
    - 若 offset 正好在 run 边界：返回 offset 前的 run idx（最靠近标题的那个）
    - 若无法定位：返回最后一个 run idx（或 -1 表示无 run）
    """
    full = _paragraph_full_text(p)
    if not p.runs:
        return -1
    if start_char_offset <= 0:
        # 标题在段首：锚点设为第一个 run（图片插在它后）
        return 0
    if start_char_offset >= len(full):
        return len(p.runs) - 1

    def _run_is_placeholder(r) -> bool:
        rt = r.text or ""
        # 下划线/边框 run 通常是“__ / 点线”占位；直接当作占位清空
        if _run_has_underline_or_border(r):
            return True
        # 仅包含空白/下划线/横线/点线等字符：当作占位清空
        try:
            if _PLACEHOLDER_CHARS.match(rt):
                return True
        except Exception:
            pass
        # 兜底：全空白也当作占位
        if not rt or not str(rt).strip():
            return True
        return False

    acc = 0
    last_nonempty_idx = 0
    for i, r in enumerate(list(p.runs)):
        rt = r.text or ""
        ln = len(rt)
        if rt:
            last_nonempty_idx = i
        seg = acc + ln
        if seg < start_char_offset:
            acc = seg
            continue
        if seg == start_char_offset:
            # offset 恰好在 run 边界：只清空“占位类”的后续 run
            for rr in list(p.runs)[i + 1 :]:
                if _run_is_placeholder(rr):
                    rr.text = ""
            return i
        # acc < offset < seg：截断当前 run 并清空后续
        cut = start_char_offset - acc
        r.text = rt[:cut]
        for rr in list(p.runs)[i + 1 :]:
            if _run_is_placeholder(rr):
                rr.text = ""
        return i
    # fallback：按原逻辑清空，锚点取最后一个含文本的 run
    _clear_runs_after_offset(p, start_char_offset)
    return last_nonempty_idx


def _move_run_after_run_idx(p: Paragraph, new_run, anchor_run_idx: int) -> None:
    """
    python-docx 只能在段末 add_run；此处把 new_run 的 XML 节点移动到 anchor_run_idx 后面，
    以保证图片紧接字段标题后（不会跑到点线/空格末尾导致换行）。
    """
    try:
        if anchor_run_idx < 0:
            return
        runs = list(p.runs)
        if not runs:
            return
        if anchor_run_idx >= len(runs):
            anchor_run_idx = len(runs) - 1
        anchor_r = runs[anchor_run_idx]._r
        new_r = new_run._r
        # 先从末尾移除，再插入到 anchor 后
        p._p.remove(new_r)
        insert_pos = p._p.index(anchor_r) + 1
        p._p.insert(insert_pos, new_r)
    except Exception:
        # 若低层 XML 操作失败，退化为段末追加（不致使签署失败）
        return


def _insert_pictures_in_paragraph(
    p: Paragraph, sig_png: Optional[bytes], date_png: Optional[bytes]
) -> None:
    """在同一段落中按需插入签名/日期（可只插入其一）。默认追加到段末。"""
    if sig_png:
        p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
    if sig_png and date_png:
        p.add_run(" ")
    if date_png:
        p.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)


def _insert_pictures_after_anchor_run_idx(
    p: Paragraph,
    anchor_run_idx: int,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
) -> None:
    """
    在 anchor_run_idx 之后插入签名/日期图片，并把对应 run 移动到锚点后，
    以保证紧贴字段标题，不会跑到点线/空格末尾导致换行。
    """
    runs_to_place = []
    if sig_png:
        rs = p.add_run()
        rs.add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
        runs_to_place.append(rs)
    if sig_png and date_png:
        runs_to_place.append(p.add_run(" "))
    if date_png:
        rd = p.add_run()
        rd.add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)
        runs_to_place.append(rd)

    # 逐个移动：每次都插到“当前锚点后”，并推进锚点以保持顺序
    cur_anchor = anchor_run_idx
    for r in runs_to_place:
        _move_run_after_run_idx(p, r, cur_anchor)
        cur_anchor += 1


def _cell_looks_like_date_header_cell(text: str) -> bool:
    t = (text or "").strip()
    return bool(_DATE_HEADER_CELL.match(t))


def _find_date_label_end_in_paragraph(p: Paragraph) -> int:
    full = _paragraph_full_text(p)
    if not full:
        return -1
    m = re.search(r"(?i)(?:\bDate\b|日期)\s*[:：]?", full)
    return m.end() if m else -1


def _insert_sig_only_at_char_offset(p: Paragraph, char_offset: int, sig_png: bytes) -> None:
    # 紧接「Author:/Approver:」等标题后插入：从插入点起清空其后文本/点线占位，
    # 再把图片 run 移到标题 run 后（避免落在点线/空格末尾导致换行）。
    anchor_idx = _truncate_after_offset_and_get_anchor_run_idx(p, char_offset)
    rn = p.add_run()
    rn.add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
    _move_run_after_run_idx(p, rn, anchor_idx)


def _insert_date_only_in_date_cell(date_cell, date_png: bytes) -> bool:
    """在「日期/Date」标签后的空白处只插入日期手写图。"""
    for p in date_cell.paragraphs:
        offd = _find_date_label_end_in_paragraph(p)
        if offd < 0:
            continue
        anchor_idx = _truncate_after_offset_and_get_anchor_run_idx(p, offd)
        rn = p.add_run()
        rn.add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)
        _move_run_after_run_idx(p, rn, anchor_idx)
        return True
    if date_cell.paragraphs:
        p0 = date_cell.paragraphs[0]
        for r in list(p0.runs):
            r.text = ""
        rn = p0.add_run()
        rn.add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)
        return True
    p = date_cell.add_paragraph()
    rn = p.add_run()
    rn.add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)
    return True


def _insert_sig_in_role_cell_for_adjacent_date_column(
    role_cell, keyword: str, sig_png: bytes
) -> bool:
    for p in role_cell.paragraphs:
        off = _find_keyword_in_paragraph(p, keyword)
        if off < 0:
            continue
        _insert_sig_only_at_char_offset(p, off, sig_png)
        return True
    return False


def _insert_sig_only_in_empty_cell(cell, sig_png: bytes) -> None:
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    for r in list(p.runs):
        r.text = ""
    p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)


def _insert_sig_and_date_in_empty_cell(cell, sig_png: Optional[bytes], date_png: Optional[bytes]) -> None:
    """把签名/日期插入到“预留空白格”中（清空原占位）。"""
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    for r in list(p.runs):
        r.text = ""
    if sig_png:
        p.add_run().add_picture(io.BytesIO(sig_png), width=_PIC_WIDTH)
    if sig_png and date_png:
        p.add_run(" ")
    if date_png:
        p.add_run().add_picture(io.BytesIO(date_png), width=_PIC_WIDTH_DATE)


def _next_distinct_cell(cells, idx: int):
    """返回同一行里 idx 之后第一个“不同的 cell”（按合并单元格去重）。"""
    try:
        base = cells[idx]
        base_tc = getattr(base, "_tc", None)
        for j in range(idx + 1, len(cells)):
            c2 = cells[j]
            if getattr(c2, "_tc", None) is not None and getattr(c2, "_tc", None) == base_tc:
                continue
            return c2
    except Exception:
        pass
    return None


def _try_table_adjacent_date_column(
    table: Table,
    keyword: str,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
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
            placed_any = False
            if sig_png:
                if date_idx == ci + 1:
                    # 只有两列且右列就是 Date 标签：签名只能写在左格关键词后
                    if not _insert_sig_in_role_cell_for_adjacent_date_column(left, keyword, sig_png):
                        continue
                else:
                    # 三列（或更多）：字段格后面通常有“预留空白签字格”，优先插到该空白格里
                    mid = _next_distinct_cell(cells, ci)
                    if mid is None or mid._tc == date_cell._tc:
                        # 没有独立签字格：退回写在字段格关键词后
                        if not _insert_sig_in_role_cell_for_adjacent_date_column(left, keyword, sig_png):
                            continue
                    else:
                        if not _is_emptyish_text(_cell_text(mid)):
                            continue
                        _insert_sig_only_in_empty_cell(mid, sig_png)
                placed_any = True
            if date_png:
                # 日期：优先插到 Date 标签格的右侧空白格；否则退回 Date 标签格内
                date_target = None
                try:
                    date_target = _next_distinct_cell(cells, date_idx)
                except Exception:
                    date_target = None
                if date_target is not None and _is_emptyish_text(_cell_text(date_target)):
                    _insert_sig_and_date_in_empty_cell(date_target, None, date_png)
                    placed_any = True
                elif not _insert_date_only_in_date_cell(date_cell, date_png):
                    # 仅有日期时若没能插入（极少），继续找其它匹配
                    if not placed_any:
                        continue
                placed_any = True
            if placed_any:
                return True
    return False


def _try_paragraph_inline(
    p: Paragraph,
    keyword: str,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
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
        anchor_idx = _truncate_after_offset_and_get_anchor_run_idx(p, off)
        _insert_pictures_after_anchor_run_idx(p, anchor_idx, sig_png, date_png)
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
        # 在关键词后（通常为冒号后）清空占位，并把图片 run 移动到标题 run 后
        anchor_idx = _truncate_after_offset_and_get_anchor_run_idx(p, off)
        _insert_pictures_after_anchor_run_idx(p, anchor_idx, sig_png, date_png)
        return True
    # 关键词后已有非占位正文：避免误删，交其它段落/表格再匹配
    if rest_stripped and not _PLACEHOLDER_CHARS.match(rest_stripped):
        return False
    anchor_idx = _truncate_after_offset_and_get_anchor_run_idx(p, off)
    _insert_pictures_after_anchor_run_idx(p, anchor_idx, sig_png, date_png)
    return True


def _try_table_role(
    table: Table,
    keyword: str,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
) -> bool:
    if _try_table_adjacent_date_column(table, keyword, sig_png, date_png):
        return True
    rows_list = list(table.rows)
    for ri, row in enumerate(rows_list):
        cells = row.cells
        for ci, cell in enumerate(cells):
            if not cell_text_matches_keyword(_cell_text(cell), keyword):
                continue
            sig_cell = _next_distinct_cell(cells, ci)
            # 优先找同一行 Date/日期 标签格，保证日期贴在字段标题后，而非整段点线末端
            date_cell = None
            for j in range(ci + 1, len(cells)):
                if _cell_looks_like_date_header_cell(_cell_text(cells[j])):
                    date_cell = cells[j]
                    break
            if date_cell is None and ri + 1 < len(rows_list):
                # 次选：下一行同列/右邻列若出现 Date/日期 标签格
                nrow = rows_list[ri + 1]
                for cand_col in (ci, ci + 1, ci + 2):
                    if cand_col < len(nrow.cells) and _cell_looks_like_date_header_cell(_cell_text(nrow.cells[cand_col])):
                        date_cell = nrow.cells[cand_col]
                        break
            placed_any = False
            if sig_png:
                # 优先写在“字段后紧接的预留空白格”（按合并单元格后的 next distinct cell）
                if sig_cell is not None and _is_emptyish_text(_cell_text(sig_cell)) and (date_cell is None or sig_cell._tc != date_cell._tc):
                    _insert_sig_and_date_in_empty_cell(sig_cell, sig_png, None)
                    placed_any = True
                else:
                    # 无预留空白格时才写回字段格关键词后（保持正文逻辑不变）
                    if _insert_sig_in_role_cell_for_adjacent_date_column(cell, keyword, sig_png):
                        placed_any = True
            if date_png and date_cell is not None:
                # 日期：优先写到 Date 标签格右侧空白格
                dt_target = _next_distinct_cell(cells, cells.index(date_cell)) if date_cell in cells else None
                if dt_target is not None and _is_emptyish_text(_cell_text(dt_target)):
                    _insert_sig_and_date_in_empty_cell(dt_target, None, date_png)
                    placed_any = True
                elif _insert_date_only_in_date_cell(date_cell, date_png):
                    placed_any = True
            if date_png and date_cell is None and (not sig_png):
                # 只有日期但没找到 Date 标签时，退化：尝试该角色格关键词后插入
                for p0 in cell.paragraphs:
                    off = _find_keyword_in_paragraph(p0, keyword)
                    if off >= 0:
                        _clear_runs_after_offset(p0, off)
                        _insert_pictures_in_paragraph(p0, None, date_png)
                        placed_any = True
                        break
            # 若标签为 “Reviewer/Date” 这类同格字段且后面有预留空白格：把签名+日期都插到预留格里
            if (sig_png or date_png) and (date_cell is None) and sig_cell is not None and _is_emptyish_text(_cell_text(sig_cell)):
                _insert_sig_and_date_in_empty_cell(sig_cell, sig_png, date_png)
                placed_any = True
            if placed_any:
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

    sig_map: dict = {}
    date_map: dict = {}
    for rid in ROLE_ID_TO_KEYWORD:
        sb = role_to_signature_png.get(rid)
        db = role_to_date_png.get(rid)
        if sb and db:
            s2, d2 = prepare_signature_date_pair_for_word(sb, db)
            if s2:
                sig_map[rid] = s2
            if d2:
                date_map[rid] = d2
        else:
            if sb:
                sig_map[rid] = prepare_png_for_word(sb) or sb
            if db:
                date_map[rid] = prepare_png_for_word(db) or db

    for role_id in ROLE_ID_TO_KEYWORD:
        sig = sig_map.get(role_id)
        dt = date_map.get(role_id)
        if not sig and not dt:
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
