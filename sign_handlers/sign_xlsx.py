# -*- coding: utf-8 -*-
"""Excel .xlsx：在关键词单元格相邻空白格插入签名/日期 PNG。"""
from __future__ import annotations

import io
import os
import re
from typing import Optional

from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl import load_workbook

from sign_handlers.config import ROLE_ID_TO_KEYWORD, role_keywords
from sign_handlers.label_match import cell_text_matches_keyword, paragraph_text_keyword_end_offset
from sign_handlers.png_word_compat import (
    prepare_png_for_word,
    prepare_signature_date_pair_for_word,
)

# Excel 里 220px 太小，签字观感发虚；提高上限，同时仍按单元格宽高自适应缩放
_MAX_IMG_WIDTH_PX = 420
_DATE_LABEL_KEYWORDS = ("日期", "Date")
_PLACEHOLDER_TAIL = re.compile(r"^[\s_\-—–\u2014\u2015\u2500\u3000\.·…．~～]+$")
_GAP_PX = 6


def _col_width_px(ws, col_idx: int) -> int:
    """Excel 列宽到像素的近似换算（足够用于锚点/缩放）。"""
    try:
        letter = get_column_letter(int(col_idx))
        w = ws.column_dimensions[letter].width
        if w is None:
            w = ws.sheet_format.defaultColWidth
        if w is None:
            w = 8.43
        # 经验公式：1 字符宽约 7px + 5px padding
        return max(24, int(float(w) * 7.0 + 5.0))
    except Exception:
        return 64


def _row_height_px(ws, row_idx: int) -> int:
    """行高（pt）到像素近似换算。"""
    try:
        h = ws.row_dimensions[int(row_idx)].height
        if h is None:
            h = ws.sheet_format.defaultRowHeight
        if h is None:
            h = 15.0
        return max(18, int(float(h) * 96.0 / 72.0))
    except Exception:
        return 20


def _fit_to_box(w: int, h: int, max_w: int, max_h: int) -> tuple[int, int]:
    if w <= 0 or h <= 0:
        return max(1, max_w), max(1, max_h)
    s = min(max_w / float(w), max_h / float(h), 1.0)
    return max(1, int(round(w * s))), max(1, int(round(h * s)))


def _find_date_cell_same_row(ws, r: int, author_col: int):
    """同行查找「Date / 日期」标签格与其右侧空白格。"""
    max_col = min(ws.max_column or 0, 64)
    if max_col < author_col + 2:
        return None, None
    for col in range(author_col + 2, max_col + 1):
        cl = ws.cell(row=r, column=col)
        for dkw in _DATE_LABEL_KEYWORDS:
            if cell_text_matches_keyword(cl.value, dkw):
                dc = ws.cell(row=r, column=col + 1)
                if _is_emptyish(dc.value):
                    return cl, dc
                return cl, None
    return None, None


def _is_emptyish(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    if not s:
        return True
    if all(c in " _-—–\u2014\u2015\u2500\u3000.·" for c in s):
        return True
    return len(s) < 2


def _add_png(ws, png_bytes: bytes, anchor: str, max_w: int = _MAX_IMG_WIDTH_PX) -> None:
    img = XLImage(io.BytesIO(png_bytes))
    w = int(getattr(img, "width", None) or max_w)
    h = int(getattr(img, "height", None) or int(max_w * 0.35))
    w2, h2 = _fit_to_box(w, h, int(max_w), int(max_w * 0.55))
    img.width = w2
    img.height = h2
    img.anchor = anchor
    ws.add_image(img)


def _merged_top_left_coordinate(ws, cell) -> str:
    """
    若 cell 位于合并单元格内，返回该合并区域左上角坐标；否则返回 cell.coordinate。
    模板里常见「标题+点线留白」做成合并格，锚到右邻格会导致图片看起来不贴标题。
    """
    try:
        merged = getattr(ws, "merged_cells", None)
        if merged is None:
            return cell.coordinate
        for rng in merged.ranges:
            if cell.coordinate in rng:
                return f"{get_column_letter(rng.min_col)}{rng.min_row}"
    except Exception:
        pass
    return cell.coordinate


def _cell_inline_insert_offset_px(cell_value, keyword: str) -> Optional[int]:
    """
    若单元格是「Author:......」这类同格标签+占位，返回应紧跟标签后插图的 x 偏移（px）。
    """
    if cell_value is None:
        return None
    txt = str(cell_value)
    off = paragraph_text_keyword_end_offset(txt, keyword)
    if off < 0:
        return None
    tail = txt[off:]
    if tail and not _PLACEHOLDER_TAIL.match(tail):
        return None
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
    return min(max(px, 4), 300)


def _add_png_after_label_in_cell(ws, png_bytes: bytes, cell, offset_x_px: int, max_w: int = _MAX_IMG_WIDTH_PX) -> None:
    """把图片锚到当前标签单元格内偏移位置，避免落到占位点线末端，并按格宽高自适应缩放。"""
    img = XLImage(io.BytesIO(png_bytes))
    w = int(getattr(img, "width", None) or max_w)
    h = int(getattr(img, "height", None) or int(max_w * 0.35))
    cell_w = _col_width_px(ws, cell.column)
    cell_h = _row_height_px(ws, cell.row)
    avail_w = max(28, cell_w - int(offset_x_px) - 4)
    max_w2 = min(int(max_w), int(avail_w))
    max_h2 = max(14, int(cell_h) - 4)
    w2, h2 = _fit_to_box(w, h, max_w2, max_h2)
    img.width = w2
    img.height = h2
    marker = AnchorMarker(
        col=cell.column - 1,
        row=cell.row - 1,
        colOff=pixels_to_EMU(offset_x_px),
        rowOff=pixels_to_EMU(1),
    )
    ext = XDRPositiveSize2D(pixels_to_EMU(int(img.width)), pixels_to_EMU(int(img.height)))
    img.anchor = OneCellAnchor(_from=marker, ext=ext)
    ws.add_image(img)


def _add_sig_and_date_after_label_in_cell(
    ws,
    sig_png: bytes,
    date_png: bytes,
    cell,
    offset_x_px: int,
    max_w: int = _MAX_IMG_WIDTH_PX,
) -> None:
    """同一单元格内：标题后先签名紧接日期（两个图片各自锚在不同 colOff）。"""
    cell_w = _col_width_px(ws, cell.column)
    cell_h = _row_height_px(ws, cell.row)
    avail_w = max(40, cell_w - int(offset_x_px) - 4)
    max_h2 = max(14, int(cell_h) - 4)

    sig_img = XLImage(io.BytesIO(sig_png))
    s_w = int(getattr(sig_img, "width", None) or max_w)
    s_h = int(getattr(sig_img, "height", None) or int(max_w * 0.35))
    date_img = XLImage(io.BytesIO(date_png))
    d_w = int(getattr(date_img, "width", None) or max_w)
    d_h = int(getattr(date_img, "height", None) or int(max_w * 0.35))

    # 先给签名一个上限（优先留位置给日期）
    sig_box_w = max(24, min(int(max_w), int(avail_w * 0.62)))
    s_w2, s_h2 = _fit_to_box(s_w, s_h, sig_box_w, max_h2)
    remain = max(24, avail_w - s_w2 - _GAP_PX)
    d_box_w = max(24, min(int(max_w), int(remain)))
    d_w2, d_h2 = _fit_to_box(d_w, d_h, d_box_w, max_h2)

    sig_img.width = s_w2
    sig_img.height = s_h2
    date_img.width = d_w2
    date_img.height = d_h2

    m1 = AnchorMarker(
        col=cell.column - 1,
        row=cell.row - 1,
        colOff=pixels_to_EMU(int(offset_x_px)),
        rowOff=pixels_to_EMU(1),
    )
    e1 = XDRPositiveSize2D(pixels_to_EMU(int(sig_img.width)), pixels_to_EMU(int(sig_img.height)))
    sig_img.anchor = OneCellAnchor(_from=m1, ext=e1)
    ws.add_image(sig_img)

    m2 = AnchorMarker(
        col=cell.column - 1,
        row=cell.row - 1,
        colOff=pixels_to_EMU(int(offset_x_px) + int(sig_img.width) + _GAP_PX),
        rowOff=pixels_to_EMU(1),
    )
    e2 = XDRPositiveSize2D(pixels_to_EMU(int(date_img.width)), pixels_to_EMU(int(date_img.height)))
    date_img.anchor = OneCellAnchor(_from=m2, ext=e2)
    ws.add_image(date_img)


def sign_xlsx(
    path: str,
    role_to_signature_png: dict,
    role_to_date_png: dict,
    out_path: Optional[str] = None,
) -> str:
    path = os.path.abspath(path)
    if out_path is None:
        base, ext = os.path.splitext(path)
        out_path = f"{base}_signed{ext}"

    wb = load_workbook(path)
    for role_id in ROLE_ID_TO_KEYWORD:
        sb = role_to_signature_png.get(role_id)
        db = role_to_date_png.get(role_id)
        if sb and db:
            sig, dt = prepare_signature_date_pair_for_word(sb, db)
        else:
            sig = prepare_png_for_word(sb) if sb else None
            dt = prepare_png_for_word(db) if db else None
            if sig is None and sb:
                sig = sb
            if dt is None and db:
                dt = db
        if not sig and not dt:
            continue
        placed = False
        for kw in role_keywords(role_id):
            if placed:
                break
            for ws in wb.worksheets:
                if placed:
                    break
                for row in ws.iter_rows():
                    if placed:
                        break
                    for cell in row:
                        if not cell_text_matches_keyword(cell.value, kw):
                            continue
                        r, c = cell.row, cell.column
                        sig_cell = ws.cell(row=r, column=c + 1)
                        # 若要贴签名，签名格必须为空；若只贴日期，则不强制要求签名格为空
                        if sig and (not _is_emptyish(sig_cell.value)):
                            continue
                        date_label_cell, date_cell = _find_date_cell_same_row(ws, r, c)
                        if date_cell is None:
                            c2 = ws.cell(row=r, column=c + 2)
                            if _is_emptyish(c2.value):
                                date_cell = c2
                            else:
                                d1 = ws.cell(row=r + 1, column=c + 1)
                                if _is_emptyish(d1.value):
                                    date_cell = d1
                                else:
                                    d0 = ws.cell(row=r + 1, column=c)
                                    if _is_emptyish(d0.value):
                                        date_cell = d0
                        placed_any = False
                        if sig:
                            # 优先：字段后紧接的空白单元格（按用户期望“插在字段后的单元格里”）
                            if _is_emptyish(sig_cell.value):
                                sig_cell.value = None
                                _add_png(ws, sig, _merged_top_left_coordinate(ws, sig_cell))
                            else:
                                # 仅当“同一单元格内留空占位”且没有可用的后续空白格时，才退化为同格插入
                                sig_inline_x = _cell_inline_insert_offset_px(cell.value, kw)
                                if sig_inline_x is not None:
                                    _add_png_after_label_in_cell(ws, sig, cell, sig_inline_x)
                            placed_any = True
                        if dt:
                            # 优先：Date/日期 标签右侧空白格；其次：启发式找到的 date_cell
                            if date_cell is not None and _is_emptyish(date_cell.value):
                                date_cell.value = None
                                _add_png(ws, dt, _merged_top_left_coordinate(ws, date_cell))
                                placed_any = True
                            else:
                                # 若 Date/日期 标签格本身是“同格留空占位”，且没有可用的右侧空白格，才退化为同格插入
                                date_inline_x = None
                                if date_label_cell is not None:
                                    date_inline_x = _cell_inline_insert_offset_px(date_label_cell.value, "Date")
                                    if date_inline_x is None:
                                        date_inline_x = _cell_inline_insert_offset_px(date_label_cell.value, "日期")
                                if date_label_cell is not None and date_inline_x is not None:
                                    _add_png_after_label_in_cell(ws, dt, date_label_cell, date_inline_x)
                                    placed_any = True
                                elif date_cell is not None:
                                    # 非空也允许覆盖（用户明确选择了日期素材）
                                    date_cell.value = None
                                    _add_png(ws, dt, _merged_top_left_coordinate(ws, date_cell))
                                    placed_any = True
                                else:
                                    # 找不到日期单元格时：放到签名格下方（兜底）
                                    below = f"{get_column_letter(c + 1)}{r + 1}"
                                    _add_png(ws, dt, below)
                                    placed_any = True
                        if not placed_any:
                            continue
                        placed = True
                        break
    wb.save(out_path)
    return out_path
