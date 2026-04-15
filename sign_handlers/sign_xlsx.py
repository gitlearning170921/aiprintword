# -*- coding: utf-8 -*-
"""Excel .xlsx：在关键词单元格相邻空白格插入签名/日期 PNG。"""
from __future__ import annotations

import io
import os
from typing import Optional

from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

from sign_handlers.config import ROLE_ID_TO_KEYWORD, role_keywords
from sign_handlers.label_match import cell_text_matches_keyword

_MAX_IMG_WIDTH_PX = 220
_DATE_LABEL_KEYWORDS = ("日期", "Date")


def _find_date_cell_same_row(ws, r: int, author_col: int):
    """同行在「Date / 日期」标签右侧空白格放日期（如 Author 与 Date 分列的封面表）。"""
    max_col = min(ws.max_column or 0, 64)
    if max_col < author_col + 2:
        return None
    for col in range(author_col + 2, max_col + 1):
        cl = ws.cell(row=r, column=col)
        for dkw in _DATE_LABEL_KEYWORDS:
            if cell_text_matches_keyword(cl.value, dkw):
                dc = ws.cell(row=r, column=col + 1)
                if _is_emptyish(dc.value):
                    return dc
    return None


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
    w = getattr(img, "width", None) or max_w
    img.width = min(int(w), max_w)
    img.anchor = anchor
    ws.add_image(img)


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
        sig = role_to_signature_png.get(role_id)
        dt = role_to_date_png.get(role_id)
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
                        date_cell = _find_date_cell_same_row(ws, r, c)
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
                            sig_cell.value = None
                            _add_png(ws, sig, sig_cell.coordinate)
                            placed_any = True
                        if dt:
                            if date_cell is not None:
                                date_cell.value = None
                                _add_png(ws, dt, date_cell.coordinate, max_w=180)
                            else:
                                # 找不到日期单元格时：放到签名格下方
                                below = f"{get_column_letter(c + 1)}{r + 1}"
                                _add_png(ws, dt, below, max_w=180)
                            placed_any = True
                        if not placed_any:
                            continue
                        placed = True
                        break
    wb.save(out_path)
    return out_path
