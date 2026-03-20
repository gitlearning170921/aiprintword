# -*- coding: utf-8 -*-
"""Excel .xlsx：在关键词单元格相邻空白格插入签名/日期 PNG。"""
from __future__ import annotations

import io
import os
from typing import Optional

from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

from sign_handlers.config import ROLE_ID_TO_KEYWORD

_MAX_IMG_WIDTH_PX = 220


def _is_emptyish(v) -> bool:
    if v is None:
        return True
    s = str(v).strip()
    if not s:
        return True
    if all(c in " _-—–\u2014\u2015\u2500\u3000.·" for c in s):
        return True
    return len(s) < 2


def _cell_contains_keyword(cell_value, keyword: str) -> bool:
    if cell_value is None:
        return False
    s = str(cell_value).strip()
    if s == keyword or s == f"{keyword}：" or s == f"{keyword}:":
        return True
    if s.startswith(keyword) and len(s) <= len(keyword) + 2:
        return True
    return False


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
    for role_id, kw in ROLE_ID_TO_KEYWORD.items():
        sig = role_to_signature_png.get(role_id)
        dt = role_to_date_png.get(role_id)
        if not sig or not dt:
            continue
        placed = False
        for ws in wb.worksheets:
            if placed:
                break
            for row in ws.iter_rows():
                if placed:
                    break
                for cell in row:
                    if not _cell_contains_keyword(cell.value, kw):
                        continue
                    r, c = cell.row, cell.column
                    sig_cell = ws.cell(row=r, column=c + 1)
                    if not _is_emptyish(sig_cell.value):
                        continue
                    date_cell = None
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
                    sig_cell.value = None
                    _add_png(ws, sig, sig_cell.coordinate)
                    if date_cell is not None:
                        date_cell.value = None
                        _add_png(ws, dt, date_cell.coordinate, max_w=180)
                    else:
                        below = f"{get_column_letter(c + 1)}{r + 1}"
                        _add_png(ws, dt, below, max_w=180)
                    placed = True
                    break
    wb.save(out_path)
    return out_path
