# -*- coding: utf-8 -*-
"""导出与自动签字落位一致缩放比例的 PNG，供工作台手动贴图。"""
from __future__ import annotations

import io
from typing import Any, Dict, Optional, Tuple

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None  # type: ignore

from sign_handlers.png_word_compat import (
    prepare_png_for_word,
    prepare_signature_date_pair_for_word,
)

# 与 sign_docx._PIC_WIDTH 一致（Cm(2.8)）
DOCX_INSERT_WIDTH_CM = 2.8
# Word 粘贴/显示常用 DPI；与 Cm(2.8) 对应的像素宽度约 106px
DOCX_INSERT_DPI = 96


def scale_png_for_docx_insert(
    png_bytes: Optional[bytes],
    *,
    width_cm: float = DOCX_INSERT_WIDTH_CM,
    dpi: int = DOCX_INSERT_DPI,
) -> Optional[bytes]:
    """将 PNG 缩放到与 Word add_picture(width=Cm(2.8)) 一致的显示像素尺寸。"""
    if not png_bytes or Image is None:
        return png_bytes
    try:
        im = Image.open(io.BytesIO(png_bytes))
    except Exception:
        return png_bytes
    w, h = im.size
    if w < 1 or h < 1:
        return png_bytes
    target_w = max(1, int(round(float(width_cm) / 2.54 * float(dpi))))
    target_h = max(1, int(round(h * (target_w / float(w)))))
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    flat = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    flat.alpha_composite(im)
    rgb = flat.convert("RGB").resize((target_w, target_h), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    rgb.save(out, format="PNG", dpi=(dpi, dpi), optimize=True)
    return out.getvalue()


def scale_png_to_pixel_box(
    png_bytes: Optional[bytes],
    target_w: int,
    target_h: int,
) -> Optional[bytes]:
    """等比缩放 PNG 到目标像素框（与 Excel 落位 _fit_to_box_excel 结果一致）。"""
    if not png_bytes or Image is None:
        return png_bytes
    tw = max(1, int(target_w))
    th = max(1, int(target_h))
    try:
        im = Image.open(io.BytesIO(png_bytes))
    except Exception:
        return png_bytes
    w, h = im.size
    if w < 1 or h < 1:
        return png_bytes
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    flat = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    flat.alpha_composite(im)
    rgb = flat.convert("RGB").resize((tw, th), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    rgb.save(out, format="PNG", optimize=True)
    return out.getvalue()


def _resolve_pair_material_bytes(
    pair: dict,
    *,
    using_mysql: bool,
    mysql_store,
    local_get_item,
    inbox_root: str,
    sid: str,
) -> Tuple[Optional[bytes], Optional[bytes], Optional[str]]:
    """解析 role-map 条目为 (sig_png, date_png, error)。"""
    if not isinstance(pair, dict):
        return None, None, "角色素材未配置"
    sig_id = (pair.get("sig") or "").strip() or None
    date_id = (pair.get("date") or "").strip() or None
    dm = pair.get("date_mode")
    diso = (pair.get("date_iso") or "").strip() or None

    def _load_item(iid: str) -> Optional[bytes]:
        if not iid:
            return None
        if using_mysql:
            row = mysql_store.get_stroke_item_row(iid)
            return (row or {}).get("png")
        return local_get_item(inbox_root, sid, iid)

    sig_png = _load_item(sig_id) if sig_id else None
    date_png = None
    if using_mysql and mysql_store and mysql_store.is_composite_date_mode(dm):
        if not sig_id or not diso:
            return sig_png, None, "拼接日期需绑定签名并填写日历日期"
        srow = mysql_store.get_stroke_item_row(sig_id) if sig_id else None
        sid0 = ((srow or {}).get("signer_id") or "").strip()
        if not sid0:
            return sig_png, None, "无法解析签署人（拼接日期）"
        try:
            lay = mysql_store.composite_mode_to_layout(dm)
            date_png, _ = mysql_store.compose_date_piece_png(sid0, diso, lay)
        except Exception as e:
            return sig_png, None, f"日期拼接失败：{e}"
    elif date_id:
        date_png = _load_item(date_id)
    return sig_png, date_png, None


def _preprocess_role_pair(
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
    *,
    ext: str,
) -> Tuple[Optional[bytes], Optional[bytes]]:
    ext_l = (ext or "").lower()
    if sig_png and date_png:
        if ext_l in (".xlsx", ".xls"):
            return prepare_signature_date_pair_for_word(
                sig_png, date_png, equalize_ink_scale=False
            )
        return prepare_signature_date_pair_for_word(sig_png, date_png)
    sig = prepare_png_for_word(sig_png) if sig_png else None
    dt = prepare_png_for_word(date_png) if date_png else None
    if sig is None and sig_png:
        sig = sig_png
    if dt is None and date_png:
        dt = date_png
    return sig, dt


def _resolve_xlsx_target_cell(wb, role_plan: dict, kind: str):
    from sign_handlers.sign_xlsx import _is_emptyish, _xlsx_row_skip_for_signoff

    if kind == "sig":
        spec = role_plan.get("xlsx_name_cell") or role_plan.get("name_cell")
    else:
        spec = role_plan.get("xlsx_date_cell") or role_plan.get("date_cell")
    if isinstance(spec, dict):
        sno = int(spec.get("sheet") or spec.get("table") or 1)
        rno = int(spec.get("row") or 0)
        cno = int(spec.get("col") or 0)
        if rno >= 1 and cno >= 1:
            if sno < 1 or sno > len(wb.worksheets):
                ws = wb.worksheets[0]
            else:
                ws = wb.worksheets[sno - 1]
            if _xlsx_row_skip_for_signoff(ws, rno):
                pass
            else:
                right_c = cno + 1
                if right_c <= int(ws.max_column or 1):
                    right = ws.cell(row=rno, column=right_c)
                    if _is_emptyish(right.value):
                        return ws, right, 0
                return ws, ws.cell(row=rno, column=cno), 0

    hints = role_plan.get("source_hints") if isinstance(role_plan.get("source_hints"), list) else []
    from sign_handlers.sign_xlsx import _parse_xlsx_row_hint

    for hint in hints:
        parsed = _parse_xlsx_row_hint(str(hint or ""))
        if not parsed:
            continue
        sno, rno = parsed
        ws_list = list(wb.worksheets) if sno is None else [wb.worksheets[sno - 1]]
        for ws in ws_list:
            if rno < 1 or rno > int(ws.max_row or 1):
                continue
            if _xlsx_row_skip_for_signoff(ws, rno):
                continue
            max_c = min(int(ws.max_column or 1), 128)
            for cno in range(1, max_c + 1):
                right = ws.cell(row=rno, column=cno + 1) if cno + 1 <= max_c else None
                if right is not None and _is_emptyish(right.value):
                    return ws, right, 0
                if _is_emptyish(ws.cell(row=rno, column=cno).value):
                    return ws, ws.cell(row=rno, column=cno), 0
    return None, None, 0


def _default_xlsx_insert_box(wb) -> Tuple[int, int]:
    from sign_handlers.sign_xlsx import (
        _EXCEL_ABS_MAX_IMG_WIDTH_PX,
        _excel_content_max_width_px,
        _iter_xlsx_cover_worksheets,
        _merged_span_box_px,
        _xlsx_row_has_cover_signoff_labels,
        _xlsx_row_skip_for_signoff,
    )

    ws_candidates = _iter_xlsx_cover_worksheets(wb)
    if not ws_candidates:
        ws_candidates = [wb.active if wb.active is not None else wb.worksheets[0]]
    for ws in ws_candidates:
        max_c = min(int(ws.max_column or 1), 128)
        for rr in range(1, min(int(ws.max_row or 1), 64) + 1):
            if _xlsx_row_has_cover_signoff_labels(ws, rr, max_c=max_c):
                cell = ws.cell(row=rr, column=min(4, max_c))
                box_w, box_h = _merged_span_box_px(ws, cell)
                max_w2 = min(int(_EXCEL_ABS_MAX_IMG_WIDTH_PX), _excel_content_max_width_px(int(box_w)))
                max_h2 = max(14, int(box_h) - 4)
                return max(24, max_w2), max_h2
    ws = ws_candidates[0]
    r_end = int(ws.max_row or 1)
    target_r = r_end
    for rr in range(r_end, 0, -1):
        if not _xlsx_row_skip_for_signoff(ws, rr):
            target_r = rr
            break
    cell = ws.cell(row=target_r, column=min(3, int(ws.max_column or 1) or 1))
    box_w, box_h = _merged_span_box_px(ws, cell)
    max_w2 = min(int(_EXCEL_ABS_MAX_IMG_WIDTH_PX), _excel_content_max_width_px(int(box_w)))
    max_h2 = max(14, int(box_h) - 4)
    return max(24, max_w2), max_h2


def _scale_for_xlsx_insert(
    png_bytes: bytes,
    file_bytes: bytes,
    role_id: str,
    kind: str,
    placement_plan: Optional[dict],
) -> bytes:
    from openpyxl import load_workbook

    from sign_handlers.sign_xlsx import compute_xlsx_insert_pixel_size

    bio = io.BytesIO(file_bytes)
    wb = load_workbook(bio, data_only=True)
    try:
        rp = (placement_plan or {}).get(str(role_id)) if isinstance(placement_plan, dict) else {}
        rp = rp if isinstance(rp, dict) else {}
        ws, cell, inline_x = _resolve_xlsx_target_cell(wb, rp, kind)
        if ws is not None and cell is not None:
            tw, th = compute_xlsx_insert_pixel_size(
                png_bytes, ws, cell, inline_offset_px=int(inline_x or 0)
            )
        else:
            tw, th = _default_xlsx_insert_box(wb)
        scaled = scale_png_to_pixel_box(png_bytes, tw, th)
        return scaled or png_bytes
    finally:
        try:
            wb.close()
        except Exception:
            pass


def export_role_material_png(
    *,
    file_bytes: bytes,
    ext: str,
    role_id: str,
    kind: str,
    pair: dict,
    placement_plan: Optional[dict] = None,
    using_mysql: bool,
    mysql_store=None,
    local_get_item=None,
    inbox_root: str = "",
    sid: str = "",
) -> Tuple[Optional[bytes], Optional[str]]:
    """
    返回 (png_bytes, error)。
    kind: sig | date
    """
    kind_l = str(kind or "").strip().lower()
    if kind_l not in ("sig", "date"):
        return None, "kind 须为 sig 或 date"

    sig_raw, date_raw, err = _resolve_pair_material_bytes(
        pair,
        using_mysql=using_mysql,
        mysql_store=mysql_store,
        local_get_item=local_get_item,
        inbox_root=inbox_root,
        sid=sid,
    )
    if err and kind_l == "date" and not date_raw:
        return None, err

    sig_prep, date_prep = _preprocess_role_pair(sig_raw, date_raw, ext=ext)
    png = sig_prep if kind_l == "sig" else date_prep
    if not png:
        if kind_l == "sig":
            return None, "签名素材未选择或不可用"
        return None, err or "日期素材未选择或不可用"

    ext_l = (ext or "").lower()
    if ext_l in (".docx", ".doc"):
        out = scale_png_for_docx_insert(png)
        return out or png, None
    if ext_l in (".xlsx", ".xls"):
        out = _scale_for_xlsx_insert(png, file_bytes, role_id, kind_l, placement_plan)
        return out, None
    return png, None
