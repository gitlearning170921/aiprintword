# -*- coding: utf-8 -*-
"""PNG 预处理：让 Word/WPS 插入时的墨色更接近浏览器预览。

透明背景 + 抗锯齿在 Word 里叠在表格底纹上时，边缘容易发灰；日期条往往留白多，观感比签名更「淡」。
入库历史若存在 900px 宽日期而签名为 1200px，同版式宽度缩放时笔画粗细也会不一致。
"""
from __future__ import annotations

import io
import math
from typing import Optional

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None  # type: ignore

# 与前端 `sign-page.js` 中 `_normalizedPngDataUrl` 的目标宽度对齐
_TARGET_EXPORT_W = 1200
# 视为「白纸」的像素：RGB 之和小于此值才参与墨迹统计/加亮（255+255+255=765）
_INK_BG_SUM_MAX = 740


def _png_to_flat_rgb_image(png_bytes: Optional[bytes]):
    """透明底展平到白底，并把过窄图拉到统一像素宽度（与单图入库逻辑一致）。"""
    if not png_bytes or Image is None:
        return None
    try:
        im = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
    except Exception:
        return None
    w, h = im.size
    if w < 2 or h < 2:
        return None
    if w < _TARGET_EXPORT_W:
        nh = max(1, int(round(h * (_TARGET_EXPORT_W / float(w)))))
        im = im.resize((_TARGET_EXPORT_W, nh), Image.Resampling.LANCZOS)
        w, h = im.size
    flat = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    flat.alpha_composite(im, (0, 0))
    return flat.convert("RGB")


def _rgb_to_png_bytes(im) -> bytes:
    out = io.BytesIO()
    im.save(out, format="PNG", optimize=True)
    return out.getvalue()


def _ink_luminance_hist(rgb):
    """非近白像素的亮度 L（0 黑 255 白）直方图。"""
    h = [0] * 256
    data = rgb.tobytes()
    n = 0
    for i in range(0, len(data), 3):
        r, g, b = data[i], data[i + 1], data[i + 2]
        if r + g + b >= _INK_BG_SUM_MAX:
            continue
        L = (r + g + b) // 3
        if L > 255:
            L = 255
        h[L] += 1
        n += 1
    return h, n


def _median_ink_L(rgb) -> tuple[float, int]:
    """墨迹像素亮度中位数（比均值更不受抗锯齿灰边影响）。无墨迹时 (255, 0)。"""
    h, n = _ink_luminance_hist(rgb)
    if n == 0:
        return 255.0, 0
    # 上中位数：n 为偶数时取中间偏「较淡」一侧，避免 n=2 时退化成最暗像素
    k = n // 2
    acc = 0
    for L in range(256):
        acc += h[L]
        if acc > k:
            return float(L), n
    return 255.0, n


def _ink_bbox_rgb(rgb):
    """返回墨迹（非近白）像素的外接矩形 (x0,y0,x1,y1)；无墨迹则 None。"""
    w, h = rgb.size
    data = rgb.tobytes()
    minx, miny, maxx, maxy = w, h, -1, -1
    any_ink = False
    idx = 0
    for y in range(h):
        row_has = False
        for x in range(w):
            r, g, b = data[idx], data[idx + 1], data[idx + 2]
            idx += 3
            if r + g + b >= _INK_BG_SUM_MAX:
                continue
            any_ink = True
            row_has = True
            if x < minx:
                minx = x
            if x > maxx:
                maxx = x
        if row_has:
            if y < miny:
                miny = y
            if y > maxy:
                maxy = y
    if not any_ink or maxx < 0:
        return None
    return (minx, miny, maxx + 1, maxy + 1)


def _crop_rgb(rgb, box):
    return rgb.crop(box)


def _resize_rgb(rgb, new_w: int, new_h: int):
    if new_w < 2 or new_h < 2:
        return rgb
    return rgb.resize((new_w, new_h), Image.Resampling.LANCZOS)


def _reexport_target_width(rgb) -> Image.Image:
    """把图重新拉到统一像素宽度（与入库/Word 缩放假设一致）。"""
    w, h = rgb.size
    if w < 2 or h < 2:
        return rgb
    if w == _TARGET_EXPORT_W:
        return rgb
    nh = max(1, int(round(h * (_TARGET_EXPORT_W / float(w)))))
    return _resize_rgb(rgb, _TARGET_EXPORT_W, nh)


def _equalize_ink_scale_pair(s_rgb, d_rgb):
    """
    让签名/日期在“笔画粗细观感”上更接近：
    - 先裁剪到墨迹外接矩形（去掉大量留白导致的缩放差异）
    - 再按墨迹面积几何均值做互相缩放（避免只靠亮度对齐仍显一粗一细）
    """
    bs = _ink_bbox_rgb(s_rgb)
    bd = _ink_bbox_rgb(d_rgb)
    if not bs or not bd:
        return s_rgb, d_rgb
    cs = _crop_rgb(s_rgb, bs)
    cd = _crop_rgb(d_rgb, bd)
    as_ = max(1, cs.size[0] * cs.size[1])
    ad = max(1, cd.size[0] * cd.size[1])
    # 面积更接近几何均值：大的略缩小、小的略放大（都有上下限，避免极端）
    g = math.sqrt(float(as_) * float(ad))
    fs = max(0.75, min(1.35, g / float(as_)))
    fd = max(0.75, min(1.35, g / float(ad)))
    sw, sh = cs.size
    dw, dh = cd.size
    cs2 = _resize_rgb(cs, max(2, int(round(sw * fs))), max(2, int(round(sh * fs))))
    cd2 = _resize_rgb(cd, max(2, int(round(dw * fd))), max(2, int(round(dh * fd))))
    # 白底铺回固定画布，避免裁剪后尺寸漂移
    def pad_to(w0, h0, im):
        canvas = Image.new("RGB", (w0, h0), (255, 255, 255))
        x = max(0, (w0 - im.size[0]) // 2)
        y = max(0, (h0 - im.size[1]) // 2)
        canvas.paste(im, (x, y))
        return canvas

    tw = max(cs2.size[0], cd2.size[0], 32)
    th = max(cs2.size[1], cd2.size[1], 32)
    return pad_to(tw, th, cs2), pad_to(tw, th, cd2)


def _apply_ink_luminance_shift_rgb(rgb, delta: int):
    """仅对非近白像素做 RGB 同步平移，使墨迹整体变亮/变暗。"""
    if delta == 0:
        return rgb
    data = rgb.tobytes()
    buf = bytearray(data)
    for i in range(0, len(buf), 3):
        r, g, b = buf[i], buf[i + 1], buf[i + 2]
        if r + g + b >= _INK_BG_SUM_MAX:
            continue
        buf[i] = max(0, min(255, r + delta))
        buf[i + 1] = max(0, min(255, g + delta))
        buf[i + 2] = max(0, min(255, b + delta))
    return Image.frombytes("RGB", rgb.size, bytes(buf))


def prepare_png_for_word(png_bytes: Optional[bytes]) -> Optional[bytes]:
    im = _png_to_flat_rgb_image(png_bytes)
    if im is None:
        return png_bytes
    return _rgb_to_png_bytes(im)


def prepare_signature_date_pair_for_word(
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
    *,
    equalize_ink_scale: bool = True,
) -> tuple[Optional[bytes], Optional[bytes]]:
    """
    同一角色同时有签名+日期时：
    - 可选：先按墨迹外接框与面积做轻微互相缩放，让笔画粗细更接近（Word 默认开启）；
    - 再把两者墨迹亮度对齐到「更淡」的一方，避免历史窄图/透明抗锯齿导致 Word/Excel 里一深一浅。
    Excel 若更追求清晰、可接受签名与日期略不一致，可传 equalize_ink_scale=False。
    """
    if not sig_png or not date_png or Image is None:
        s = prepare_png_for_word(sig_png) if sig_png else None
        d = prepare_png_for_word(date_png) if date_png else None
        return s, d
    s_rgb = _png_to_flat_rgb_image(sig_png)
    d_rgb = _png_to_flat_rgb_image(date_png)
    if s_rgb is None or d_rgb is None:
        return (
            prepare_png_for_word(sig_png) or sig_png,
            prepare_png_for_word(date_png) or date_png,
        )
    if equalize_ink_scale:
        s_rgb, d_rgb = _equalize_ink_scale_pair(s_rgb, d_rgb)
        s_rgb = _reexport_target_width(s_rgb)
        d_rgb = _reexport_target_width(d_rgb)
    ms, ns = _median_ink_L(s_rgb)
    md, nd = _median_ink_L(d_rgb)
    if ns == 0 or nd == 0:
        return _rgb_to_png_bytes(s_rgb), _rgb_to_png_bytes(d_rgb)
    # 亮度越高 = 墨迹越淡；对齐到较淡一侧（必要时把更深的图整体提亮）
    target = max(ms, md)
    s2 = _apply_ink_luminance_shift_rgb(s_rgb, int(round(target - ms)))
    d2 = _apply_ink_luminance_shift_rgb(d_rgb, int(round(target - md)))
    return _rgb_to_png_bytes(s2), _rgb_to_png_bytes(d2)
