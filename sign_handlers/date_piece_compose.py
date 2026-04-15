"""日期笔迹元件横向拼接：数字 0–9、英文月份整词（pm01–pm12）、句点「.」。
支持版式：中文 2026.04.15；英文 15 April 2026；英文 15.April.2026。"""

from __future__ import annotations

import io
import re
from datetime import date
from typing import Any, Dict, List, Optional, Sequence, Tuple

# 月份与 pm01..pm12 一一对应（整张手写图，不是逐字母）
_MONTH_NAMES = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)

# 常见三字母简称（September → Sep）
_MONTH_ABBREV = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def normalize_piece_kind(kind: str) -> Optional[str]:
    k = (kind or "").strip().lower()
    if re.fullmatch(r"pd[0-9]", k):
        return k
    if re.fullmatch(r"pm(0[1-9]|1[0-2])", k):
        return k
    if re.fullmatch(r"pma(0[1-9]|1[0-2])", k):
        return k
    if k == "pdot":
        return k
    return None


def all_piece_kinds() -> List[str]:
    out = [f"pd{d}" for d in range(10)]
    out.extend([f"pm{m:02d}" for m in range(1, 13)])
    out.extend([f"pma{m:02d}" for m in range(1, 13)])
    out.append("pdot")
    return out


def month_abbr_token(mo: int) -> str:
    return f"pma{mo:02d}"


def month_full_token(mo: int) -> str:
    return f"pm{mo:02d}"


def _parse_iso_ymd(iso: str) -> Tuple[int, int, int]:
    s = (iso or "").strip()
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", s)
    if not m:
        raise ValueError("日期须为 YYYY-MM-DD")
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        date(y, mo, d)
    except ValueError as e:
        raise ValueError("无效日期") from e
    return y, mo, d


def kinds_zh_ymd_dot(iso: str) -> Tuple[List[str], str]:
    """中文数字笔迹：2026.04.15 → 数字 + 句点分隔，需 pd0–pd9 与 pdot。"""
    y, mo, d = _parse_iso_ymd(iso)
    tokens: List[str] = []
    for ch in str(y):
        tokens.append(f"pd{ch}")
    tokens.append("pdot")
    for ch in f"{mo:02d}":
        tokens.append(f"pd{ch}")
    tokens.append("pdot")
    for ch in f"{d:02d}":
        tokens.append(f"pd{ch}")
    label = f"{y:04d}.{mo:02d}.{d:02d}"
    return tokens, label


def kinds_en_dmy_space(iso: str) -> Tuple[List[str], str, List[int]]:
    """英文「15 Apr 2026」：日数字 + 月简称 + 年数字；段间用大间距模拟空格（无需空格笔迹）。"""
    y, mo, d = _parse_iso_ymd(iso)
    ds = str(d)
    ys = str(y)
    tokens: List[str] = []
    for ch in ds:
        tokens.append(f"pd{ch}")
    # 仅要求月份简称元件（pma01..pma12）；不再依赖整词 pm01..pm12
    tokens.append(month_abbr_token(mo))
    for ch in ys:
        tokens.append(f"pd{ch}")
    label = f"{d} {_MONTH_ABBREV[mo - 1]} {y}"
    nd = len(ds)
    n = len(tokens)
    if n < 2:
        gaps: List[int] = []
    else:
        # 字符间距整体收紧；仅在「日-月」「月-年」之间略放大
        gaps = [2] * (n - 1)
        # 最后一笔「日」与「月」、 「月」与「年首位」之间加大间距（模拟空格）
        if nd >= 1:
            gaps[nd - 1] = max(gaps[nd - 1], 10)
        if nd < len(gaps):
            gaps[nd] = max(gaps[nd], 10)
    return tokens, label, gaps


def kinds_en_dot_dmy(iso: str) -> Tuple[List[str], str]:
    """「15.April.2026」：日数字 + . + 月单词 + . + 年数字（原 composite_en）。"""
    y, mo, d = _parse_iso_ymd(iso)
    tokens: List[str] = []
    for ch in str(d):
        tokens.append(f"pd{ch}")
    tokens.append("pdot")
    tokens.append(f"pm{mo:02d}")
    tokens.append("pdot")
    for ch in str(y):
        tokens.append(f"pd{ch}")
    return tokens, f"{d}.{_MONTH_NAMES[mo - 1]}.{y}"


def kinds_for_iso_date(iso: str) -> Tuple[List[str], str]:
    """兼容旧名：等同 kinds_en_dot_dmy。"""
    return kinds_en_dot_dmy(iso)


def _scale_to_height(img: Any, target_h: int) -> Any:
    from PIL import Image

    if target_h <= 0:
        return img
    w, h = img.size
    if h <= 0:
        return img
    if h == target_h:
        return img
    nw = max(1, int(round(w * (target_h / float(h)))))
    return img.resize((nw, target_h), Image.Resampling.LANCZOS)


def compose_png_horizontal(
    png_parts: Sequence[bytes],
    *,
    gap: int = 6,
    gaps: Optional[Sequence[int]] = None,
    target_h: int = 360,
) -> bytes:
    """将多张 PNG 横向拼接，统一高度为 target_h（保持各自宽高比）。
    gaps：长度须为 len(png_parts)-1，表示相邻两张图之间的像素间距；不传则全部使用 gap。
    """
    if not png_parts:
        raise ValueError("没有可拼接的图片")
    try:
        from PIL import Image
    except ImportError as e:
        raise RuntimeError("需要安装 Pillow 才能拼接日期笔迹（pip install Pillow）") from e

    imgs: List[Any] = []
    for b in png_parts:
        if not b:
            raise ValueError("某张笔迹图片为空")
        im = Image.open(io.BytesIO(b)).convert("RGBA")
        im = _scale_to_height(im, target_h)
        # 录入元件通常来自固定宽高画布（例如 900x360），四周空白很大；
        # 拼接前先裁剪到“有效笔迹”的包围盒，避免字符间距被空白撑开。
        try:
            px = im.load()
            w, h = im.size
            minx, miny, maxx, maxy = w, h, -1, -1
            for y in range(h):
                for x in range(w):
                    r, g, b2, a = px[x, y]
                    # “有墨迹”的判定：alpha 足够且不是接近白色
                    if a > 8 and (r + g + b2) < 740:
                        if x < minx:
                            minx = x
                        if y < miny:
                            miny = y
                        if x > maxx:
                            maxx = x
                        if y > maxy:
                            maxy = y
            if maxx >= 0 and maxy >= 0:
                pad = 8
                minx = max(0, minx - pad)
                miny = max(0, miny - pad)
                maxx = min(w - 1, maxx + pad)
                maxy = min(h - 1, maxy + pad)
                im = im.crop((minx, miny, maxx + 1, maxy + 1))
        except Exception:
            pass
        imgs.append(im)

    n = len(imgs)
    if n == 1:
        gap_list: List[int] = []
    elif gaps is not None:
        if len(gaps) != n - 1:
            raise ValueError("gaps 长度须为 len(png_parts)-1")
        gap_list = list(gaps)
    else:
        gap_list = [gap] * (n - 1)

    extra = sum(gap_list)
    total_w = sum(im.size[0] for im in imgs) + extra
    h = target_h
    # 用不透明白底，避免浏览器用深色背景渲染透明区域导致“黑底黑字”
    out = Image.new("RGBA", (total_w, h), (255, 255, 255, 255))
    x = 0
    for i, im in enumerate(imgs):
        if i > 0:
            x += gap_list[i - 1]
        # 底对齐：连接符/字母/数字裁剪后高度不同，顶对齐会导致「.」浮在中间或偏上；
        # 底对齐能让基线更一致，中文日期中的连接符更自然（靠右下）。
        y = max(0, h - im.size[1])
        out.paste(im, (x, y), im)
        x += im.size[0]
    buf = io.BytesIO()
    out.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def overlay_dot_on_right_bottom(
    base_png: bytes,
    dot_png: bytes,
    *,
    target_h: int = 360,
    dot_scale: float = 0.28,
    margin: int = 2,
) -> bytes:
    """将句点笔迹叠加到 base 的右下角（用于中文 2026.04.15，使 . 贴近前一位数字）。"""
    if not base_png or not dot_png:
        return base_png or dot_png
    try:
        from PIL import Image
    except ImportError as e:
        raise RuntimeError("需要安装 Pillow 才能拼接日期笔迹（pip install Pillow）") from e

    def _prep(b: bytes) -> Any:
        im = Image.open(io.BytesIO(b)).convert("RGBA")
        im = _scale_to_height(im, target_h)
        # 与 compose_png_horizontal 一致的“去空白边”裁剪逻辑（简化复用）
        try:
            px = im.load()
            w, h = im.size
            minx, miny, maxx, maxy = w, h, -1, -1
            for y in range(h):
                for x in range(w):
                    r, g, b2, a = px[x, y]
                    if a > 8 and (r + g + b2) < 740:
                        if x < minx:
                            minx = x
                        if y < miny:
                            miny = y
                        if x > maxx:
                            maxx = x
                        if y > maxy:
                            maxy = y
            if maxx >= 0 and maxy >= 0:
                pad = 8
                minx = max(0, minx - pad)
                miny = max(0, miny - pad)
                maxx = min(w - 1, maxx + pad)
                maxy = min(h - 1, maxy + pad)
                im = im.crop((minx, miny, maxx + 1, maxy + 1))
        except Exception:
            pass
        return im

    base = _prep(base_png)
    dot = _prep(dot_png)
    bw, bh = base.size
    if bw < 2 or bh < 2:
        return base_png
    # 缩小句点并贴到右下角
    dh = max(1, int(round(bh * float(dot_scale))))
    dot = _scale_to_height(dot, dh)
    dw, dh2 = dot.size
    x = max(0, bw - dw - margin)
    y = max(0, bh - dh2 - margin)
    out = Image.new("RGBA", (bw, bh), (255, 255, 255, 0))
    out.paste(base, (0, 0), base)
    out.paste(dot, (x, y), dot)
    buf = io.BytesIO()
    out.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def render_dot_cell_right_bottom(
    dot_png: bytes,
    *,
    target_h: int = 360,
    cell_w: int = 110,
    dot_scale: float = 0.28,
    margin: int = 2,
) -> bytes:
    """把句点渲染为“单独一格”，但点落在该格的右下角（视觉上贴近前一位数字的右下角）。"""
    if not dot_png:
        raise ValueError("dot_png 为空")
    try:
        from PIL import Image
    except ImportError as e:
        raise RuntimeError("需要安装 Pillow 才能拼接日期笔迹（pip install Pillow）") from e

    im = Image.open(io.BytesIO(dot_png)).convert("RGBA")
    im = _scale_to_height(im, target_h)
    # 复用与拼接一致的裁剪逻辑，避免 dot 自带大空白
    try:
        px = im.load()
        w, h = im.size
        minx, miny, maxx, maxy = w, h, -1, -1
        for y in range(h):
            for x in range(w):
                r, g, b2, a = px[x, y]
                if a > 8 and (r + g + b2) < 740:
                    if x < minx:
                        minx = x
                    if y < miny:
                        miny = y
                    if x > maxx:
                        maxx = x
                    if y > maxy:
                        maxy = y
        if maxx >= 0 and maxy >= 0:
            pad = 8
            minx = max(0, minx - pad)
            miny = max(0, miny - pad)
            maxx = min(w - 1, maxx + pad)
            maxy = min(h - 1, maxy + pad)
            im = im.crop((minx, miny, maxx + 1, maxy + 1))
    except Exception:
        pass

    # 缩小 dot
    dh = max(1, int(round(target_h * float(dot_scale))))
    im = _scale_to_height(im, dh)
    dw, dh2 = im.size
    cw = max(int(cell_w), dw + margin * 2)
    out = Image.new("RGBA", (cw, target_h), (255, 255, 255, 255))
    x = max(0, cw - dw - margin)
    y = max(0, target_h - dh2 - margin)
    out.paste(im, (x, y), im)
    buf = io.BytesIO()
    out.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def month_label_for_kind(pm_kind: str) -> str:
    k = normalize_piece_kind(pm_kind)
    if not k:
        return pm_kind
    if k.startswith("pma") and len(k) == 5:
        try:
            idx = int(k[3:]) - 1
        except ValueError:
            return pm_kind
        if 0 <= idx < 12:
            return _MONTH_ABBREV[idx]
        return pm_kind
    if k.startswith("pm") and len(k) == 4:
        try:
            idx = int(k[2:]) - 1
        except ValueError:
            return pm_kind
        if 0 <= idx < 12:
            return _MONTH_NAMES[idx]
        return pm_kind
    return pm_kind


def piece_kind_label(kind: str) -> str:
    """用于列表展示的中文标签。"""
    k = normalize_piece_kind(kind) or (kind or "").strip().lower()
    if k == "pdot":
        return "日期元件·连接符 ."
    if k.startswith("pd") and len(k) == 3:
        return "日期元件·数字 " + k[2]
    if k.startswith("pma") and len(k) == 5:
        try:
            i = int(k[3:]) - 1
        except ValueError:
            return kind or ""
        if 0 <= i < 12:
            return (
                "日期元件·月份简称 "
                + _MONTH_ABBREV[i]
                + "（"
                + _MONTH_NAMES[i]
                + "）"
            )
        return kind or ""
    if k.startswith("pm") and len(k) == 4:
        try:
            i = int(k[2:]) - 1
        except ValueError:
            return kind or ""
        if 0 <= i < 12:
            return (
                "日期元件·月份整词 "
                + _MONTH_NAMES[i]
                + "（"
                + _MONTH_ABBREV[i]
                + "）"
            )
        return kind or ""
    return kind or ""
