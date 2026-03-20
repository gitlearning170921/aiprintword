# -*- coding: utf-8 -*-
"""
Excel 文档处理：可选接受修订、打印
依赖 Windows + 已安装的 WPS 表格 或 Microsoft Excel，使用 COM (pywin32)
默认使用 WPS（KET.Application）；环境变量 USE_OFFICE=1 时使用 Excel。
"""
import os
import re
import time
from collections import Counter

try:
    import win32com.client
    import pythoncom
    import pywintypes
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None

from config import EXCEL_PROGID

RPC_E_SERVER_UNAVAILABLE = -2147023170
RPC_E_SERVER_UNAVAILABLE_ALT = -2147023174
RPC_E_CALL_REJECTED = -2147418111
COM_E_EXCEPTION = -2147352567
COM_E_FAIL = -2147467259
COM_E_OBJDEF = -2146827864  # Excel 对象/应用定义错误
RPC_RETRY_CODES = (RPC_E_SERVER_UNAVAILABLE, RPC_E_SERVER_UNAVAILABLE_ALT, RPC_E_CALL_REJECTED)

_EXCEL_COM_FRIENDLY = (
    "文档处理时发生意外，请确认 WPS/Excel 已安装、文档未被占用，或稍后重试。"
)
_COM_USER_CODES = (COM_E_EXCEPTION, COM_E_FAIL, COM_E_OBJDEF)


def _reraise_with_step(step: str, e: BaseException) -> None:
    """将异常包装为带【步骤名】的 RuntimeError，便于界面定位失败环节。"""
    code = None
    try:
        if e.args:
            code = e.args[0]
    except Exception:
        pass
    if isinstance(code, int) and code in _COM_USER_CODES:
        raise RuntimeError(f"【{step}】{_EXCEL_COM_FRIENDLY}") from e
    raise RuntimeError(f"【{step}】{e}") from e


def _get_excel_app(visible=False):
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    for attempt in range(3):
        try:
            excel = win32com.client.dynamic.Dispatch(EXCEL_PROGID)
            excel.Visible = False
            excel.DisplayAlerts = False
            return excel
        except Exception as e:
            if getattr(e, "args", (None,))[0] in RPC_RETRY_CODES and attempt < 2:
                time.sleep(2)
                continue
            raise


def _save_wb(wb, original_path, save_path):
    """保存工作簿：同路径用 Save() 避免格式弹窗，不同路径用 SaveAs()。"""
    if os.path.normcase(os.path.abspath(save_path)) == os.path.normcase(os.path.abspath(original_path)):
        wb.Save()
    else:
        wb.SaveAs(os.path.abspath(save_path))


# 常见标黄/高亮 ColorIndex（已不再用于自动清除，保留供可选逻辑）
YELLOW_HIGHLIGHT_INDICES = (6, 36, 43, 19)
# 工作表类型：仅处理普通工作表，跳过图表表等
XL_WORKSHEET = -4167
# 单表最多逐格处理单元格数，避免超大表卡死 COM
_MAX_CELLS_PER_SHEET = 80000


def _iter_used_cells(ws, ur):
    """按行列迭代 UsedRange 内单元格，避免 ur.Cells 枚举器在 WPS 上触发 COM 错误。"""
    if ur is None:
        return
    try:
        top, left = int(ur.Row), int(ur.Column)
        nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
    except Exception:
        return
    if nr * nc > _MAX_CELLS_PER_SHEET:
        nr = max(1, _MAX_CELLS_PER_SHEET // max(nc, 1))
    for i in range(nr):
        for j in range(nc):
            try:
                yield ws.Cells(top + i, left + j)
            except Exception:
                continue


def _apply_font_black_sheet(ws):
    """
    将工作表已用区域字体设为黑色。合并单元格下整块 UsedRange.Font 常会报 COM 错，故分级降级。
    """
    try:
        ur = ws.UsedRange
    except Exception:
        return
    if ur is None:
        return
    try:
        ur.Font.Color = 0
        return
    except Exception:
        pass
    try:
        for area in ur.Areas:
            try:
                area.Font.Color = 0
            except Exception:
                for cell in _iter_used_cells(ws, area):
                    try:
                        cell.Font.Color = 0
                    except Exception:
                        pass
    except Exception:
        for cell in _iter_used_cells(ws, ur):
            try:
                cell.Font.Color = 0
            except Exception:
                pass


# 风险矩阵外需清除的填充色 ColorIndex（黄/绿/红等）
_XL_COLOR_NONE = -4142
_FILL_CLEAR_COLORINDEX = frozenset(
    {3, 4, 6, 7, 10, 14, 19, 22, 35, 36, 38, 43, 45, 50, 53, 54}
)
_MAX_ROW_AUTOFIT = 8000
_MAX_COL_AUTOFIT = 256
# Excel Borders 索引：外框 + 内部网格
_XL_BORDER_LEFT = 7
_XL_BORDER_TOP = 8
_XL_BORDER_BOTTOM = 9
_XL_BORDER_RIGHT = 10
_XL_BORDER_INSIDE_V = 11
_XL_BORDER_INSIDE_H = 12
_XL_LINE_CONTINUOUS = 1
_XL_BORDER_THIN = 2


def _xl_border_line_visible(linestyle):
    if linestyle is None:
        return False
    try:
        v = int(linestyle)
    except (TypeError, ValueError):
        return False
    return v > 0 and v != -4142


def _tl_for_cell(ws, r, c):
    try:
        ce = ws.Cells(r, c)
        if ce.MergeCells:
            return ce.MergeArea.Cells(1, 1)
        return ce
    except Exception:
        return None


def _row_tls_for_horizontal_edge(ws, r, left, nc, bottom=True):
    """行 r 上参与底边/顶边的逻辑格，按列排序。"""
    seen = set()
    out = []
    for c in range(left, left + nc):
        try:
            ce = ws.Cells(r, c)
            if ce.MergeCells:
                ma = ce.MergeArea
                mr = int(ma.Row)
                mrows = int(ma.Rows.Count)
                mc = int(ma.Column)
                mcols = int(ma.Columns.Count)
                if bottom:
                    if r != mr + mrows - 1:
                        continue
                else:
                    if r != mr:
                        continue
                key = (mr, mc, mrows, mcols)
                tl = ma.Cells(1, 1)
            else:
                key = (r, c, 1, 1)
                tl = ce
            if key in seen:
                continue
            seen.add(key)
            out.append((c, tl))
        except Exception:
            pass
    out.sort(key=lambda x: x[0])
    return out


def _copy_border_style_from_sample(dst_border, sample_border):
    try:
        dst_border.LineStyle = int(sample_border.LineStyle)
    except Exception:
        try:
            dst_border.LineStyle = _XL_LINE_CONTINUOUS
        except Exception:
            return
    try:
        dst_border.Color = sample_border.Color
    except Exception:
        try:
            dst_border.Color = 0
        except Exception:
            pass
    try:
        dst_border.Weight = sample_border.Weight
    except Exception:
        try:
            dst_border.Weight = _XL_BORDER_THIN
        except Exception:
            pass


def _cell_has_value_for_border(cell):
    """单元格是否有内容（空、纯空白视为无，不参与框线修复也不被改框线）。"""
    try:
        v = cell.Value
        if v is None:
            return False
        if isinstance(v, str):
            return bool(v.strip())
        return True
    except Exception:
        return False


def _excel_unify_row_borders_inconsistent_only(wb):
    """
    只修正同一行内框线不统一：仅针对**有值**的单元格；空格不参与比较、也不被补线。
    有值格之间若底/顶边或相邻竖线一侧有一侧无，则补成与同行有值格一致。不修改风险矩阵格。
    """
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
        except Exception:
            continue
        if ur is None:
            continue
        top = int(ur.Row)
        left = int(ur.Column)
        nr = min(int(ur.Rows.Count), _MAX_ROW_AUTOFIT)
        nc = min(int(ur.Columns.Count), _MAX_COL_AUTOFIT)
        for ri in range(nr):
            r = top + ri
            for edge_bottom in (True, False):
                pairs = _row_tls_for_horizontal_edge(ws, r, left, nc, bottom=edge_bottom)
                tls = [
                    tl
                    for _, tl in pairs
                    if not _cell_is_risk_matrix_cell(tl) and _cell_has_value_for_border(tl)
                ]
                if len(tls) < 2:
                    continue
                bid = _XL_BORDER_BOTTOM if edge_bottom else _XL_BORDER_TOP
                vis = []
                for tl in tls:
                    try:
                        vis.append(_xl_border_line_visible(tl.Borders(bid).LineStyle))
                    except Exception:
                        vis.append(False)
                if not any(vis) or not any(not v for v in vis):
                    continue
                sample = None
                for tl, v in zip(tls, vis):
                    if v:
                        sample = tl.Borders(bid)
                        break
                if not sample:
                    continue
                for tl, v in zip(tls, vis):
                    if not v and _cell_has_value_for_border(tl):
                        try:
                            _copy_border_style_from_sample(tl.Borders(bid), sample)
                        except Exception:
                            pass
            for c in range(left, left + nc - 1):
                try:
                    tl_l = _tl_for_cell(ws, r, c)
                    tl_r = _tl_for_cell(ws, r, c + 1)
                    if tl_l is None or tl_r is None:
                        continue
                    if _cell_is_risk_matrix_cell(tl_l) or _cell_is_risk_matrix_cell(tl_r):
                        continue
                    if tl_l == tl_r:
                        continue
                    bl = _xl_border_line_visible(tl_l.Borders(_XL_BORDER_RIGHT).LineStyle)
                    br = _xl_border_line_visible(tl_r.Borders(_XL_BORDER_LEFT).LineStyle)
                    if bl and not br and _cell_has_value_for_border(tl_r):
                        _copy_border_style_from_sample(
                            tl_r.Borders(_XL_BORDER_LEFT), tl_l.Borders(_XL_BORDER_RIGHT)
                        )
                    elif br and not bl and _cell_has_value_for_border(tl_l):
                        _copy_border_style_from_sample(
                            tl_l.Borders(_XL_BORDER_RIGHT), tl_r.Borders(_XL_BORDER_LEFT)
                        )
                except Exception:
                    pass


def _cell_text_contains_roman_risk_level(val):
    """文本中包含 I / II / III 即视为风险等级相关（不要求整格只有等级）。"""
    if val is None:
        return False
    s = str(val).strip()
    if not s:
        return False
    u = s.upper()
    if "III" in u:
        return True
    # II 且不是 III 的一部分；可匹配「等级II」等
    if re.search(r"(?<![I])II(?!I)", s, re.I):
        return True
    # 单独的 I（避免匹配 Item、Risk 等词内的 I）
    if re.search(r"\bI\b", s, re.I):
        return True
    return False


def _cell_has_color_fill(cell):
    try:
        idx = int(cell.Interior.ColorIndex)
        if idx <= 0 or idx == 2 or idx == _XL_COLOR_NONE:
            return False
        return True
    except Exception:
        return False


def _cell_is_risk_matrix_cell(cell):
    """同时满足：文本含 I / II / III，且单元格有底色。"""
    try:
        return _cell_text_contains_roman_risk_level(cell.Value) and _cell_has_color_fill(cell)
    except Exception:
        return False


def _excel_clear_fill_outside_matrix(wb):
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
        except Exception:
            continue
        if ur is None:
            continue
        for cell in _iter_used_cells(ws, ur):
            if _cell_is_risk_matrix_cell(cell):
                continue
            try:
                idx = cell.Interior.ColorIndex
                if idx is None:
                    continue
                try:
                    idx = int(idx)
                except Exception:
                    continue
                if idx <= 0 or idx == 2 or idx == _XL_COLOR_NONE:
                    continue
                if idx in _FILL_CLEAR_COLORINDEX:
                    cell.Interior.ColorIndex = _XL_COLOR_NONE
            except Exception:
                pass


def _cell_font_size_pt(cell):
    """单元格字号（磅），异常或缺省时按 11。"""
    try:
        s = cell.Font.Size
        if s is None:
            return 11.0
        return max(8.0, min(96.0, float(s)))
    except Exception:
        return 11.0


def _estimate_min_row_height_pt(text, col_width_chars, font_size_pt=11.0):
    """按字数、列宽、字号估算最小行高（磅）；略留余量防裁字，避免过大导致整表/封面被撑换页。"""
    if not text:
        return 0
    fs = max(8.0, min(96.0, float(font_size_pt or 11.0)))
    per_line = max(15.5, fs * 1.12 + 5.5)
    pad = max(9.0, fs * 0.12 + 5.0)
    t = str(text).replace("\r\n", "\n").replace("\r", "\n")
    lines_from_breaks = t.count("\n") + 1
    flat = "".join(t.split())
    cw = max(5, min(100, float(col_width_chars) or 10))
    wrap_lines = max(1, int((len(flat) + max(1, int(cw * 0.65)) - 1) / max(1, int(cw * 0.65))))
    total = max(lines_from_breaks, wrap_lines)
    return min(409.0, max(14.0, total * per_line + pad))


_XL_V_ALIGN_TOP = -4160


def _excel_row_min_height_by_max_font(ws, row, left, nc):
    """该行有内容格中最大字号对应的最小行高（合并格只扫左上角）。"""
    mx_fs = 0.0
    for jj in range(nc):
        col = left + jj
        try:
            ce = ws.Cells(row, col)
            if ce.MergeCells:
                ma = ce.MergeArea
                if int(ce.Row) != int(ma.Row) or int(ce.Column) != int(ma.Column):
                    continue
                probe = ma.Cells(1, 1)
            else:
                probe = ce
            v = probe.Value
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            mx_fs = max(mx_fs, _cell_font_size_pt(probe))
        except Exception:
            pass
    if mx_fs <= 0:
        return 13.0
    return min(409.0, max(13.5, mx_fs * 1.18 + 4.0))


def _excel_autofit_rows_outside_matrix(wb):
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
        except Exception:
            continue
        if ur is None:
            continue
        top = int(ur.Row)
        left = int(ur.Column)
        nr = min(int(ur.Rows.Count), _MAX_ROW_AUTOFIT)
        nc = min(int(ur.Columns.Count), _MAX_COL_AUTOFIT)
        for ri in range(nr):
            row = top + ri
            try:
                need_h = 14.0
                for jj in range(nc):
                    col = left + jj
                    try:
                        ce = ws.Cells(row, col)
                        if ce.MergeCells:
                            ma = ce.MergeArea
                            if int(ce.Row) != int(ma.Row) or int(ce.Column) != int(ma.Column):
                                continue
                            cw = 8.0
                            try:
                                for k in range(1, int(ma.Columns.Count) + 1):
                                    cw += float(
                                        ws.Cells(int(ma.Row), int(ma.Column) + k - 1).ColumnWidth
                                    )
                            except Exception:
                                cw = 40.0
                        else:
                            cw = float(ce.ColumnWidth) if ce.ColumnWidth else 10.0
                        if _cell_is_risk_matrix_cell(ce):
                            continue
                        v = ce.Value
                        if v is None:
                            continue
                        txt = str(v).strip()
                        if not txt:
                            continue
                        tl = ce.MergeArea.Cells(1, 1) if ce.MergeCells else ce
                        fs = _cell_font_size_pt(tl)
                        try:
                            if ce.MergeCells:
                                ce.MergeArea.Cells(1, 1).WrapText = True
                            else:
                                ce.WrapText = True
                        except Exception:
                            pass
                        est = _estimate_min_row_height_pt(txt, cw, fs)
                        need_h = max(need_h, est)
                    except Exception:
                        pass
                try:
                    need_h = max(need_h, _excel_row_min_height_by_max_font(ws, row, left, nc))
                except Exception:
                    pass
                try:
                    ws.Rows(row).AutoFit()
                except Exception:
                    pass
                try:
                    ws.Rows(row).AutoFit()
                except Exception:
                    pass
                try:
                    cur = float(ws.Rows(row).RowHeight)
                except Exception:
                    cur = 15.0
                target = max(cur, need_h)
                if target > cur + 0.5:
                    ws.Rows(row).RowHeight = min(409.0, target)
                if float(ws.Rows(row).RowHeight) < 14:
                    ws.Rows(row).RowHeight = 14
            except Exception:
                pass


def _header_col_empty_across_rows(ws, top, left, ci, h):
    """表头区内该列在各行均无文本，视为表与表之间的空列分隔。"""
    col = left + ci
    for ri in range(h):
        try:
            v = ws.Cells(top + ri, col).Value
            if v is not None and str(v).strip():
                return False
        except Exception:
            return False
    return True


def _header_row_top_left_cells(ws, row, left, c0, c1):
    """该行在 [c0,c1] 内每个「逻辑格」的左上角单元格（横向合并只算一格）。"""
    seen_merge = set()
    out = []
    for ci in range(c0, c1 + 1):
        try:
            ce = ws.Cells(row, left + ci)
            if ce.MergeCells:
                ma = ce.MergeArea
                key = (
                    int(ma.Row),
                    int(ma.Column),
                    int(ma.Rows.Count),
                    int(ma.Columns.Count),
                )
                if key in seen_merge:
                    continue
                seen_merge.add(key)
                out.append(ma.Cells(1, 1))
            else:
                out.append(ce)
        except Exception:
            pass
    return out


def _collect_header_fill_counts_for_range(ws, row, left, c0, c1):
    """本行本块内**有内容**的表头格统计多数底色（空单元格不参与统计）。"""
    tls = _header_row_top_left_cells(ws, row, left, c0, c1)
    counts = Counter()
    for tl in tls:
        try:
            if not _cell_has_value_for_border(tl):
                continue
            if _cell_is_risk_matrix_cell(tl):
                continue
            idx = int(tl.Interior.ColorIndex)
            if 15 <= idx <= 25 or idx in (47, 48, 49, 50, 52):
                counts[idx] += 1
        except Exception:
            pass
    if not counts:
        for tl in tls:
            try:
                if not _cell_has_value_for_border(tl):
                    continue
                if _cell_is_risk_matrix_cell(tl):
                    continue
                idx = int(tl.Interior.ColorIndex)
                if (
                    idx > 0
                    and idx not in (2, 3, 4, 6, 7, 10)
                    and idx not in _FILL_CLEAR_COLORINDEX
                    and idx != _XL_COLOR_NONE
                    and idx >= 10
                ):
                    counts[idx] += 1
            except Exception:
                pass
    return counts


def _unify_header_fill_gray(wb):
    """
    同一表头内统一底色：仅处理**有内容**的单元格；空格不统计多数色、也不补底。
    按空列分表块，每块每行内有值格多数色为参照，仅给仍有值且白/无填充的格补色。
    不覆盖风险矩阵格。
    """
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
        except Exception:
            continue
        if ur is None:
            continue
        top = int(ur.Row)
        left = int(ur.Column)
        nr = int(ur.Rows.Count)
        nc = int(ur.Columns.Count)
        if nc < 4:
            continue
        h = min(5, nr) if nc >= 8 else min(3, nr)
        segments = []
        i = 0
        while i < nc:
            while i < nc and _header_col_empty_across_rows(ws, top, left, i, h):
                i += 1
            if i >= nc:
                break
            j = i
            while j < nc and not _header_col_empty_across_rows(ws, top, left, j, h):
                j += 1
            segments.append((i, j - 1))
            i = j
        for c0, c1 in segments:
            for ri in range(h):
                row = top + ri
                counts = _collect_header_fill_counts_for_range(ws, row, left, c0, c1)
                if not counts:
                    continue
                target = counts.most_common(1)[0][0]
                seen = set()
                for ci in range(c0, c1 + 1):
                    try:
                        ce = ws.Cells(row, left + ci)
                        if ce.MergeCells:
                            ma = ce.MergeArea
                            key = (
                                int(ma.Row),
                                int(ma.Column),
                                int(ma.Rows.Count),
                                int(ma.Columns.Count),
                            )
                            if key in seen:
                                continue
                            seen.add(key)
                            tl = ma.Cells(1, 1)
                        else:
                            tl = ce
                        if _cell_is_risk_matrix_cell(tl):
                            continue
                        if not _cell_has_value_for_border(tl):
                            continue
                        idx = int(tl.Interior.ColorIndex)
                        if idx in (0, 2, _XL_COLOR_NONE, -4105):
                            try:
                                tl.Interior.Pattern = 1
                                tl.Interior.ColorIndex = target
                            except Exception:
                                pass
                    except Exception:
                        pass


def excel_normalize_matrix_and_layout(xl_path, save_path=None):
    """
    风险矩阵格判定：文本含 I/II/III 且带填充色；矩阵外清黄/绿/红底；表头底色仅对有内容格按块按行补白底；
    矩阵外行换行并估算行高。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(xl_path)
    save_path = save_path or path
    save_path = os.path.abspath(save_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    excel = None
    wb = None
    try:
        try:
            excel = _get_excel_app(visible=False)
        except Exception as e:
            _reraise_with_step("启动 WPS/Excel（矩阵外排版）", e)
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=False, UpdateLinks=0)
        except Exception as e:
            _reraise_with_step("打开工作簿（矩阵外排版）", e)
        try:
            _excel_clear_fill_outside_matrix(wb)
            _unify_header_fill_gray(wb)
            _excel_autofit_rows_outside_matrix(wb)
            _excel_unify_row_borders_inconsistent_only(wb)
            _ensure_print_area_covers_used_range(wb)
        except Exception:
            pass
        try:
            _save_wb(wb, path, save_path)
        except Exception as e:
            _reraise_with_step("保存工作簿（矩阵外排版）", e)
        return True
    finally:
        if wb:
            try:
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def check_formal_document(xl_path, check_highlight=True, check_comments=True, check_font_color=True):
    """
    第一步：正式性检查（覆盖判定点）。判定工作簿是否为“正式文档”。
    返回 (passed: bool, issues: list[str])。
    """
    if win32com is None:
        return True, []
    path = os.path.abspath(xl_path)
    if not os.path.isfile(path):
        return False, ["文件不存在"]
    issues = []
    excel = None
    wb = None
    try:
        try:
            excel = _get_excel_app(visible=False)
        except Exception as e:
            return False, [f"【正式性检查-启动表格】{e}"]
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=True, UpdateLinks=0)
        except Exception as e:
            try:
                if e.args and e.args[0] in _COM_USER_CODES:
                    return False, [f"【正式性检查-打开工作簿】{_EXCEL_COM_FRIENDLY}"]
            except Exception:
                pass
            return False, [f"【正式性检查-打开工作簿】{e}"]
        if check_highlight:
            has_highlight = False
            for ws in wb.Worksheets:
                try:
                    if ws.Type != XL_WORKSHEET:
                        continue
                    ur = ws.UsedRange
                    if ur is None:
                        continue
                    for c in _iter_used_cells(ws, ur):
                        try:
                            if c.Interior.ColorIndex in YELLOW_HIGHLIGHT_INDICES:
                                has_highlight = True
                                break
                        except Exception:
                            pass
                    if has_highlight:
                        break
                except Exception:
                    pass
            if has_highlight:
                issues.append("存在单元格标黄/高亮")
        if check_comments:
            comment_count = 0
            for ws in wb.Worksheets:
                try:
                    if ws.Type != XL_WORKSHEET:
                        continue
                    ur = ws.UsedRange
                    if ur is None:
                        continue
                    for c in _iter_used_cells(ws, ur):
                        try:
                            if c.Comment is not None:
                                comment_count += 1
                        except Exception:
                            pass
                except Exception:
                    pass
            if comment_count > 0:
                issues.append(f"存在 {comment_count} 个单元格批注")
        if check_font_color:
            has_non_black = False
            for ws in wb.Worksheets:
                try:
                    if ws.Type != XL_WORKSHEET:
                        continue
                    ur = ws.UsedRange
                    if ur is None:
                        continue
                    try:
                        if ur.Font.Color != 0 and ur.Font.Color != 16777215:
                            has_non_black = True
                            break
                    except Exception:
                        for c in _iter_used_cells(ws, ur):
                            try:
                                col = c.Font.Color
                                if col not in (0, 16777215, -4105):  # -4105 = 自动色
                                    has_non_black = True
                                    break
                            except Exception:
                                pass
                            if has_non_black:
                                break
                except Exception:
                    pass
                if has_non_black:
                    break
            if has_non_black:
                issues.append("存在非黑色字体")
        passed = len(issues) == 0
        return passed, issues
    except Exception as e:
        return False, [f"正式性检查异常: {e}"]
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def auto_fix_formal_excel(xl_path, save_path=None):
    """
    正式性检查失败后自动修复：接受修订、删批注、统一黑色；
    风险矩阵外清除标黄/绿/红底，矩阵外行自动拉高（规则同 excel_normalize_matrix_and_layout）。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(xl_path)
    save_path = save_path or path
    save_path = os.path.abspath(save_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    excel = None
    wb = None
    try:
        try:
            excel = _get_excel_app(visible=False)
        except Exception as e:
            _reraise_with_step("启动 WPS/Excel（自动修复）", e)
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=False, UpdateLinks=0)
        except Exception as e:
            _reraise_with_step("打开工作簿（自动修复，需可写）", e)
        try:
            wb.AcceptAllChanges()
        except Exception:
            pass
        for ws in wb.Worksheets:
            try:
                if ws.Type == XL_WORKSHEET:
                    ws.Cells.ClearComments()
            except Exception:
                pass
        for ws in wb.Worksheets:
            try:
                if ws.Type == XL_WORKSHEET:
                    _apply_font_black_sheet(ws)
            except Exception:
                pass
        try:
            _excel_clear_fill_outside_matrix(wb)
            _unify_header_fill_gray(wb)
            _excel_autofit_rows_outside_matrix(wb)
            _excel_unify_row_borders_inconsistent_only(wb)
            _ensure_print_area_covers_used_range(wb)
        except Exception:
            pass
        try:
            _save_wb(wb, path, save_path)
        except Exception as e:
            _reraise_with_step("保存工作簿（自动修复）", e)
        return True
    finally:
        if wb:
            try:
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def ensure_font_black(xl_path, save_path=None):
    """
    将工作簿中所有单元格文字颜色统一为黑色，符合正式文档要求。
    save_path 为 None 时覆盖原文件。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(xl_path)
    save_path = save_path or path
    save_path = os.path.abspath(save_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    excel = None
    wb = None
    try:
        try:
            excel = _get_excel_app(visible=False)
        except Exception as e:
            _reraise_with_step("启动 WPS/Excel（字体改黑）", e)
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=False, UpdateLinks=0)
        except Exception as e:
            _reraise_with_step("打开工作簿（字体改黑）", e)
        for ws in wb.Worksheets:
            try:
                if ws.Type == XL_WORKSHEET:
                    _apply_font_black_sheet(ws)
            except Exception:
                pass
        try:
            _save_wb(wb, path, save_path)
        except Exception as e:
            _reraise_with_step("保存工作簿（字体改黑）", e)
        return True
    finally:
        if wb:
            try:
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _remove_cell_highlights(wb):
    """去除工作簿中单元格的标黄/高亮填充（逐格访问，避免 WPS 对 ur.Cells 枚举报错）。"""
    xlNone = -4142
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
        except Exception:
            continue
        for c in _iter_used_cells(ws, ur):
            try:
                if c.Interior.ColorIndex in YELLOW_HIGHLIGHT_INDICES:
                    c.Interior.ColorIndex = xlNone
            except Exception:
                pass


def accept_all_changes_and_save(xl_path, save_path=None, accept_revisions=True, remove_highlights=False):
    """
    若 accept_revisions 且工作簿启用了跟踪修订，则接受所有修订并保存。
    remove_highlights=True 时会清除常见标黄底色（默认 False，保留风险矩阵等配色）。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(xl_path)
    save_path = save_path or path
    save_path = os.path.abspath(save_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    excel = None
    wb = None
    try:
        excel = _get_excel_app(visible=False)
        wb = excel.Workbooks.Open(path, ReadOnly=False, UpdateLinks=0)
        if accept_revisions:
            try:
                wb.AcceptAllChanges()
            except Exception:
                pass
        if remove_highlights:
            try:
                _remove_cell_highlights(wb)
            except Exception:
                pass
        _save_wb(wb, path, save_path)
        return True
    finally:
        if wb:
            try:
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _apply_excel_print_layout_for_readability(wb):
    """
    仅打印前在内存中生效：**不修改打印方向（横向/纵向）**，沿用源工作簿设置。
    - 水平居中；
    - 列数较多时：FitToPagesWide=1（一页宽不拆列）、FitToPagesTall=999（纵向可多页）；
    - 不缩小 PrintArea；页码起始用自动。
    完整数据范围依赖 `_ensure_print_area_covers_used_range`。
    """
    try:
        inch = float(wb.Application.InchesToPoints(1))
    except Exception:
        inch = 72.0
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
            nc = int(ur.Columns.Count) if ur else 0
        except Exception:
            nc = 0
        try:
            ps = ws.PageSetup
            try:
                ps.CenterHorizontally = True
            except Exception:
                pass

            if nc >= 6:
                try:
                    ps.Zoom = False
                except Exception:
                    pass
                try:
                    ps.FitToPagesWide = 1
                except Exception:
                    pass
                try:
                    ps.FitToPagesTall = 999
                except Exception:
                    pass

            try:
                ps.LeftMargin = inch * 0.4
                ps.RightMargin = inch * 0.4
                # 拉开页眉/页脚与正文距离，避免重影或覆盖正文
                ps.TopMargin = inch * 0.75
                ps.BottomMargin = inch * 0.6
                ps.HeaderMargin = inch * 0.3
                ps.FooterMargin = inch * 0.3
            except Exception:
                pass
            try:
                ps.FirstPageNumber = -4105  # xlAutomatic
            except Exception:
                pass
        except Exception:
            pass


_SAFE_PAGE_NUM_PAT = re.compile(r"(?i)\bpage\s*[:：]?\s*\d+(\s*/\s*\d+)?\b")


def _used_range_address_for_print(ur):
    """COM 下 Address 多为只读属性，勿用关键字参数调用。"""
    if ur is None:
        return None
    try:
        a = ur.Address
        if isinstance(a, str) and a.strip():
            return a.strip()
    except Exception:
        pass
    try:
        return str(ur.GetAddress(True, True, 1, False, None)).strip()
    except Exception:
        return None


def _ensure_print_area_covers_used_range(wb):
    """
    将**每个**普通工作表的 `PrintArea` 设为当前 `UsedRange` 地址，保证整本工作簿各页内容按 Excel 已用区域完整打印；
    不修改纸张方向（横向/纵向），不改缩放模式以外的版式。
    在规范化保存与打印前各执行一次（打印为内存，保存会写回文件）。
    """
    try:
        wb.Application.Calculate()
    except Exception:
        pass
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ur = ws.UsedRange
            if ur is None:
                continue
            addr = _used_range_address_for_print(ur)
            if addr:
                ws.PageSetup.PrintArea = addr
        except Exception:
            pass


def _normalize_page_markers_safe(wb):
    """
    仅替换页眉/页脚中的“Page:数字[/数字]”片段为动态页码，避免改动正文编号等字段。
    """
    for ws in wb.Worksheets:
        try:
            if ws.Type != XL_WORKSHEET:
                continue
        except Exception:
            continue
        try:
            ps = ws.PageSetup
            for attr in ("LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter"):
                try:
                    cur = getattr(ps, attr)
                    if cur is None:
                        continue
                    txt = str(cur)
                    if not txt.strip():
                        continue
                    if "&P" in txt.upper():
                        continue
                    newv = _SAFE_PAGE_NUM_PAT.sub("Page:&P/&N", txt)
                    if newv != txt:
                        setattr(ps, attr, newv)
                except Exception:
                    pass
        except Exception:
            pass


def _printer_line_for_excel_com(printer_name):
    """
    Excel/WPS ActivePrinter 需要「打印机名 on 端口」形式；
    界面传入的多为 EnumPrinters 的 pPrinterName，直接赋值会 COM 报错。
    """
    if not printer_name:
        return None
    try:
        import win32print

        h = win32print.OpenPrinter(printer_name)
        try:
            info = win32print.GetPrinter(h, 2)
            pname = (info.get("pPrinterName") or printer_name).strip()
            port = (info.get("pPortName") or "").strip()
            if pname and port:
                return f"{pname} on {port}"
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        pass
    return None


def print_excel_workbook(xl_path, printer_name=None, copies=1):
    """
    打印前（仅内存，不写回文件）：矩阵外行高、表头底色、框线；打印布局仅居中与一页宽适配，**不改纸张方向**；
    安全修正页码标记（仅 Page:数字 片段）。
    若 PrintArea 未覆盖已用区域，会扩展到 UsedRange，保证记录打全。
    指定打印机：先 ActivePrinter，失败则临时改系统默认打印机。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(xl_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    excel = None
    wb = None
    old_default_printer = None
    try:
        try:
            excel = _get_excel_app(visible=False)
        except Exception as e:
            _reraise_with_step("启动 WPS/Excel（打印）", e)
        try:
            wb = excel.Workbooks.Open(path, ReadOnly=False, UpdateLinks=0)
        except Exception as e:
            _reraise_with_step("打开工作簿（打印）", e)
        try:
            _excel_autofit_rows_outside_matrix(wb)
        except Exception:
            pass
        try:
            _unify_header_fill_gray(wb)
        except Exception:
            pass
        try:
            _excel_unify_row_borders_inconsistent_only(wb)
        except Exception:
            pass
        try:
            _apply_excel_print_layout_for_readability(wb)
        except Exception:
            pass
        try:
            _ensure_print_area_covers_used_range(wb)
        except Exception:
            pass
        try:
            _normalize_page_markers_safe(wb)
        except Exception:
            pass
        if printer_name:
            line = _printer_line_for_excel_com(printer_name)
            applied = False
            for cand in (c for c in (line, printer_name) if c):
                try:
                    excel.ActivePrinter = cand
                    applied = True
                    break
                except Exception:
                    continue
            if not applied:
                try:
                    import win32print

                    old_default_printer = win32print.GetDefaultPrinter()
                    win32print.SetDefaultPrinter(printer_name)
                except Exception as e:
                    _reraise_with_step("设置打印机（COM 与系统默认均失败）", e)
        try:
            wb.PrintOut(
                From=1,
                To=9999,
                Copies=int(copies),
                Collate=True,
            )
        except Exception as e:
            _reraise_with_step("执行打印（PrintOut）", e)
        time.sleep(3)
        return True
    finally:
        if old_default_printer:
            try:
                import win32print

                win32print.SetDefaultPrinter(old_default_printer)
            except Exception:
                pass
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
