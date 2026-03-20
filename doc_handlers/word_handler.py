# -*- coding: utf-8 -*-
"""
Word 文档处理：检查封面签字、接受所有修订、统一字体为黑色、打印
依赖 Windows + 已安装的 WPS 文字 或 Microsoft Word，使用 COM (pywin32)
默认使用 WPS（KWPS.Application）；环境变量 USE_OFFICE=1 时使用 Word。
"""
import os
import re
import time

try:
    import win32com.client
    import pythoncom
    import pywintypes
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None

from config import WORD_PROGID

# 可重试的 COM 错误码
RPC_E_SERVER_UNAVAILABLE = -2147023170      # 远程过程调用失败
RPC_E_SERVER_UNAVAILABLE_ALT = -2147023174  # RPC 服务器不可用
RPC_E_CALL_REJECTED = -2147418111           # 被呼叫方拒绝接收呼叫（应用正忙/弹窗中）
COM_E_EXCEPTION = -2147352567
COM_E_FAIL = -2147467259
RPC_RETRY_CODES = (RPC_E_SERVER_UNAVAILABLE, RPC_E_SERVER_UNAVAILABLE_ALT, RPC_E_CALL_REJECTED)

# Word 常量
wdGoToPage = 1
wdGoToFirst = 1
wdGoToNext = 2
wdColorBlack = 0
wdColorAutomatic = -1  # 自动颜色视为黑色
wdNoHighlight = 0
# 格式检查/规范化时统一使用的字体（宋体）
FONT_NAME_SONG = "SimSun"
# 页眉页脚：1= primary, 2= first page, 3= even
wdHeaderFooterPrimary = 1
# 域类型：页码、目录
wdFieldPage = 33
wdFieldTOC = 37
# 表格行高：自动适应内容
wdRowHeightAuto = 0
wdRowHeightExactly = 2  # 固定行高，易导致打印裁字
# 表格自动适应：按窗口宽度（行高相对紧凑）；按内容易把行撑得过高换页
wdAutoFitContent = 1
wdAutoFitWindow = 2
# 单元格垂直对齐：顶端，减少行高略紧时上下被裁
wdCellAlignVerticalTop = 1
# 框线：单线、黑色，用于统一表格框线
wdLineStyleSingle = 1
wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4
wdBorderHorizontal = -5
wdBorderVertical = -6


def _normalize_table_borders(doc):
    """统一文档内所有表格的框线为单线、黑色，与其他表格框线一致。WPS 可能不支持部分 Border 属性，逐项 try。"""
    try:
        for i in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(i)
            try:
                for bid in (wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight, wdBorderHorizontal, wdBorderVertical):
                    try:
                        b = tbl.Borders(bid)
                        b.LineStyle = wdLineStyleSingle
                        b.Color = wdColorBlack
                    except Exception:
                        pass
                try:
                    tbl.Borders.InsideLineStyle = wdLineStyleSingle
                except Exception:
                    pass
                try:
                    tbl.Borders.InsideColor = wdColorBlack
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass


def _tables_of_contents_safe(doc):
    """安全获取 TablesOfContents 集合，WPS 可能不支持或返回异常。返回 (count, getter) 或 (0, None)。"""
    try:
        toc = getattr(doc, "TablesOfContents", None)
        if toc is None:
            return 0, None
        n = int(toc.Count)
        return n, toc
    except Exception:
        return 0, None


def _is_risk_matrix_color(c):
    """判断是否为风险矩阵用到的红/黄/绿色（底纹或字体，Word 中常用 BGR/OLE 值）。"""
    if c is None:
        return False
    try:
        v = int(c)
    except (TypeError, ValueError):
        return False
    if v in (wdColorBlack, wdColorAutomatic, -1, 0):
        return False
    if v < 0:
        return False
    if v <= 255 or (128 <= v <= 1000):
        return True
    if 32768 <= v <= 65535:
        return True
    if 8421376 <= v <= 16777215:
        return True
    return False


def _cell_has_risk_color(cell):
    """表格单元格是否含红/黄/绿底纹或字体（风险矩阵特征）。"""
    try:
        try:
            sh = cell.Shading.BackgroundPatternColor
            if _is_risk_matrix_color(sh):
                return True
        except Exception:
            pass
        try:
            r = cell.Range
            if r and r.Font.Color is not None:
                if _is_risk_matrix_color(r.Font.Color):
                    return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def _is_risk_matrix_table(table):
    """判断表格是否为风险矩阵（多格含红/黄/绿标色）。"""
    try:
        nr = int(table.Rows.Count)
        nc = int(table.Columns.Count)
        if nr < 2 or nc < 2:
            return False
        count_risk = 0
        total = 0
        for ri in range(1, nr + 1):
            for ci in range(1, nc + 1):
                try:
                    cell = table.Cell(ri, ci)
                    total += 1
                    if _cell_has_risk_color(cell):
                        count_risk += 1
                except Exception:
                    pass
        return total >= 4 and count_risk >= max(4, total // 4)
    except Exception:
        return False


def _save_risk_matrix_formats(doc):
    """保存所有风险矩阵表格的单元格底纹与字体颜色，返回可传给 _restore_risk_matrix_formats 的数据。"""
    out = []
    try:
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            if not _is_risk_matrix_table(tbl):
                continue
            try:
                nr = int(tbl.Rows.Count)
                nc = int(tbl.Columns.Count)
                cells_data = []
                for ri in range(1, nr + 1):
                    for ci in range(1, nc + 1):
                        try:
                            cell = tbl.Cell(ri, ci)
                            sh = None
                            fc = None
                            try:
                                sh = int(cell.Shading.BackgroundPatternColor)
                            except Exception:
                                pass
                            try:
                                fc = int(cell.Range.Font.Color)
                            except Exception:
                                pass
                            cells_data.append((ri, ci, sh, fc))
                        except Exception:
                            pass
                if cells_data:
                    out.append((ti, cells_data))
            except Exception:
                pass
    except Exception:
        pass
    return out


def _restore_risk_matrix_formats(doc, saved):
    """恢复风险矩阵表格的底纹与字体颜色（正式性修复后保留矩阵标色）。"""
    if not saved:
        return
    try:
        for ti, cells_data in saved:
            try:
                tbl = doc.Tables(ti)
                for ri, ci, sh, fc in cells_data:
                    try:
                        cell = tbl.Cell(ri, ci)
                        if sh is not None:
                            try:
                                cell.Shading.BackgroundPatternColor = sh
                            except Exception:
                                pass
                        if fc is not None:
                            try:
                                cell.Range.Font.Color = fc
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass


def _range_in_risk_matrix(doc, start, end):
    """判断 [start,end) 是否落在任意风险矩阵表格范围内。"""
    try:
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            if not _is_risk_matrix_table(tbl):
                continue
            try:
                r = tbl.Range
                if r.Start <= start and end <= r.End:
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def _get_word_app(visible=False):
    """获取或创建 Word Application 实例。当前线程需已调用 CoInitialize。"""
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    for attempt in range(3):
        try:
            word = win32com.client.dynamic.Dispatch(WORD_PROGID)
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone，禁止一切弹窗
            return word
        except Exception as e:
            is_rpc = getattr(e, "args", (None,))[0] in RPC_RETRY_CODES
            if is_rpc and attempt < 2:
                time.sleep(2)
                continue
            raise


def _com_call(func, *args, retries=3, delay=2, **kwargs):
    """
    带重试的 COM 调用包装器。
    当 WPS/Word 正忙（弹窗、处理中）抛出 RPC_E_CALL_REJECTED 等错误时自动重试。
    """
    for attempt in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            code = getattr(e, "args", (None,))[0]
            if code in RPC_RETRY_CODES and attempt < retries - 1:
                time.sleep(delay)
                continue
            raise


def _save_doc(doc, original_path, save_path):
    """
    保存文档：当目标路径与原路径相同时用 doc.Save() 避免格式弹窗；
    不同路径时用 doc.SaveAs()。
    """
    if os.path.normcase(os.path.abspath(save_path)) == os.path.normcase(os.path.abspath(original_path)):
        doc.Save()
    else:
        doc.SaveAs(os.path.abspath(save_path))


def _is_value_empty(s):
    """判断签字/日期等值是否为空（含占位符）。"""
    if s is None:
        return True
    s = str(s).strip()
    if not s:
        return True
    # 移除常见占位符：空格、下划线、横线、点等
    meaningful = re.sub(r"[\s_\-\u2014\u2015\u2500./\\:\u3000]+", "", s)
    return len(meaningful) < 2


def _extract_field_value(cover_text, keyword, next_keywords):
    """从封面文本中提取关键词后的值（到换行或下一个关键词为止）。"""
    if not cover_text or keyword not in cover_text:
        return None
    idx = cover_text.find(keyword)
    start = idx + len(keyword)
    # 跳过 "：" 或 ":"
    while start < len(cover_text) and cover_text[start] in "：:\t ":
        start += 1
    # 截取到换行或下一个关键词
    end = len(cover_text)
    for nk in next_keywords:
        pos = cover_text.find(nk, start)
        if 0 <= pos < end:
            end = pos
    for sep in ("\r", "\n", "\t"):
        pos = cover_text.find(sep, start)
        if 0 <= pos < end:
            end = pos
    return cover_text[start:end].strip()


def check_cover_signature(doc_path):
    """
    检查 Word 文档封面页签字是否完成。
    通过封面页（第一页）的 作者、审核、批准、日期 是否为空判断。
    任一为空则视为签字未完成。
    """
    if win32com is None:
        return True, "未安装 pywin32，跳过签字检查"
    path = os.path.abspath(doc_path)
    if not os.path.isfile(path):
        return False, "文件不存在"
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=True, AddToRecentFiles=False)
        # 获取第一页文本
        try:
            r1 = doc.GoTo(wdGoToPage, wdGoToFirst, 1)
            p1_start = r1.Start
            r2 = doc.GoTo(wdGoToPage, wdGoToNext, 1)
            p2_start = r2.Start
        except Exception:
            p1_start = 0
            p2_start = doc.Content.End
        if p2_start <= p1_start:
            p2_start = min(p1_start + 8000, doc.Content.End)
        cover_range = doc.Range(p1_start, min(p2_start, doc.Content.End))
        cover_text = cover_range.Text or ""
        # 检查 作者、审核、批准、日期
        next_kw = ["审核", "批准", "日期"]
        author = _extract_field_value(cover_text, "作者", next_kw)
        next_kw = ["批准", "日期"]
        review = _extract_field_value(cover_text, "审核", next_kw)
        next_kw = ["日期"]
        approve = _extract_field_value(cover_text, "批准", next_kw)
        date_val = _extract_field_value(cover_text, "日期", [])
        if _is_value_empty(author):
            return False, "封面页「作者」为空"
        if _is_value_empty(review):
            return False, "封面页「审核」为空"
        if _is_value_empty(approve):
            return False, "封面页「批准」为空"
        if _is_value_empty(date_val):
            return False, "封面页「日期」为空"
        return True, "封面签字检查通过（作者、审核、批准、日期均已填写）"
    except Exception as e:
        return False, f"签字检查异常: {e}"
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# 草稿水印关键词（页眉页脚 Shape 文本包含任一则视为非正式）
DRAFT_WATERMARK_KEYWORDS = ("草稿", "DRAFT", "内部", "初稿", "待审核", "副本")


def _has_draft_watermark(doc):
    """检测页眉页脚中是否存在草稿类水印。"""
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):  # 页眉
                try:
                    hf = sec.Headers(hf_type)
                    if not hf.Exists:
                        continue
                    for i in range(1, hf.Shapes.Count + 1):
                        try:
                            sh = hf.Shapes(i)
                            if sh.TextFrame.HasText:
                                t = (sh.TextFrame.TextRange.Text or "").upper()
                                for kw in DRAFT_WATERMARK_KEYWORDS:
                                    if kw.upper() in t:
                                        return True
                        except Exception:
                            pass
                except Exception:
                    pass
            for hf_type in (1, 2, 3):  # 页脚
                try:
                    hf = sec.Footers(hf_type)
                    if not hf.Exists:
                        continue
                    for i in range(1, hf.Shapes.Count + 1):
                        try:
                            sh = hf.Shapes(i)
                            if sh.TextFrame.HasText:
                                t = (sh.TextFrame.TextRange.Text or "").upper()
                                for kw in DRAFT_WATERMARK_KEYWORDS:
                                    if kw.upper() in t:
                                        return True
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _has_unupdated_fields(doc):
    """检测是否存在未更新或占位符的域。"""
    try:
        for i in range(1, doc.Fields.Count + 1):
            f = doc.Fields(i)
            try:
                if f.Result.Text.strip() in ("", "0", "错误!") or "{" in (f.Result.Text or ""):
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def _remove_draft_watermark_shapes(doc):
    """删除页眉页脚中包含草稿关键词的 Shape（水印）。"""
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):
                for hf_getter in (sec.Headers, sec.Footers):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        to_delete = []
                        for i in range(1, hf.Shapes.Count + 1):
                            try:
                                sh = hf.Shapes(i)
                                if sh.TextFrame.HasText:
                                    t = (sh.TextFrame.TextRange.Text or "").upper()
                                    for kw in DRAFT_WATERMARK_KEYWORDS:
                                        if kw.upper() in t:
                                            to_delete.append(i)
                                            break
                            except Exception:
                                pass
                        for i in reversed(to_delete):
                            try:
                                hf.Shapes(i).Delete()
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass


def _has_page_number(doc):
    """检测文档是否在页眉/页脚中已有页码域（PAGE）。首页/封面可不含页码，仅要求至少有一处存在 PAGE。"""
    try:
        for si in range(1, int(doc.Sections.Count) + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):
                for hf_getter in (sec.Headers, sec.Footers):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        r = hf.Range
                        if r is None or r.Fields.Count == 0:
                            continue
                        for fi in range(1, r.Fields.Count + 1):
                            try:
                                f = r.Fields(fi)
                                if int(f.Type) == wdFieldPage:
                                    return True
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass
    return False


def _ensure_page_numbers(doc):
    """为缺少页码的节在主页脚中插入 PAGE 域。"""
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            try:
                foot = sec.Footers(wdHeaderFooterPrimary)
                if not foot.Exists:
                    continue
                r = foot.Range
                has_page = False
                for fi in range(1, r.Fields.Count + 1):
                    try:
                        if int(r.Fields(fi).Type) == wdFieldPage:
                            has_page = True
                            break
                    except Exception:
                        pass
                if not has_page:
                    r.Collapse(0)  # 0 = wdCollapseEnd
                    r.Fields.Add(r, wdFieldPage, "", False)
            except Exception:
                pass
    except Exception:
        pass


def _has_headers_footers_issue(doc):
    """检测页眉页脚是否异常：多节时首节有页脚而后续节缺失主页脚。首页/封面无页脚不报。"""
    try:
        n = int(doc.Sections.Count)
        if n <= 1:
            return False
        first_has_footer = False
        try:
            f = doc.Sections(1).Footers(wdHeaderFooterPrimary)
            if f.Exists and f.Range and f.Range.Text.strip():
                first_has_footer = True
        except Exception:
            pass
        if not first_has_footer:
            return False
        for si in range(2, n + 1):
            try:
                f = doc.Sections(si).Footers(wdHeaderFooterPrimary)
                if not f.Exists or not f.Range or not (f.Range.Text or "").strip():
                    return True
            except Exception:
                return True
    except Exception:
        pass
    return False


def _ensure_headers_footers_consistent(doc):
    """若首节有页脚而后续节无内容，且首节为纯文本（无域）时复制到后续节；否则由 _ensure_page_numbers 补页码。"""
    try:
        n = int(doc.Sections.Count)
        if n <= 1:
            return
        try:
            src = doc.Sections(1).Footers(wdHeaderFooterPrimary)
            if not src.Exists or not src.Range:
                return
            if getattr(src.Range, "Fields", None) and int(src.Range.Fields.Count) > 0:
                return
            src_text = (src.Range.Text or "").strip()
            if not src_text:
                return
        except Exception:
            return
        for si in range(2, n + 1):
            try:
                dst = doc.Sections(si).Footers(wdHeaderFooterPrimary)
                if not dst.Exists:
                    continue
                dst_r = dst.Range
                if (dst_r.Text or "").strip():
                    continue
                dst_r.Text = src_text
                try:
                    dst_r.Fields.Update()
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass


def _has_toc_error_or_unupdated(doc):
    """检测是否存在目录域未更新或显示错误。"""
    try:
        for i in range(1, doc.Fields.Count + 1):
            f = doc.Fields(i)
            try:
                t = int(getattr(f, "Type", -1))
                if t == wdFieldTOC:
                    res = (f.Result.Text or "").strip()
                    if not res or "错误" in res or "Error" in res.lower():
                        return True
            except Exception:
                pass
    except Exception:
        pass
    n_toc, toc_coll = _tables_of_contents_safe(doc)
    if n_toc and toc_coll:
        try:
            for i in range(1, n_toc + 1):
                try:
                    toc = toc_coll(i)
                    r = (toc.Range.Text or "").strip()
                    if not r or "错误" in r or "Error" in r.lower():
                        return True
                except Exception:
                    pass
        except Exception:
            pass
    return False


def _ensure_toc_updated(doc):
    """更新所有目录域与域结果。"""
    n_toc, toc_coll = _tables_of_contents_safe(doc)
    if n_toc and toc_coll:
        try:
            for i in range(1, n_toc + 1):
                try:
                    toc_coll(i).Update()
                except Exception:
                    pass
        except Exception:
            pass
    try:
        doc.Fields.Update()
    except Exception:
        pass


def _has_tables_not_fit(doc):
    """检测是否存在行高固定、可能导致内容显示不全的表格。"""
    try:
        for i in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(i)
            try:
                for ri in range(1, tbl.Rows.Count + 1):
                    row = tbl.Rows(ri)
                    hr = int(getattr(row, "HeightRule", -1))
                    if hr == wdRowHeightExactly:
                        return True
            except Exception:
                pass
    except Exception:
        pass
    return False


def _auto_fit_tables(doc):
    """保守修复：仅取消固定行高，避免重排表格导致字段位置错位。"""
    try:
        for i in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(i)
            try:
                nr = int(tbl.Rows.Count)
                for ri in range(1, nr + 1):
                    try:
                        tbl.Rows(ri).HeightRule = wdRowHeightAuto
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass


def auto_fix_formal_word(doc_path, save_path=None):
    """
    正式性检查失败后自动修复。执行顺序：表格排版 → 统一表格框线 → 去高亮/统一字体/去水印/统一宋体 → 目录更新 → 接受修订 → 删除批注 → 保存。
    非风险矩阵中的标红、标黄、标绿会被修复为黑色；风险矩阵内保留。页码、页眉页脚仅检查不修改。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=False, AddToRecentFiles=False)
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        saved_risk_matrices = _save_risk_matrix_formats(doc)
        _auto_fit_tables(doc)
        _normalize_table_borders(doc)
        wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        for stType in wdStoryTypes:
            try:
                r = doc.StoryRanges(stType)
                while r is not None:
                    try:
                        r.HighlightColorIndex = wdNoHighlight
                        r.Font.Color = wdColorBlack
                        r.Font.Name = FONT_NAME_SONG
                        r.Font.Hidden = False
                    except Exception:
                        pass
                    try:
                        r = r.NextStoryRange
                    except Exception:
                        break
            except Exception:
                pass
        try:
            doc.Content.HighlightColorIndex = wdNoHighlight
            doc.Content.Font.Color = wdColorBlack
            doc.Content.Font.Name = FONT_NAME_SONG
            doc.Content.Font.Hidden = False
        except Exception:
            pass
        _remove_draft_watermark_shapes(doc)
        _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True)
        _restore_risk_matrix_formats(doc, saved_risk_matrices)
        _ensure_toc_updated(doc)
        _accept_all_revisions_in_document(doc)
        try:
            doc.DeleteAllComments()
        except Exception:
            pass
        _save_doc(doc, path, save_path)
        return True
    except Exception as e:
        try:
            if getattr(e, "args", (None,))[0] in (COM_E_EXCEPTION, COM_E_FAIL):
                raise RuntimeError(
                    "文档处理时发生意外，请确认 WPS 已安装、文档未被占用，或稍后重试。"
                ) from e
        except RuntimeError:
            raise
        raise
    finally:
        if doc:
            try:
                doc.Saved = True
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def check_formal_document(
    doc_path,
    check_highlight=True,
    check_revisions=True,
    check_comments=True,
    check_font_color=True,
    check_hidden_text=True,
    check_draft_watermark=True,
    check_unupdated_fields=False,
    check_digital_signature=False,
    check_page_number=True,
    check_headers_footers=True,
    check_toc=True,
    check_tables_fit=True,
):
    """
    第一步：正式性检查（覆盖所有判定点，含页眉页脚、页码、目录、表格排版）。
    返回 (passed: bool, issues: list[str])。
    """
    if win32com is None:
        return True, []
    path = os.path.abspath(doc_path)
    if not os.path.isfile(path):
        return False, ["文件不存在"]
    issues = []
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=True, AddToRecentFiles=False)

        if check_revisions:
            n = int(doc.Revisions.Count)
            if n > 0:
                issues.append(f"存在 {n} 处修订（跟踪的更改）")

        if check_comments:
            n = int(doc.Comments.Count)
            if n > 0:
                issues.append(f"存在 {n} 条批注/评论")

        if check_digital_signature:
            try:
                sigs = doc.Signatures
                if int(sigs.Count) > 0:
                    for j in range(1, sigs.Count + 1):
                        s = sigs.Item(j)
                        if not getattr(s, "IsSigned", True):
                            issues.append("存在未签署的签名行")
                            break
                        if getattr(s, "IsValid", True) is False:
                            issues.append("存在无效的数字签名")
                            break
            except Exception:
                pass

        if check_draft_watermark and _has_draft_watermark(doc):
            issues.append("存在草稿/内部水印")

        if check_unupdated_fields and _has_unupdated_fields(doc):
            issues.append("存在未更新或错误的域")

        if check_page_number:
            try:
                n_pages = int(doc.ComputeStatistics(2))  # 2 = wdStatisticPages
                if n_pages > 1 and not _has_page_number(doc):
                    issues.append("多页文档未设置页码（首页/封面可不设，其余页需有页码域）")
            except Exception:
                pass

        if check_headers_footers and _has_headers_footers_issue(doc):
            issues.append("页眉页脚不一致或后续节缺少页脚（首页/封面可不设页脚）")

        if check_toc and _has_toc_error_or_unupdated(doc):
            issues.append("目录未更新或存在错误")

        if check_tables_fit and _has_tables_not_fit(doc):
            issues.append("存在表格行高固定可能导致内容显示不全")

        if check_highlight or check_font_color or check_hidden_text:
            wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
            has_highlight = False
            has_non_black = False
            has_hidden = False
            for stType in wdStoryTypes:
                try:
                    r = doc.StoryRanges(stType)
                    while r is not None:
                        try:
                            try:
                                npar = int(r.Paragraphs.Count)
                            except Exception:
                                npar = 0
                            if npar >= 1:
                                for pi in range(1, npar + 1):
                                    try:
                                        pr = r.Paragraphs(pi).Range
                                        if pr.Tables.Count > 0:
                                            try:
                                                if _is_risk_matrix_table(pr.Tables(1)):
                                                    continue
                                            except Exception:
                                                pass
                                        if check_highlight and pr.HighlightColorIndex != wdNoHighlight:
                                            has_highlight = True
                                        if check_font_color:
                                            c = getattr(pr.Font, "Color", wdColorBlack)
                                            if c != wdColorBlack and c != wdColorAutomatic:
                                                has_non_black = True
                                        if check_hidden_text and getattr(pr.Font, "Hidden", 0) != 0:
                                            has_hidden = True
                                        if has_highlight and has_non_black and (not check_hidden_text or has_hidden):
                                            break
                                    except Exception:
                                        pass
                            else:
                                if not _range_in_risk_matrix(doc, r.Start, r.End):
                                    if check_highlight and r.HighlightColorIndex != wdNoHighlight:
                                        has_highlight = True
                                    if check_font_color:
                                        c = getattr(r.Font, "Color", wdColorBlack)
                                        if c != wdColorBlack and c != wdColorAutomatic:
                                            has_non_black = True
                                    if check_hidden_text and getattr(r.Font, "Hidden", 0) != 0:
                                        has_hidden = True
                            if has_highlight and has_non_black and (not check_hidden_text or has_hidden):
                                break
                        except Exception:
                            pass
                        try:
                            r = r.NextStoryRange
                        except Exception:
                            break
                except Exception:
                    pass
                if has_highlight and has_non_black and (not check_hidden_text or has_hidden):
                    break
            if check_highlight and has_highlight:
                issues.append("存在标黄/高亮")
            if check_font_color and has_non_black:
                issues.append("存在非黑色字体")
            if check_hidden_text and has_hidden:
                issues.append("存在隐藏文字")

        passed = len(issues) == 0
        return passed, issues
    except Exception as e:
        try:
            if getattr(e, "args", (None,))[0] in (COM_E_EXCEPTION, COM_E_FAIL):
                return False, ["正式性检查时发生意外，请确认 WPS 已安装且文档未被占用，或稍后重试。"]
        except Exception:
            pass
        return False, [f"正式性检查异常: {e}"]
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def has_revisions(doc_path):
    """检查文档是否存在修订（跟踪更改）。"""
    if win32com is None:
        return False, 0
    path = os.path.abspath(doc_path)
    if not os.path.isfile(path):
        return False, 0
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=True, AddToRecentFiles=False)
        revs = doc.Revisions
        count = int(revs.Count)
        return count > 0, count
    except Exception:
        return False, 0
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def ensure_font_black(doc_path, save_path=None, remove_highlights=True):
    """
    将文档中所有文字颜色统一为黑色，去除标黄高亮，符合正式文档要求。
    save_path 为 None 时覆盖原文件。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = save_path or path
    save_path = os.path.abspath(save_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=False, AddToRecentFiles=False)
        try:
            doc.Content.Font.Color = wdColorBlack
            doc.Content.Font.Name = FONT_NAME_SONG
            wdStoryTypes = (1, 7, 8, 9, 10, 11, 12)
            for st in wdStoryTypes:
                try:
                    r = doc.StoryRanges(st)
                    while r is not None:
                        try:
                            r.Font.Color = wdColorBlack
                            r.Font.Name = FONT_NAME_SONG
                            if remove_highlights:
                                r.HighlightColorIndex = wdNoHighlight
                        except Exception:
                            pass
                        try:
                            r = r.NextStoryRange
                        except Exception:
                            break
                except Exception:
                    pass
            if remove_highlights:
                doc.Content.HighlightColorIndex = wdNoHighlight
        except Exception:
            try:
                doc.Content.Font.Color = wdColorBlack
                doc.Content.Font.Name = FONT_NAME_SONG
                if remove_highlights:
                    doc.Content.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
        _save_doc(doc, path, save_path)
        return True
    finally:
        if doc:
            try:
                doc.Saved = True
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True):
    """
    逐段将字体统一为宋体（同一段落内字体一致），可选同时设为黑色、去高亮。
    风险矩阵表格内的段落不处理，保留标红/标黄/标绿。
    """
    try:
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                p = doc.Paragraphs(i)
                r = p.Range
                if r is None:
                    continue
                if r.Tables.Count > 0:
                    try:
                        if _is_risk_matrix_table(r.Tables(1)):
                            continue
                    except Exception:
                        pass
                r.Font.Name = FONT_NAME_SONG
                if set_black:
                    r.Font.Color = wdColorBlack
                if remove_highlight:
                    r.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
    except Exception:
        pass


def _accept_all_revisions_in_document(doc):
    """
    接受文档中所有修订，包括正文、目录页、页眉页脚、文本框等。

    关键：必须先关闭 TrackRevisions，否则后续 Fields.Update / TOC.Update
    会把目录重新生成的内容当作新修订记录下来，导致修订永远清不完。
    """
    # ===== 0) 关闭修订跟踪与修订显示，防止后续操作产生新修订 =====
    try:
        doc.TrackRevisions = False
    except Exception:
        pass
    try:
        doc.ShowRevisions = False
    except Exception:
        pass

    # ===== 1) Document.AcceptAllRevisions（全文档含脚注/目录等） =====
    try:
        doc.AcceptAllRevisions()
    except Exception:
        pass
    # 再用 Revisions.AcceptAll 兜底
    try:
        if int(doc.Revisions.Count) > 0:
            doc.Revisions.AcceptAll()
    except Exception:
        pass

    # ===== 2) 逐段接受（含目录页每一段） =====
    try:
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                r = doc.Paragraphs(i).Range
                if r is not None and int(r.Revisions.Count) > 0:
                    r.Revisions.AcceptAll()
            except Exception:
                pass
    except Exception:
        pass

    # ===== 3) 全文 Content =====
    try:
        r = doc.Content
        if r is not None and int(r.Revisions.Count) > 0:
            r.Revisions.AcceptAll()
    except Exception:
        pass

    # ===== 4) 按 Story 类型 =====
    for stType in range(1, 12):
        try:
            r = doc.StoryRanges(stType)
            while r is not None:
                try:
                    if int(r.Revisions.Count) > 0:
                        r.Revisions.AcceptAll()
                except Exception:
                    pass
                try:
                    r = r.NextStoryRange
                except Exception:
                    break
        except Exception:
            pass

    # ===== 5) 按节 =====
    try:
        for si in range(1, doc.Sections.Count + 1):
            try:
                r = doc.Sections(si).Range
                if r is not None and int(r.Revisions.Count) > 0:
                    r.Revisions.AcceptAll()
            except Exception:
                pass
    except Exception:
        pass

    # ===== 6) 页眉页脚 =====
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):
                for hf_getter in (sec.Headers, sec.Footers):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        r = hf.Range
                        if r is not None and int(r.Revisions.Count) > 0:
                            r.Revisions.AcceptAll()
                    except Exception:
                        pass
    except Exception:
        pass

    # ===== 7) Shape 文本框（页眉页脚 + 主文档） =====
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            for hf_getter in (sec.Headers, sec.Footers):
                for hf_type in (1, 2, 3):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        for i in range(1, hf.Shapes.Count + 1):
                            try:
                                sh = hf.Shapes(i)
                                if sh.TextFrame.HasText and sh.TextFrame.TextRange is not None:
                                    r = sh.TextFrame.TextRange
                                    if int(r.Revisions.Count) > 0:
                                        r.Revisions.AcceptAll()
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass
    try:
        for i in range(1, doc.Shapes.Count + 1):
            try:
                sh = doc.Shapes(i)
                if sh.TextFrame.HasText and sh.TextFrame.TextRange is not None:
                    r = sh.TextFrame.TextRange
                    if int(r.Revisions.Count) > 0:
                        r.Revisions.AcceptAll()
            except Exception:
                pass
    except Exception:
        pass

    # ===== 8) 文档级扫尾 =====
    try:
        if int(doc.Revisions.Count) > 0:
            doc.Revisions.AcceptAll()
    except Exception:
        pass

    # ===== 9) 逐条接受残留修订 =====
    try:
        for _ in range(int(doc.Revisions.Count) + 1):
            if int(doc.Revisions.Count) == 0:
                break
            try:
                doc.Revisions(1).Accept()
            except Exception:
                break
    except Exception:
        pass

    # ===== 10) 更新目录（TrackRevisions 已关闭，不会产生新修订） =====
    n_toc, toc_coll = _tables_of_contents_safe(doc)
    if n_toc and toc_coll:
        try:
            for i in range(1, n_toc + 1):
                try:
                    toc_coll(i).Update()
                except Exception:
                    pass
        except Exception:
            pass
    try:
        doc.Fields.Update()
    except Exception:
        pass

    # ===== 11) 更新目录后再次清理（以防万一） =====
    try:
        if int(doc.Revisions.Count) > 0:
            doc.AcceptAllRevisions()
    except Exception:
        pass
    try:
        if int(doc.Revisions.Count) > 0:
            doc.Revisions.AcceptAll()
    except Exception:
        pass
    try:
        for _ in range(int(doc.Revisions.Count) + 1):
            if int(doc.Revisions.Count) == 0:
                break
            try:
                doc.Revisions(1).Accept()
            except Exception:
                break
    except Exception:
        pass


def accept_all_revisions_and_save(doc_path, save_path=None, ensure_font_black_too=True):
    """
    接受文档中所有修订，去除标黄高亮，可选统一字体黑色宋体，并保存。
    save_path 为 None 时覆盖原文件。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _com_call(word.Documents.Open, path, ReadOnly=False, AddToRecentFiles=False)
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        _accept_all_revisions_in_document(doc)
        _normalize_table_borders(doc)
        saved_risk_matrices = _save_risk_matrix_formats(doc)
        wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        for stType in wdStoryTypes:
            try:
                r = doc.StoryRanges(stType)
                while r is not None:
                    try:
                        r.HighlightColorIndex = wdNoHighlight
                        if ensure_font_black_too:
                            r.Font.Color = wdColorBlack
                            r.Font.Name = FONT_NAME_SONG
                    except Exception:
                        pass
                    try:
                        r = r.NextStoryRange
                    except Exception:
                        break
            except Exception:
                pass
        if ensure_font_black_too:
            try:
                doc.Content.Font.Color = wdColorBlack
                doc.Content.Font.Name = FONT_NAME_SONG
                doc.Content.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
        if ensure_font_black_too:
            _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True)
        _restore_risk_matrix_formats(doc, saved_risk_matrices)
        _save_doc(doc, path, save_path)
        return True
    except Exception as e:
        try:
            if getattr(e, "args", (None,))[0] in (COM_E_EXCEPTION, COM_E_FAIL):
                raise RuntimeError(
                    "文档处理时发生意外，请确认 WPS 已安装、文档未被占用，或稍后重试。"
                ) from e
        except RuntimeError:
            raise
        raise
    finally:
        if doc:
            try:
                doc.Saved = True
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def print_word_document(doc_path, printer_name=None, copies=1):
    """
    使用 WPS 文字 / Word 打印文档。
    打印后等待一段时间再关闭，以便打印窗口能弹出且多份文件时下一份能正常弹窗。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    old_printer = None
    try:
        word = _get_word_app(visible=False)
        if printer_name:
            old_printer = word.ActivePrinter
            word.ActivePrinter = printer_name
        doc = _com_call(word.Documents.Open, path, ReadOnly=True, AddToRecentFiles=False)
        doc.PrintOut(
            Background=False,
            Append=False,
            Range=0,
            Item=0,
            Copies=int(copies),
            Collate=True,
        )
        # 等待打印窗口弹出并允许用户操作，避免立即关闭导致第二份不弹窗
        time.sleep(3)
        return True
    except Exception as e:
        try:
            if getattr(e, "args", (None,))[0] in (COM_E_EXCEPTION, COM_E_FAIL):
                raise RuntimeError(
                    "文档处理时发生意外，请确认 WPS 已安装、文档未被占用，或稍后重试。"
                ) from e
        except RuntimeError:
            raise
        raise
    finally:
        if word and old_printer:
            try:
                word.ActivePrinter = old_printer
            except Exception:
                pass
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        # 给 WPS/Word 时间完全退出，再处理下一份时新建实例不会冲突
        time.sleep(1)
        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
