# -*- coding: utf-8 -*-
"""
Word 文档处理：检查封面签字、接受所有修订、统一字体为黑色、打印
依赖 Windows + 已安装的 WPS 文字 或 Microsoft Word，使用 COM (pywin32)
默认使用 WPS（KWPS.Application）；环境变量 USE_OFFICE=1 时使用 Word。
"""
import os
import re
import shutil
import tempfile
import time
import uuid
import logging
import subprocess
from typing import Optional

try:
    import win32com.client
    import pythoncom
    import pywintypes
    import win32process
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None
    win32process = None

from config import get_word_progid, use_wps_runtime
from doc_handlers.office_pid import (
    hwnd_chain_to_pid,
    refresh_pid_after_doc_open,
    resolve_office_app_pid,
    tasklist_pids_for_images,
)

logger = logging.getLogger("aiprintword.word")


def convert_doc_to_docx(src_path: str, dst_path: Optional[str] = None) -> str:
    """
    将 .doc 转为 .docx（用于启用 XML 表格计数与更稳定的后续处理）。
    失败时抛异常；调用方可自行回退按原文件处理。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    src = os.path.abspath(src_path)
    if not os.path.isfile(src):
        raise FileNotFoundError(src)
    ext = os.path.splitext(src)[1].lower()
    if ext != ".doc":
        return src
    if dst_path:
        dst = os.path.abspath(dst_path)
    else:
        d = os.path.dirname(src) or os.getcwd()
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            d = os.path.dirname(src)
        stem = os.path.splitext(os.path.basename(src))[0]
        dst = os.path.join(d, f"{stem}._aiprint_{uuid.uuid4().hex[:10]}.docx")
        dst = os.path.abspath(dst)
    word = None
    doc = None
    try:
        word = _get_word_app(visible=False)
        doc = _open_word_document(word, src, read_only=False)
        # Word 常量：wdFormatXMLDocument = 12
        try:
            _com_call(doc.SaveAs2, dst, FileFormat=12)
        except Exception:
            _com_call(doc.SaveAs, dst)
        return dst
    finally:
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)


def docx_table_count_hint(path: str) -> int:
    """
    对 .docx/.docm 快速统计 word/document.xml 中 w:tbl 数量（不启动 Word）。
    失败或非 OOXML 时返回 -1。
    """
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        p = os.path.abspath(path)
        ext = os.path.splitext(p)[1].lower()
        if ext not in (".docx", ".docm"):
            return -1
        with zipfile.ZipFile(p, "r") as zf:
            data = zf.read("word/document.xml")
        root = ET.fromstring(data)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        return len(root.findall(".//w:tbl", ns))
    except Exception:
        return -1


def docx_table_row_count_hint(path: str) -> int:
    """
    对 .docx/.docm 快速统计 word/document.xml 中所有表格行 w:tr 总数（含嵌套表，不启动 Word）。
    失败或非 OOXML 时返回 -1。
    """
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        p = os.path.abspath(path)
        ext = os.path.splitext(p)[1].lower()
        if ext not in (".docx", ".docm"):
            return -1
        with zipfile.ZipFile(p, "r") as zf:
            data = zf.read("word/document.xml")
        root = ET.fromstring(data)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        return len(root.findall(".//w:tbl//w:tr", ns))
    except Exception:
        return -1


def word_lite_repair_many_tables_save(
    doc_path,
    save_path=None,
    change_notes=None,
    full_change_log=None,
):
    """
    表格行数多、处理耗时场景：仅接受修订、去标黄/高亮、按字体策略统一、更新目录与页码后保存。
    不执行全文正式性检查、删删除线、表格自适应、单元格内图片缩放、水印清理、图片风险对比等重步骤。

    为控制 COM 调用量：不做风险矩阵底纹快照（表格极多时逐格读写很慢）；不按段落再扫一遍字体
   （StoryRanges + Content 已覆盖主故事）；接受修订阶段不重复执行内置全量域更新（由后续
    _ensure_toc_updated 统一处理）。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    backup_path = None
    try:
        try:
            fd, backup_path = _mkstemp_word_backup()
            os.close(fd)
            shutil.copyfile(path, backup_path)
        except Exception:
            backup_path = None

        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=False)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        n_rev = 0
        try:
            n_rev = int(doc.Revisions.Count)
        except Exception:
            pass
        _accept_all_revisions_in_document(doc, skip_builtin_toc_pass=True)
        if change_notes is not None and n_rev > 0:
            change_notes.append(f"已接受全部修订（约 {n_rev} 处跟踪更改）")
        wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        for stType in wdStoryTypes:
            try:
                r = doc.StoryRanges(stType)
                while r is not None:
                    try:
                        r.HighlightColorIndex = wdNoHighlight
                        r.Font.Color = wdColorBlack
                        _apply_font_profile_to_range(r)
                    except Exception:
                        pass
                    try:
                        r = r.NextStoryRange
                    except Exception:
                        break
            except Exception:
                pass
        try:
            doc.Content.Font.Color = wdColorBlack
            _apply_font_profile_to_range(doc.Content)
            doc.Content.HighlightColorIndex = wdNoHighlight
        except Exception:
            pass
        n_com = 0
        try:
            n_com = int(doc.Comments.Count)
        except Exception:
            pass
        _delete_all_comments_robust(doc)
        if change_notes is not None and n_com > 0:
            change_notes.append(f"已删除 {n_com} 条批注")
        if change_notes is not None:
            change_notes.append(
                "【多表格轻量】已去除高亮与标黄、将非黑色字体改为黑色并统一为「"
                + _word_font_profile_label()
                + "」（为缩短耗时未做风险矩阵颜色保留；已跳过全文正式性检查、删删除线、"
                "表格自适应与图片缩放等；若须保留风险矩阵色，请提高「表格行轻量阈值」使该文档走完整流程）"
            )
        _ensure_toc_updated(doc)
        _sync_page_numbers_after_edit(doc)
        _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path)
        _finalize_no_track_changes_state(doc)
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
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
        if backup_path:
            try:
                if os.path.isfile(backup_path):
                    os.remove(backup_path)
            except Exception:
                pass

# 记录最近创建的 Word 应用 PID，供超时/取消时强制结束（避免 COM 永久卡死）
_LAST_WORD_PID = None

# 整批连续 Word：同一线程内复用 Application，仅开关文档（run_batch 内开启；遇 Excel/PDF 或批末关闭）
_word_batch_session_active = False
_word_batch_cached_app = None


def word_batch_session_active() -> bool:
    return bool(_word_batch_session_active)


def word_batch_session_begin() -> None:
    """在即将处理 Word 文件前调用；已在会话中则为空操作。"""
    global _word_batch_session_active
    _word_batch_session_active = True


def word_batch_session_end() -> None:
    """结束批量 Word 会话：关闭残留文档并退出进程（无会话则为空操作）。"""
    global _word_batch_session_active, _word_batch_cached_app, _LAST_WORD_PID
    if not _word_batch_session_active:
        return
    _word_batch_session_active = False
    w = _word_batch_cached_app
    _word_batch_cached_app = None
    if w:
        try:
            for _ in range(80):
                try:
                    if int(w.Documents.Count) <= 0:
                        break
                except Exception:
                    break
                try:
                    w.Documents(1).Close(SaveChanges=False)
                except Exception:
                    break
        except Exception:
            pass
        try:
            w.Quit()
        except Exception:
            pass
    _LAST_WORD_PID = None
    time.sleep(1)
    if pythoncom:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _word_batch_drop_cached_app() -> None:
    """进程已被 kill 或 COM 失效时丢弃缓存，下次将重新 Dispatch。"""
    global _word_batch_cached_app, _LAST_WORD_PID
    _word_batch_cached_app = None
    _LAST_WORD_PID = None


def _word_post_document_cleanup(word, doc, *, mark_doc_saved=False):
    """
    关闭当前文档；批量复用模式下不 Quit Word。
    mark_doc_saved：规范化保存路径在 Close 前标记已保存，避免提示。
    """
    if doc:
        try:
            if mark_doc_saved:
                doc.Saved = True
            doc.Close(SaveChanges=False)
        except Exception:
            pass
    if word_batch_session_active():
        time.sleep(0.25)
        return
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


def _word_host_image_names():
    return ("wps.exe",) if use_wps_runtime() else ("WINWORD.EXE",)


def _try_get_app_pid(app):
    """从 COM Application 的窗口链推导 PID（WPS 常 Hwnd=0，需多源尝试）。"""
    return hwnd_chain_to_pid(app, win32process)


def kill_last_word_app(reason="timeout_or_cancel"):
    """强制结束最近创建的 Word/WPS Writer 进程（仅本工具创建的那个 PID）。"""
    global _LAST_WORD_PID
    pid = _LAST_WORD_PID
    if not pid:
        logger.warning(
            "kill_last_word_app skipped: no PID (reason=%s)；单文件超时/跳过可能无法打断卡死的 COM",
            reason,
        )
        return False
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        logger.warning("killed word pid=%s reason=%s", pid, reason)
        _word_batch_drop_cached_app()
        return True
    except Exception:
        return False

# 修改明细文件中单条文本预览长度（删除线等）
_MAX_CHANGE_TEXT_PREVIEW = 100
_MAX_FILE_DETAIL_PREVIEW = 800


def _sanitize_change_preview(text, max_len=_MAX_CHANGE_TEXT_PREVIEW):
    """用于结果展示的片段：去控制符、压缩空白、截断。"""
    if not text:
        return ""
    try:
        t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(text))
    except Exception:
        t = str(text)
    t = t.replace("\r", " ").replace("\n", " ").strip()
    if len(t) > max_len:
        return t[:max_len] + "…"
    return t


# Word 模块级开关：由 apply_resolved_word_base_settings() 从库内配置 / .env / 默认值加载；
# 每批任务开始前 batch_print 会再次调用，随后 set_runtime_options 可按单次请求覆盖。
WORD_CONTENT_PRESERVE = True
WORD_IMAGE_RISK_GUARD = True
WORD_PRESERVE_LINKED_IMAGES = False
WORD_HEADER_FOOTER_LAYOUT_FIX = True
WORD_STEP_TIMEOUT_SEC = 3600.0
WORD_SKIP_FILE_ON_TIMEOUT = True
# 页数保护备份目录：SSD 路径可减少慢盘与杀毒实时扫描干扰（留空用系统 TEMP）
WORD_BACKUP_TEMP_DIR = ""


class WordStepTimeout(RuntimeError):
    pass


def _mkstemp_word_backup(suffix=".docx"):
    """页数保护等临时备份；目录由 WORD_BACKUP_TEMP_DIR / 配置项指定。"""
    d = (WORD_BACKUP_TEMP_DIR or "").strip()
    if d:
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            d = ""
    kw = {"prefix": "aiprintword_wordbak_", "suffix": suffix}
    if d:
        kw["dir"] = d
    return tempfile.mkstemp(**kw)


class FileAbortRequested(Exception):
    """单文件总超时/取消/手动跳过：在下一 COM 调用前中止（与看门狗 kill 配合）。"""


# 由 batch_print 在每文件开始时设置，结束时清空；用于在 _com_call 等处协作退出。
WORD_FILE_ABORT_EVENT = None


def set_word_file_abort_event(ev):
    """ev 为 threading.Event；置位后下一 COM 包装调用将抛出 FileAbortRequested。"""
    global WORD_FILE_ABORT_EVENT
    WORD_FILE_ABORT_EVENT = ev


def clear_word_file_abort_event():
    global WORD_FILE_ABORT_EVENT
    WORD_FILE_ABORT_EVENT = None


def _check_file_abort():
    ev = WORD_FILE_ABORT_EVENT
    if ev is not None and ev.is_set():
        raise FileAbortRequested("单文件总耗时已超过限制，已中止处理")


def _step_timed_out(t0, timeout_sec=None):
    try:
        limit = float(timeout_sec) if timeout_sec is not None else float(WORD_STEP_TIMEOUT_SEC)
    except Exception:
        limit = 3600.0
    if limit <= 0:
        return False
    return (time.perf_counter() - float(t0)) >= limit


def _raise_or_mark_timeout(step_name):
    if WORD_SKIP_FILE_ON_TIMEOUT:
        raise WordStepTimeout(f"【超时跳过】步骤「{step_name}」超过 {int(WORD_STEP_TIMEOUT_SEC)} 秒，已标记该文件需人工处理")
    raise WordStepTimeout(f"【超时】步骤「{step_name}」超过 {int(WORD_STEP_TIMEOUT_SEC)} 秒")


def _snapshot_visual_objects(doc):
    """
    采样文档可视对象数量，用于检测处理后是否发生疑似图片/图形丢失。
    仅做保守计数，不依赖具体 Shape 类型常量，兼容 WPS/Word 差异。
    """
    out = {
        "body_inline": 0,
        "body_shapes": 0,
        "hf_shapes": 0,
        "body_inline_pics": 0,
        "body_shape_pics": 0,
        "hf_shape_pics": 0,
    }
    try:
        inlines = getattr(doc, "InlineShapes", None)
        out["body_inline"] = int(inlines.Count)
        for i in range(1, out["body_inline"] + 1):
            try:
                ish = inlines(i)
                t = int(getattr(ish, "Type", -1))
                # 常见图片类型：wdInlineShapePicture(3), wdInlineShapeLinkedPicture(4)
                if t in (3, 4):
                    out["body_inline_pics"] += 1
                    continue
                # 兜底：能访问 PictureFormat 视为图片对象
                try:
                    _ = ish.PictureFormat
                    out["body_inline_pics"] += 1
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        shapes = getattr(doc, "Shapes", None)
        out["body_shapes"] = int(shapes.Count)
        for i in range(1, out["body_shapes"] + 1):
            try:
                sh = shapes(i)
                st = int(getattr(sh, "Type", -1))
                # 常见图片类型：msoLinkedPicture(11), msoPicture(13)
                if st in (11, 13):
                    out["body_shape_pics"] += 1
                    continue
                try:
                    _ = sh.PictureFormat
                    out["body_shape_pics"] += 1
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        for si in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):
                for hf_getter in (sec.Headers, sec.Footers):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        out["hf_shapes"] += int(hf.Shapes.Count)
                        for i in range(1, int(hf.Shapes.Count) + 1):
                            try:
                                sh = hf.Shapes(i)
                                st = int(getattr(sh, "Type", -1))
                                if st in (11, 13):
                                    out["hf_shape_pics"] += 1
                                    continue
                                try:
                                    _ = sh.PictureFormat
                                    out["hf_shape_pics"] += 1
                                except Exception:
                                    pass
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass
    return out


def _visual_objects_lost(before, after):
    """
    正文内嵌图/浮动图数量若减少则视为高风险（含仅 1 张图被误删的情况）。
    说明：普通 Shape 可能是文本框/流程图，本函数只比较已识别的“图片类”计数。
    """
    try:
        b_pic = int(before.get("body_inline_pics", 0)) + int(before.get("body_shape_pics", 0))
        a_pic = int(after.get("body_inline_pics", 0)) + int(after.get("body_shape_pics", 0))
        if b_pic <= 0:
            return False
        if a_pic >= b_pic:
            return False
        drop = b_pic - a_pic
        logger.warning(
            "image-object-drop before(pics=%s) after(pics=%s) drop=%s details_before=%s details_after=%s",
            b_pic,
            a_pic,
            drop,
            {
                "inline_pics": before.get("body_inline_pics", 0),
                "shape_pics": before.get("body_shape_pics", 0),
                "hf_pics": before.get("hf_shape_pics", 0),
            },
            {
                "inline_pics": after.get("body_inline_pics", 0),
                "shape_pics": after.get("body_shape_pics", 0),
                "hf_pics": after.get("hf_shape_pics", 0),
            },
        )
        return True
    except Exception:
        return False
    return False


def _safe_accept_and_normalize_word(doc):
    """
    保真兜底规范化：仍执行“接受修订 + 基础规范化”，但避免高风险文本/对象清理。
    """
    try:
        doc.TrackRevisions = False
    except Exception:
        pass
    _accept_all_revisions_in_document(doc)
    _normalize_table_borders(doc)
    _ensure_toc_updated(doc)
    try:
        doc.DeleteAllComments()
    except Exception:
        pass


def _is_pdf_printer(printer_name):
    if not printer_name:
        return False
    s = str(printer_name).strip().lower()
    return "pdf" in s


def _desktop_pdf_path(src_path):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    stem = os.path.splitext(os.path.basename(src_path))[0]
    out = os.path.join(desktop, f"{stem}_printed.pdf")
    if not os.path.exists(out):
        return out
    ts = time.strftime("%Y%m%d_%H%M%S")
    return os.path.join(desktop, f"{stem}_printed_{ts}.pdf")


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
# 格式检查/规范化时默认字体
FONT_NAME_SONG = "SimSun"
FONT_NAME_LATIN = "Times New Roman"
# 统一字体策略：chinese=全文宋体 | english=全文 Times New Roman | mixed=西文 TNR + 中文宋体
WORD_FONT_PROFILE = "mixed"
# 页眉页脚：1= primary, 2= first page, 3= even
wdHeaderFooterPrimary = 1
# 域类型：页码、目录（保全链接图时用选择性更新，不全量 Fields.Update）
wdFieldPageRef = 25
wdFieldNumPages = 26
wdFieldSection = 28
wdFieldSectionPages = 30
wdFieldPage = 33
wdFieldTOC = 37
# 统计类型：页数（强制重算布局时常用）
wdStatisticPages = 2
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

# 总页数保护：处理前后 ComputeStatistics 页数不一致则中止保存并恢复备份（默认开启）
WORD_PRESERVE_PAGE_COUNT = True


def apply_resolved_word_base_settings():
    """
    从 runtime_settings（MySQL / 环境变量 / 默认值）刷新 Word 相关全局开关。
    在 run_batch 开头调用；单次请求仍可通过 set_runtime_options 覆盖。
    """
    from runtime_settings.resolve import get_setting

    global WORD_CONTENT_PRESERVE, WORD_IMAGE_RISK_GUARD, WORD_PRESERVE_LINKED_IMAGES
    global WORD_HEADER_FOOTER_LAYOUT_FIX, WORD_STEP_TIMEOUT_SEC, WORD_SKIP_FILE_ON_TIMEOUT
    global WORD_PRESERVE_PAGE_COUNT, WORD_BACKUP_TEMP_DIR
    try:
        WORD_CONTENT_PRESERVE = bool(get_setting("WORD_CONTENT_PRESERVE"))
        WORD_IMAGE_RISK_GUARD = bool(get_setting("WORD_IMAGE_RISK_GUARD"))
        WORD_PRESERVE_LINKED_IMAGES = bool(get_setting("WORD_PRESERVE_LINKED_IMAGES"))
        WORD_HEADER_FOOTER_LAYOUT_FIX = bool(get_setting("WORD_HEADER_FOOTER_LAYOUT_FIX"))
        WORD_STEP_TIMEOUT_SEC = float(get_setting("WORD_STEP_TIMEOUT_SEC"))
        WORD_SKIP_FILE_ON_TIMEOUT = bool(get_setting("WORD_SKIP_FILE_ON_TIMEOUT"))
        WORD_PRESERVE_PAGE_COUNT = bool(get_setting("WORD_PRESERVE_PAGE_COUNT"))
        WORD_BACKUP_TEMP_DIR = str(
            get_setting("AIPRINTWORD_WORD_BACKUP_TEMP_DIR") or ""
        ).strip()
    except Exception:
        pass


apply_resolved_word_base_settings()


def set_runtime_options(
    *,
    word_content_preserve=None,
    word_preserve_page_count=None,
    word_image_risk_guard=None,
    word_preserve_linked_images=None,
    word_step_timeout_sec=None,
    word_skip_file_on_timeout=None,
    word_header_footer_layout_fix=None,
    word_font_profile=None,
):
    """按请求级别动态设置运行开关（本进程内生效）。"""
    global WORD_CONTENT_PRESERVE, WORD_PRESERVE_PAGE_COUNT, WORD_IMAGE_RISK_GUARD, WORD_PRESERVE_LINKED_IMAGES
    global WORD_STEP_TIMEOUT_SEC, WORD_SKIP_FILE_ON_TIMEOUT
    global WORD_HEADER_FOOTER_LAYOUT_FIX, WORD_FONT_PROFILE
    if word_content_preserve is not None:
        WORD_CONTENT_PRESERVE = bool(word_content_preserve)
    if word_preserve_page_count is not None:
        WORD_PRESERVE_PAGE_COUNT = bool(word_preserve_page_count)
    if word_image_risk_guard is not None:
        WORD_IMAGE_RISK_GUARD = bool(word_image_risk_guard)
    if word_preserve_linked_images is not None:
        WORD_PRESERVE_LINKED_IMAGES = bool(word_preserve_linked_images)
    if word_step_timeout_sec is not None:
        try:
            WORD_STEP_TIMEOUT_SEC = float(word_step_timeout_sec)
        except Exception:
            pass
    if word_skip_file_on_timeout is not None:
        WORD_SKIP_FILE_ON_TIMEOUT = bool(word_skip_file_on_timeout)
    if word_header_footer_layout_fix is not None:
        WORD_HEADER_FOOTER_LAYOUT_FIX = bool(word_header_footer_layout_fix)
    if word_font_profile is not None:
        p = str(word_font_profile).strip().lower()
        if p in ("chinese", "english", "mixed"):
            WORD_FONT_PROFILE = p


def _apply_font_profile_to_range(r):
    """按 WORD_FONT_PROFILE 设置字体。"""
    try:
        pr = (WORD_FONT_PROFILE or "mixed").strip().lower()
        if pr not in ("chinese", "english", "mixed"):
            pr = "mixed"
        if pr == "english":
            r.Font.Name = FONT_NAME_LATIN
        elif pr == "chinese":
            r.Font.Name = FONT_NAME_SONG
        else:
            try:
                r.Font.NameAscii = FONT_NAME_LATIN
                r.Font.NameFarEast = FONT_NAME_SONG
            except Exception:
                r.Font.Name = FONT_NAME_SONG
    except Exception:
        try:
            r.Font.Name = FONT_NAME_SONG
        except Exception:
            pass


def _word_font_profile_label():
    pr = (WORD_FONT_PROFILE or "mixed").strip().lower()
    if pr == "chinese":
        return "全文宋体"
    if pr == "english":
        return "全文 Times New Roman"
    return "中英混排（西文 Times New Roman，中文宋体）"


def _get_doc_page_count(doc):
    """当前文档总页数（WPS/Word 布局重算后统计）。"""
    try:
        doc.Repaginate()
    except Exception:
        pass
    try:
        return int(doc.ComputeStatistics(wdStatisticPages))
    except Exception:
        return None


def _restore_file_and_raise_page_change(backup_path, target_path, before, after):
    if backup_path and os.path.isfile(backup_path):
        try:
            shutil.copyfile(backup_path, target_path)
        except Exception:
            pass
    raise RuntimeError(
        f"【页数变化】处理前总页数 {before}，处理后 {after}，已恢复原文件并中止保存"
    )


def _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path):
    """若启用页数保护且页数变化，恢复备份并抛错。"""
    if not WORD_PRESERVE_PAGE_COUNT or pages_before is None:
        return
    pages_after = _get_doc_page_count(doc)
    if pages_after is None:
        return
    if pages_after != pages_before:
        logger.warning(
            "page count changed: before=%s after=%s path=%s",
            pages_before,
            pages_after,
            path,
        )
        _restore_file_and_raise_page_change(backup_path, path, pages_before, pages_after)


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
        slots = nr * nc
        if slots < 4:
            return False
        need = max(4, slots // 4)
        count_risk = 0
        k = 0
        for ri in range(1, nr + 1):
            for ci in range(1, nc + 1):
                k += 1
                try:
                    cell = table.Cell(ri, ci)
                    if _cell_has_risk_color(cell):
                        count_risk += 1
                        if count_risk >= need:
                            return True
                except Exception:
                    pass
                if count_risk + (slots - k) < need:
                    return False
        return count_risk >= need
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
    """获取或创建 Word Application 实例。批量复用模式下返回同一线程内已缓存的 Application。"""
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    _check_file_abort()
    global _word_batch_cached_app
    if word_batch_session_active() and _word_batch_cached_app is not None:
        try:
            _ = _word_batch_cached_app.Documents.Count
            return _word_batch_cached_app
        except Exception:
            _word_batch_drop_cached_app()
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    before = tasklist_pids_for_images(_word_host_image_names())
    for attempt in range(3):
        try:
            _check_file_abort()
            word = win32com.client.dynamic.Dispatch(get_word_progid())
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone，禁止一切弹窗
            try:
                global _LAST_WORD_PID
                _LAST_WORD_PID = resolve_office_app_pid(
                    word, win32process, _word_host_image_names(), before
                )
            except Exception:
                _LAST_WORD_PID = _try_get_app_pid(word)
            if word_batch_session_active():
                _word_batch_cached_app = word
            return word
        except Exception as e:
            is_rpc = getattr(e, "args", (None,))[0] in RPC_RETRY_CODES
            if is_rpc and attempt < 2:
                time.sleep(2)
                continue
            raise


def _open_word_document(word, path, *, read_only=True):
    """打开文档并在打开后刷新 _LAST_WORD_PID（窗口句柄常在此后才可用）。"""
    doc = _com_call(
        word.Documents.Open,
        path,
        ReadOnly=read_only,
        AddToRecentFiles=False,
    )
    try:
        global _LAST_WORD_PID
        np = refresh_pid_after_doc_open(word, win32process)
        if np:
            _LAST_WORD_PID = np
    except Exception:
        pass
    return doc


def verify_word_saved_no_pending_revisions(path: str):
    """
    仅保存任务落盘后：只读再次打开，检查是否仍有未接受修订。
    用于避免「界面显示成功但 ZIP/成品仍为旧稿或未合并修订」的假成功。
    返回 (True, "") 或 (False, 原因说明)。
    """
    if win32com is None:
        return True, ""
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        return False, "成品路径不存在，无法校验"
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".doc", ".docx", ".docm"):
        return True, ""
    word = None
    doc = None
    out_ok, out_msg = True, ""
    try:
        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=True)
        try:
            n_rev = int(doc.Revisions.Count)
        except Exception:
            n_rev = -1
        if n_rev > 0:
            out_ok, out_msg = (
                False,
                f"仍检测到约 {n_rev} 处未接受修订，请用 Word/WPS 打开合并修订后保存",
            )
        else:
            try:
                if bool(doc.TrackRevisions):
                    out_ok, out_msg = (
                        False,
                        "文档仍开启「修订」跟踪，请关闭修订跟踪后保存",
                    )
            except Exception:
                pass
    except Exception as e:
        out_ok, out_msg = False, "二次打开校验失败：" + str(e)
    finally:
        try:
            _word_post_document_cleanup(word, doc, mark_doc_saved=False)
        except Exception:
            pass
    return out_ok, out_msg


def _com_call(func, *args, retries=3, delay=2, **kwargs):
    """
    带重试的 COM 调用包装器。
    当 WPS/Word 正忙（弹窗、处理中）抛出 RPC_E_CALL_REJECTED 等错误时自动重试。
    """
    for attempt in range(retries):
        try:
            _check_file_abort()
            return func(*args, **kwargs)
        except FileAbortRequested:
            raise
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
    _check_file_abort()
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


def _check_cover_signature_on_doc(doc):
    """
    对已打开的 Word 文档做封面签字检查（不 Open / Quit），供单次会话内调用。
    """
    try:
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
        doc = _open_word_document(word, path, read_only=True)
        return _check_cover_signature_on_doc(doc)
    except Exception as e:
        return False, f"签字检查异常: {e}"
    finally:
        _word_post_document_cleanup(word, doc, mark_doc_saved=False)


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


def _embed_linked_pictures(doc, change_notes=None, full_change_log=None):
    """
    强制断开链接图片并内嵌到文档中，避免后续处理或跨机打开时出现红叉丢图。
    处理范围：正文 InlineShapes、正文 Shapes、页眉页脚 Shapes。
    """
    if not WORD_PRESERVE_LINKED_IMAGES:
        return
    t0 = time.perf_counter()
    converted = 0
    failures = 0

    def _break_link_obj(obj):
        nonlocal converted, failures
        try:
            lf = getattr(obj, "LinkFormat", None)
            if lf is None:
                return
            src = ""
            try:
                src = str(getattr(lf, "SourceFullName", "") or "")
            except Exception:
                src = ""
            try:
                lf.BreakLink()
                converted += 1
                if full_change_log is not None:
                    pv = _sanitize_change_preview(src, _MAX_FILE_DETAIL_PREVIEW) if src else ""
                    if pv:
                        full_change_log.append(f"图片断链内嵌：{pv}")
                    else:
                        full_change_log.append("图片断链内嵌：已处理 1 处链接图片")
            except Exception:
                failures += 1
        except Exception:
            failures += 1

    try:
        for i in range(1, int(doc.InlineShapes.Count) + 1):
            if _step_timed_out(t0):
                _raise_or_mark_timeout("图片断链内嵌")
            try:
                _break_link_obj(doc.InlineShapes(i))
            except Exception:
                failures += 1
    except Exception:
        pass

    try:
        for i in range(1, int(doc.Shapes.Count) + 1):
            if _step_timed_out(t0):
                _raise_or_mark_timeout("图片断链内嵌")
            try:
                _break_link_obj(doc.Shapes(i))
            except Exception:
                failures += 1
    except Exception:
        pass

    try:
        for si in range(1, int(doc.Sections.Count) + 1):
            sec = doc.Sections(si)
            for hf_type in (1, 2, 3):
                for hf_getter in (sec.Headers, sec.Footers):
                    try:
                        hf = hf_getter(hf_type)
                        if not hf.Exists:
                            continue
                        for i in range(1, int(hf.Shapes.Count) + 1):
                            if _step_timed_out(t0):
                                _raise_or_mark_timeout("图片断链内嵌")
                            try:
                                _break_link_obj(hf.Shapes(i))
                            except Exception:
                                failures += 1
                    except Exception:
                        pass
    except Exception:
        pass

    if change_notes is not None and (converted > 0 or failures > 0):
        msg = f"已执行图片断链内嵌：成功 {converted} 处"
        if failures > 0:
            msg += f"，失败 {failures} 处"
        change_notes.append(msg)


def _remove_draft_watermark_shapes(doc, change_notes=None, full_change_log=None):
    """删除页眉页脚中包含草稿关键词的 Shape（水印）。"""
    deleted = 0
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
                                    raw = sh.TextFrame.TextRange.Text or ""
                                    t = raw.upper()
                                    for kw in DRAFT_WATERMARK_KEYWORDS:
                                        if kw.upper() in t:
                                            to_delete.append((i, raw))
                                            break
                            except Exception:
                                pass
                        for i, raw in sorted(to_delete, key=lambda x: x[0], reverse=True):
                            try:
                                hf.Shapes(i).Delete()
                                deleted += 1
                                if full_change_log is not None:
                                    pv = _sanitize_change_preview(
                                        raw, _MAX_FILE_DETAIL_PREVIEW
                                    )
                                    if pv:
                                        full_change_log.append(
                                            f"水印/草稿图形：已删除 Shape，文字预览「{pv}」"
                                        )
                                    else:
                                        full_change_log.append(
                                            "水印/草稿图形：已删除 Shape（无文字预览）"
                                        )
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass
    if change_notes is not None and deleted > 0:
        if full_change_log is not None:
            change_notes.append(
                f"已删除 {deleted} 个草稿/水印类图形对象（逐条见修改明细）"
            )
        else:
            msg = f"已删除 {deleted} 个草稿/水印类图形对象"
            change_notes.append(msg)


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


def _apply_formal_header_footer_fixes(doc):
    """
    正式性收尾（页眉页脚）：多节页脚空缺时按规则从首节补全；缺页码的主页脚补 PAGE 域。
    须在 _sync_page_numbers_after_edit 之前调用，以便随后统一刷新域。
    """
    if not WORD_HEADER_FOOTER_LAYOUT_FIX:
        return
    try:
        _ensure_headers_footers_consistent(doc)
    except Exception:
        pass
    try:
        _ensure_page_numbers(doc)
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


def _update_toc_tables_only(doc):
    """仅通过 TablesOfContents 更新目录（与全量 doc.Fields.Update 分离）。"""
    n_toc, toc_coll = _tables_of_contents_safe(doc)
    if not n_toc or not toc_coll:
        return
        try:
            for i in range(1, n_toc + 1):
                try:
                    toc_coll(i).Update()
                except Exception:
                    pass
        except Exception:
            pass


def _field_is_toc_or_pagination(field):
    """是否目录/页码/节页等与分页、目录相关的域（不含 HYPERLINK、INCLUDEPICTURE 等）。"""
    try:
        ft = int(getattr(field, "Type", -999))
    except Exception:
        return False
    return ft in (
        wdFieldPageRef,
        wdFieldNumPages,
        wdFieldSection,
        wdFieldSectionPages,
        wdFieldPage,
        wdFieldTOC,
    )


def _update_pagination_fields_selective(doc):
    """不调用 doc.Fields.Update()，只刷新目录/页码类域（正文 + 页眉页脚），降低链接图被全量刷新破坏的风险。"""
    try:
        for i in range(1, int(doc.Fields.Count) + 1):
            try:
                f = doc.Fields(i)
                if _field_is_toc_or_pagination(f):
                    f.Update()
            except Exception:
                pass
    except Exception:
        pass
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
                        if r is None:
                            continue
                        for fi in range(1, int(r.Fields.Count) + 1):
                            try:
                                f2 = r.Fields(fi)
                                if _field_is_toc_or_pagination(f2):
                                    f2.Update()
                            except Exception:
                                pass
                    except Exception:
                        pass
    except Exception:
        pass


def _ensure_toc_updated(doc):
    """更新所有目录域与域结果。"""
    _update_toc_tables_only(doc)
    if WORD_PRESERVE_LINKED_IMAGES:
        _update_pagination_fields_selective(doc)
        return
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


def _word_scale_table_inline_pictures_to_fit(doc):
    """
    非风险矩阵表格：单元格内嵌入图若大于单元格可视宽高，则等比例缩小（宽与高同比例）。
    文字显示不全仍依赖行高自动（_auto_fit_tables 等），不通过拉伸单维解决图片。
    """
    wdInlineShapePicture = 3
    wdInlineShapeLinkedPicture = 4
    t0 = time.perf_counter()
    try:
        for ti in range(1, doc.Tables.Count + 1):
            if _step_timed_out(t0):
                _raise_or_mark_timeout("表格图片缩放")
            tbl = doc.Tables(ti)
            try:
                if _is_risk_matrix_table(tbl):
                    continue
            except Exception:
                pass
            try:
                nc = int(tbl.Cells.Count)
            except Exception:
                continue
            for ci in range(1, nc + 1):
                if _step_timed_out(t0):
                    _raise_or_mark_timeout("表格图片缩放")
                try:
                    cell = tbl.Cells(ci)
                    rng = cell.Range
                    try:
                        max_w = float(cell.Width)
                        max_h = float(cell.Height)
                    except Exception:
                        try:
                            max_w = float(rng.Width)
                            max_h = float(rng.Height)
                        except Exception:
                            continue
                    if max_w < 8.0 or max_h < 8.0:
                        continue
                    nsh = int(rng.InlineShapes.Count)
                    for j in range(1, nsh + 1):
                        if _step_timed_out(t0):
                            _raise_or_mark_timeout("表格图片缩放")
                        try:
                            ish = rng.InlineShapes(j)
                            it = int(ish.Type)
                            if it not in (wdInlineShapePicture, wdInlineShapeLinkedPicture):
                                continue
                            iw = float(ish.Width)
                            ih = float(ish.Height)
                            if iw <= 0 or ih <= 0:
                                continue
                            fit_scale = min(max_w / iw, max_h / ih) * 0.98
                            # 小图可放大，但设上限避免过度失真
                            if fit_scale > 1.0:
                                scale = min(fit_scale, 2.0)
                            else:
                                scale = fit_scale
                            if abs(scale - 1.0) < 0.03:
                                continue
                            try:
                                ish.LockAspectRatio = True
                            except Exception:
                                pass
                            ish.Width = iw * scale
                            ish.Height = ih * scale
                            # 放大后若仍偏小：适度增加所在行高度，给图片更多展示空间
                            try:
                                nw = float(ish.Width)
                                nh = float(ish.Height)
                                if nw < max_w * 0.78 or nh < max_h * 0.78:
                                    need_h = max_h
                                    target_h = max(nh / 0.85, need_h)
                                    grow_h = min(1.30, max(1.0, target_h / max(need_h, 1.0)))
                                    try:
                                        row_idx = int(cell.RowIndex)
                                        row_obj = tbl.Rows(row_idx)
                                        cur_h = float(getattr(row_obj, "Height", 0) or 0)
                                        row_obj.HeightRule = wdRowHeightAuto
                                        if cur_h > 1.0:
                                            row_obj.Height = min(1584.0, cur_h * grow_h)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        except Exception:
                            pass
                except Exception:
                    pass
    except Exception:
        pass


def _remove_strikethrough_text(doc, change_notes=None, full_change_log=None):
    """
    删除全文（含页眉页脚等 StoryRanges）中带删除线格式的文本。
    目的：在正式性修订时，直接去除人工标记为删除线的内容，避免其继续留在成文里。
    full_change_log 为列表时，逐条追加每条删除线文本（供修改明细文件）；不写入 change_notes 以免与页面摘要重复。
    """
    wd_find_stop = 0
    wd_replace_none = 0
    wd_story_types = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    strike_count = 0
    for st_type in wd_story_types:
        try:
            r = doc.StoryRanges(st_type)
        except Exception:
            r = None
        while r is not None:
            try:
                f = r.Find
                try:
                    f.ClearFormatting()
                except Exception:
                    pass
                try:
                    f.Replacement.ClearFormatting()
                except Exception:
                    pass
                f.Text = ""
                f.Replacement.Text = ""
                f.Forward = True
                f.Wrap = wd_find_stop
                f.Format = True
                try:
                    f.Font.StrikeThrough = True
                except Exception:
                    pass
                while True:
                    try:
                        found = bool(f.Execute(Replace=wd_replace_none))
                    except TypeError:
                        found = bool(f.Execute())
                    if not found:
                        break
                    try:
                        # 清空 Range 会连带删除内嵌图，故含 InlineShape 的命中跳过并前移，避免仅保存模式丢图。
                        try:
                            if int(r.InlineShapes.Count) > 0:
                                r.Collapse(0)  # wdCollapseEnd
                                continue
                        except Exception:
                            pass
                        snippet = (r.Text or "").strip()
                        if snippet:
                            pv = _sanitize_change_preview(
                                snippet, _MAX_FILE_DETAIL_PREVIEW
                            )
                            if full_change_log is not None:
                                strike_count += 1
                                if pv:
                                    full_change_log.append(
                                        f"删除线：已删除文本「{pv}」"
                                    )
                                else:
                                    full_change_log.append(
                                        "删除线：已删除文本（无可见预览）"
                                    )
                            elif change_notes is not None:
                                strike_count += 1
                        r.Text = ""
                    except Exception:
                        break
            except Exception:
                pass
            try:
                r = r.NextStoryRange
            except Exception:
                break
    if change_notes is not None and strike_count > 0:
        if full_change_log is not None:
            change_notes.append(
                f"已删除 {strike_count} 处带删除线文本（逐条见修改明细）"
            )
        else:
            change_notes.append("已删除文中带删除线格式的文本")


def _cleanup_extra_page_breaks(doc, change_notes=None):
    """
    清理“空白页/多余分页符”：
    - 合并连续手动分页符（^m^m -> ^m）
    - 删除文末多余的分页符（最后一页为空白的常见原因）

    仅处理主文档内容（doc.Content），不动页眉页脚，尽量降低对模板的副作用。
    """
    touched = False
    wd_find_stop = 0
    wd_replace_one = 1
    wd_replace_all = 2
    try:
        r = doc.Content
    except Exception:
        return

    # 1) 连续分页符压缩：^m^m -> ^m（多次循环直到不存在）
    try:
        f = r.Find
        try:
            f.ClearFormatting()
        except Exception:
            pass
        try:
            f.Replacement.ClearFormatting()
        except Exception:
            pass
        f.Forward = True
        f.Wrap = wd_find_stop
        f.Format = False
        f.Text = "^m^m"
        f.Replacement.Text = "^m"
        for _ in range(50):  # 防止极端情况死循环
            try:
                changed = bool(f.Execute(Replace=wd_replace_all))
            except TypeError:
                changed = bool(f.Execute())
            if not changed:
                break
            touched = True
    except Exception:
        pass

    # 2) 删除文末多余分页符（\f），以及分页符前的空段落
    try:
        # Word 中手动分页符在 Range.Text 里通常是 \x0c（FormFeed）
        for _ in range(50):
            t = str(r.Text or "")
            if not t:
                break
            # 先去尾部空白（CR/LF/空格/制表符）
            t2 = t.rstrip(" \t\r\n")
            if t2 != t:
                end = doc.Range(r.Start, r.Start + len(t2))
                r = end
                touched = True
                continue
            if t2.endswith("\x0c"):
                # 删除最后一个分页符
                last = doc.Range(r.End - 1, r.End)
                last.Text = ""
                touched = True
                try:
                    r = doc.Content
                except Exception:
                    break
                continue
            break
    except Exception:
        pass
    if change_notes is not None and touched:
        change_notes.append("已调整主文档分页符（合并连续分页或删除文末多余分页符）")


def auto_fix_formal_word(
    doc_path,
    save_path=None,
    change_notes=None,
    full_change_log=None,
    *,
    run_cover_check=False,
):
    """
    正式性检查失败后自动修复。执行顺序：表格排版 → 统一表格框线 → 去高亮/统一字体/去水印 → 目录更新 → 接受修订 → 删除批注 → 页眉页脚与页码补齐 → 保存。
    非风险矩阵中的标红、标黄、标绿会被修复为黑色；风险矩阵内保留。
    change_notes：传入 list 时追加格式/内容变更说明（供仅保存模式展示）。
    full_change_log：传入 list 时追加删除线、水印等逐条记录（写入下载包「修改明细」）。
    run_cover_check：为 True 时在保存前对当前文档做封面签字检查（仅追加提示，不中断保存）。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    backup_path = None
    try:
        try:
            fd, backup_path = _mkstemp_word_backup()
            os.close(fd)
            shutil.copyfile(path, backup_path)
        except Exception:
            backup_path = None

        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=False)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        _embed_linked_pictures(doc, change_notes, full_change_log)
        saved_risk_matrices = _save_risk_matrix_formats(doc)
        _auto_fit_tables(doc)
        _word_scale_table_inline_pictures_to_fit(doc)
        _normalize_table_borders(doc)
        wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        for stType in wdStoryTypes:
            try:
                r = doc.StoryRanges(stType)
                while r is not None:
                    try:
                        r.HighlightColorIndex = wdNoHighlight
                        r.Font.Color = wdColorBlack
                        _apply_font_profile_to_range(r)
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
            _apply_font_profile_to_range(doc.Content)
            doc.Content.Font.Hidden = False
        except Exception:
            pass
        # 保真模式下不主动删 Shape（水印/图形可能与业务图片共用对象类型）。
        if not WORD_CONTENT_PRESERVE:
            _remove_draft_watermark_shapes(doc, change_notes, full_change_log)
        _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True)
        _restore_risk_matrix_formats(doc, saved_risk_matrices)
        if change_notes is not None:
            change_notes.append(
                "【格式】已去除高亮与标黄、将非黑色字体改为黑色并统一为「"
                + _word_font_profile_label()
                + "」（风险矩阵内保留原色；删除线/水印逐条见修改明细）"
            )
        _ensure_toc_updated(doc)
        n_rev = 0
        try:
            n_rev = int(doc.Revisions.Count)
        except Exception:
            pass
        _accept_all_revisions_in_document(doc)
        if change_notes is not None and n_rev > 0:
            change_notes.append(f"已接受全部修订（约 {n_rev} 处跟踪更改）")
        # 修订合并前后 InlineShape 计数可能不一致（含“修订插入”图），在此之后拍 before 才与后续步骤可比。
        before_visual = _snapshot_visual_objects(doc)
        # 保真模式下不删除删除线文本、不清理分页符，避免误删内容/锚点导致图片异常。
        if not WORD_CONTENT_PRESERVE:
            _remove_strikethrough_text(doc, change_notes, full_change_log)
            _cleanup_extra_page_breaks(doc, change_notes)
        n_com = 0
        try:
            n_com = int(doc.Comments.Count)
        except Exception:
            pass
        _delete_all_comments_robust(doc)
        if change_notes is not None and n_com > 0:
            change_notes.append(f"已删除 {n_com} 条批注")
        after_visual = _snapshot_visual_objects(doc)
        if WORD_IMAGE_RISK_GUARD and _visual_objects_lost(before_visual, after_visual):
            # 先走兜底流程；若仍疑似异常，不中断主流程，避免误报导致任务全失败。
            _safe_accept_and_normalize_word(doc)
            after_visual2 = _snapshot_visual_objects(doc)
            if _visual_objects_lost(before_visual, after_visual2):
                raise RuntimeError("【图片完整性风险】检测到疑似图片/图形对象减少，请手动打印原文")
        _apply_formal_header_footer_fixes(doc)
        _sync_page_numbers_after_edit(doc)
        _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path)
        if run_cover_check:
            ok_cov, msg_cov = _check_cover_signature_on_doc(doc)
            if not ok_cov and change_notes is not None:
                change_notes.append("封面签字未通过（仍继续保存）：" + msg_cov)
        _finalize_no_track_changes_state(doc)
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
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
        if backup_path:
            try:
                if os.path.isfile(backup_path):
                    os.remove(backup_path)
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
        doc = _open_word_document(word, path, read_only=True)

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
        _word_post_document_cleanup(word, doc, mark_doc_saved=False)


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
        doc = _open_word_document(word, path, read_only=True)
        revs = doc.Revisions
        count = int(revs.Count)
        return count > 0, count
    except Exception:
        return False, 0
    finally:
        _word_post_document_cleanup(word, doc, mark_doc_saved=False)


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
    backup_path = None
    try:
        if WORD_PRESERVE_PAGE_COUNT:
            try:
                fd, backup_path = _mkstemp_word_backup()
                os.close(fd)
                shutil.copyfile(path, backup_path)
            except Exception:
                backup_path = None
        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=False)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.Content.Font.Color = wdColorBlack
            _apply_font_profile_to_range(doc.Content)
            # 与各 Story（含页眉页脚、脚注等）一致，避免仅改主文遗漏页眉页脚非黑字
            wdStoryTypes = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
            for st in wdStoryTypes:
                try:
                    r = doc.StoryRanges(st)
                    while r is not None:
                        try:
                            r.Font.Color = wdColorBlack
                            _apply_font_profile_to_range(r)
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
                _apply_font_profile_to_range(doc.Content)
                if remove_highlights:
                    doc.Content.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
        _apply_formal_header_footer_fixes(doc)
        _sync_page_numbers_after_edit(doc)
        _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path)
        _save_doc(doc, path, save_path)
        return True
    finally:
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
        if backup_path:
            try:
                if os.path.isfile(backup_path):
                    os.remove(backup_path)
            except Exception:
                pass


def _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True):
    """
    逐段按当前字体策略统一字体（同一段落内一致），可选同时设为黑色、去高亮。
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
                _apply_font_profile_to_range(r)
                if set_black:
                    r.Font.Color = wdColorBlack
                if remove_highlight:
                    r.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
    except Exception:
        pass


def _finalize_no_track_changes_state(doc):
    """
    保存前强制关闭修订跟踪与「显示标记」视图。
    否则部分宿主（尤其 WPS）落盘后仍呈修订/草稿视图，或 settings 仍带跟踪标记。
    """
    try:
        doc.TrackRevisions = False
    except Exception:
        pass
    try:
        doc.ShowRevisions = False
    except Exception:
        pass
    for win_try in (
        lambda: doc.ActiveWindow,
        lambda: doc.Windows(1) if int(getattr(doc.Windows, "Count", 0) or 0) >= 1 else None,
    ):
        try:
            aw = win_try()
            if aw is None:
                continue
            v = aw.View
            try:
                v.ShowRevisionsAndComments = False
            except Exception:
                pass
            try:
                v.RevisionsView = 0
            except Exception:
                pass
        except Exception:
            pass


def _accept_revisions_residual_loop(doc, max_iter=100000):
    """逐条 Accept 残留修订，避免 Count 很大时 range(Count+1) 不够。"""
    try:
        for _ in range(max_iter):
            try:
                n = int(doc.Revisions.Count)
            except Exception:
                break
            if n <= 0:
                break
            try:
                doc.Revisions(1).Accept()
            except Exception:
                break
    except Exception:
        pass


def _delete_all_comments_robust(doc):
    """
    尽可能删除所有批注/评论。
    某些 WPS/Word 场景下 doc.DeleteAllComments() 可能无效或部分失败，因此增加按集合逐条删除兜底。
    """
    try:
        n = int(getattr(getattr(doc, "Comments", None), "Count", 0) or 0)
    except Exception:
        n = 0
    if n <= 0:
        return 0
    try:
        doc.DeleteAllComments()
    except Exception:
        pass
    deleted = 0
    try:
        for _ in range(n + 5):
            try:
                c = doc.Comments(1)
            except Exception:
                break
            try:
                c.Delete()
                deleted += 1
            except Exception:
                break
    except Exception:
        pass
    try:
        left = int(doc.Comments.Count)
    except Exception:
        left = 0
    return max(0, n - left)


def _accept_all_revisions_in_document(doc, *, skip_builtin_toc_pass=False):
    """
    接受文档中所有修订，包括正文、目录页、页眉页脚、文本框等。

    关键：必须先关闭 TrackRevisions，否则后续 Fields.Update / TOC.Update
    会把目录重新生成的内容当作新修订记录下来，导致修订永远清不完。

    skip_builtin_toc_pass=True：不执行末尾的目录/全量域更新与紧随其后的二次修订清理
    （供多表格轻量路径等调用方在字体等处理之后再统一 _ensure_toc_updated，避免重复 Fields.Update）。
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

    t0 = time.perf_counter()
    try:
        n_rev_after_bulk = int(doc.Revisions.Count)
    except Exception:
        n_rev_after_bulk = -1
    # ===== 2) 逐段接受（含目录页每一段）— 无残留修订时跳过，避免表格文档上万段 COM =====
    if n_rev_after_bulk > 0:
        try:
            for i in range(1, doc.Paragraphs.Count + 1):
                if _step_timed_out(t0):
                    _raise_or_mark_timeout("接受修订（深扫）")
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
                if _step_timed_out(t0):
                    _raise_or_mark_timeout("接受修订（深扫）")
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
    _accept_revisions_residual_loop(doc)

    if skip_builtin_toc_pass:
        return

    # ===== 10) 更新目录（TrackRevisions 已关闭，不会产生新修订） =====
    # 图片保全模式下避免全量 Fields.Update，仍刷新目录与页码相关域。
    _update_toc_tables_only(doc)
    if WORD_PRESERVE_LINKED_IMAGES:
        _update_pagination_fields_selective(doc)
    else:
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
    _accept_revisions_residual_loop(doc)
    _finalize_no_track_changes_state(doc)


def accept_all_revisions_and_save(
    doc_path,
    save_path=None,
    ensure_font_black_too=True,
    change_notes=None,
    full_change_log=None,
):
    """
    接受文档中所有修订，去除标黄高亮，可选按字体策略统一为黑色，并保存。
    save_path 为 None 时覆盖原文件。
    change_notes：传入 list 时追加变更说明（供仅保存模式展示）。
    full_change_log：传入 list 时追加删除线等逐条记录（写入下载包「修改明细」）。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    backup_path = None
    try:
        try:
            fd, backup_path = _mkstemp_word_backup()
            os.close(fd)
            shutil.copyfile(path, backup_path)
        except Exception:
            backup_path = None

        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=False)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        _embed_linked_pictures(doc, change_notes, full_change_log)
        n_rev = 0
        try:
            n_rev = int(doc.Revisions.Count)
        except Exception:
            pass
        _accept_all_revisions_in_document(doc)
        if change_notes is not None and n_rev > 0:
            change_notes.append(f"已接受全部修订（约 {n_rev} 处跟踪更改）")
        # 同上：先合并修订再拍快照，避免修订插入图在合并前后计数差异误报。
        before_visual = _snapshot_visual_objects(doc)
        if not WORD_CONTENT_PRESERVE:
            _remove_strikethrough_text(doc, change_notes, full_change_log)
            _cleanup_extra_page_breaks(doc, change_notes)
        _auto_fit_tables(doc)
        _word_scale_table_inline_pictures_to_fit(doc)
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
                            _apply_font_profile_to_range(r)
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
                _apply_font_profile_to_range(doc.Content)
                doc.Content.HighlightColorIndex = wdNoHighlight
            except Exception:
                pass
        if ensure_font_black_too:
            _unify_paragraph_fonts(doc, set_black=True, remove_highlight=True)
        _restore_risk_matrix_formats(doc, saved_risk_matrices)
        if change_notes is not None:
            change_notes.append(
                "【格式】已去除高亮与标黄、将非黑色字体改为黑色并统一为「"
                + _word_font_profile_label()
                + "」（风险矩阵内保留原色；删除线逐条见修改明细）"
            )
        after_visual = _snapshot_visual_objects(doc)
        if WORD_IMAGE_RISK_GUARD and _visual_objects_lost(before_visual, after_visual):
            _safe_accept_and_normalize_word(doc)
            after_visual2 = _snapshot_visual_objects(doc)
            if _visual_objects_lost(before_visual, after_visual2):
                raise RuntimeError("【图片完整性风险】检测到疑似图片/图形对象减少，请手动打印原文")
        _apply_formal_header_footer_fixes(doc)
        _sync_page_numbers_after_edit(doc)
        _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path)
        _finalize_no_track_changes_state(doc)
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
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
        if backup_path:
            try:
                if os.path.isfile(backup_path):
                    os.remove(backup_path)
            except Exception:
                pass


def accept_revisions_basic_word(doc_path, save_path=None):
    """
    轻量处理：仅接受修订、去除标黄/高亮、统一表格框线、删批注后保存。
    不做整篇改宋体/逐段改色等重处理，用于疑似图片风险文档仍自动打印的场景。
    """
    if win32com is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    path = os.path.abspath(doc_path)
    save_path = os.path.abspath(save_path) if save_path else path
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    word = None
    doc = None
    backup_path = None
    try:
        if WORD_PRESERVE_PAGE_COUNT:
            try:
                fd, backup_path = _mkstemp_word_backup()
                os.close(fd)
                shutil.copyfile(path, backup_path)
            except Exception:
                backup_path = None
        word = _get_word_app(visible=False)
        doc = _open_word_document(word, path, read_only=False)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        _accept_all_revisions_in_document(doc)
        _normalize_table_borders(doc)
        wd_story_basic = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
        for st_type in wd_story_basic:
            try:
                r = doc.StoryRanges(st_type)
                while r is not None:
                    try:
                        r.HighlightColorIndex = wdNoHighlight
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
        except Exception:
            pass
        _delete_all_comments_robust(doc)
        _apply_formal_header_footer_fixes(doc)
        _sync_page_numbers_after_edit(doc)
        _check_page_count_unchanged_or_restore(doc, pages_before, backup_path, path)
        _finalize_no_track_changes_state(doc)
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
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
        if backup_path:
            try:
                if os.path.isfile(backup_path):
                    os.remove(backup_path)
            except Exception:
                pass


def _sync_page_numbers_after_edit(doc):
    """
    批量接受修订/规范化后，总页数或分页常会变，需重算版式并刷新页码域（PAGE/NUMPAGES 等）。
    正文与每节页眉页脚中的域分别更新，兼容页码在页眉页脚的情况。
    """
    preserve = WORD_PRESERVE_LINKED_IMAGES
    try:
        doc.Repaginate()
    except Exception:
        pass
    if preserve:
        _update_pagination_fields_selective(doc)
    else:
        try:
            for i in range(1, doc.Fields.Count + 1):
                try:
                    doc.Fields(i).Update()
                except Exception:
                    pass
        except Exception:
            pass
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
                            if r is None:
                                continue
                            n = int(r.Fields.Count)
                            for fi in range(1, n + 1):
                                try:
                                    r.Fields(fi).Update()
                                except Exception:
                                    pass
                        except Exception:
                            pass
        except Exception:
            pass
    _update_toc_tables_only(doc)
    if preserve:
        _update_pagination_fields_selective(doc)
    else:
        try:
            doc.Fields.Update()
        except Exception:
            pass
    try:
        doc.Content.ComputeStatistics(wdStatisticPages)
    except Exception:
        pass
    # 分页稳定后再刷一次域，避免总页数仍滞后
    try:
        doc.Repaginate()
    except Exception:
        pass
    if preserve:
        _update_pagination_fields_selective(doc)
        _update_toc_tables_only(doc)
    else:
        try:
            doc.Fields.Update()
        except Exception:
            pass
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
                            if r is None:
                                continue
                            for fi in range(1, int(r.Fields.Count) + 1):
                                try:
                                    r.Fields(fi).Update()
                                except Exception:
                                    pass
                        except Exception:
                            pass
        except Exception:
            pass


def _refresh_word_pagination_before_print(doc):
    """
    所有修改完成后、PrintOut 前刷新分页与页码相关域，减少换页/页码与版面不一致。
    WPS 可能不支持部分 API，逐项 try。
    """
    _sync_page_numbers_after_edit(doc)


def _ensure_no_markup_print(word, doc):
    """
    确保打印/导出 PDF 不携带修订标记/批注气泡等。
    说明：是否“打印标记”常是宿主全局选项（WPS/Word UI 设置），仅隐藏视图不一定影响 PrintOut/ExportAsFixedFormat。
    """
    # 1) 先尽量切到“无标记”视图（不落盘）
    try:
        _finalize_no_track_changes_state(doc)
    except Exception:
        pass
    # 2) 再关掉宿主打印选项（不同版本/WPS 可能不支持部分字段，逐项 try）
    try:
        opt = getattr(word, "Options", None)
        if opt is not None:
            for k, v in (
                ("PrintRevisions", False),
                ("PrintComments", False),
                ("PrintHiddenText", False),
                ("PrintFieldCodes", False),
            ):
                try:
                    setattr(opt, k, v)
                except Exception:
                    pass
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
        doc = _open_word_document(word, path, read_only=True)
        _ensure_no_markup_print(word, doc)
        _refresh_word_pagination_before_print(doc)
        if _is_pdf_printer(printer_name):
            out_pdf = _desktop_pdf_path(path)
            doc.ExportAsFixedFormat(out_pdf, 17)  # wdExportFormatPDF
            return out_pdf
        doc.PrintOut(
            Background=True,
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
        _word_post_document_cleanup(word, doc, mark_doc_saved=False)


def print_word_with_basic_processing_no_save(
    doc_path,
    printer_name=None,
    copies=1,
    accept_revisions=True,
    remove_highlights=True,
):
    """
    保真打印（不落盘）：
    - 在内存中执行基础处理（接受修订、可选去标黄）；
    - 同步页码后直接打印/导出 PDF；
    - 关闭时不保存，避免 WPS 保存阶段导致图片丢失。
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
        doc = _open_word_document(word, path, read_only=False)
        _ensure_no_markup_print(word, doc)
        pages_before = _get_doc_page_count(doc) if WORD_PRESERVE_PAGE_COUNT else None
        try:
            doc.TrackRevisions = False
        except Exception:
            pass
        if accept_revisions:
            _accept_all_revisions_in_document(doc)
        if remove_highlights:
            wd_story_basic = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
            for st_type in wd_story_basic:
                try:
                    r = doc.StoryRanges(st_type)
                    while r is not None:
                        try:
                            r.HighlightColorIndex = wdNoHighlight
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
            except Exception:
                pass
        _apply_formal_header_footer_fixes(doc)
        _sync_page_numbers_after_edit(doc)
        if WORD_PRESERVE_PAGE_COUNT and pages_before is not None:
            pages_after = _get_doc_page_count(doc)
            if pages_after is not None and pages_after != pages_before:
                logger.warning(
                    "no-save print would change page count; before=%s after=%s path=%s",
                    pages_before,
                    pages_after,
                    path,
                )
                raise RuntimeError(
                    "【页数变化】基础处理将改变总页数，请直接打印原文件（未做会改变页数的处理）"
                )
        if _is_pdf_printer(printer_name):
            out_pdf = _desktop_pdf_path(path)
            doc.ExportAsFixedFormat(out_pdf, 17)  # wdExportFormatPDF
            return out_pdf
        doc.PrintOut(
            Background=True,
            Append=False,
            Range=0,
            Item=0,
            Copies=int(copies),
            Collate=True,
        )
        time.sleep(3)
        return True
    finally:
        if word and old_printer:
            try:
                word.ActivePrinter = old_printer
            except Exception:
                pass
        _word_post_document_cleanup(word, doc, mark_doc_saved=True)
