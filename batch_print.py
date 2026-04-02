# -*- coding: utf-8 -*-
"""
批量打印工具：在打印前检查封面签字（Word）、修订记录，并自动接受所有修订后打印。
支持：.doc/.docx、.xls/.xlsx、.pdf
"""
from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
import threading

# 支持的扩展名（小写）
WORD_EXT = {".doc", ".docx", ".docm"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm"}
PDF_EXT = {".pdf"}

# COM「发生意外」错误码，WPS/Office/Excel 易触发
_COM_ERR_CODES = (-2147352567, -2147467259, -2146827864)
_COM_ERR_MSG = "文档处理时发生意外，请确认 WPS/Excel 已安装、文档未被占用，或稍后重试。"

logger = logging.getLogger("aiprintword.batch")


def _batch_log_and_web(msg, *args):
    """批处理日志同时打到 aiprintword.batch 与 aiprintword.web（避免只配置了 web 时看不到 precheck）。"""
    logger.info(msg, *args)
    try:
        logging.getLogger("aiprintword.web").info(msg, *args)
    except Exception:
        pass


def _emit_word_lite_precheck(path: str, raw_print: bool) -> None:
    """
    在处理 Word 之前输出「阈值 / 表格行合计 / 计划是否走 lite」。
    注意：planned_lite 仅由 XML 计数得到；.doc 需转 docx 后 _process_word 内才会最终判定。
    """
    try:
        from runtime_settings.resolve import get_setting

        th = int(get_setting("WORD_MANY_TABLE_ROWS_LITE_THRESHOLD"))
    except Exception:
        th = 100
    try:
        ext = os.path.splitext(path)[1].lower()
    except Exception:
        ext = ""
    base = os.path.basename(path)
    if raw_print:
        _batch_log_and_web(
            "word precheck: raw_print=True threshold=%s ext=%s planned_lite=N/A file=%s",
            th,
            ext or "?",
            base,
        )
        return
    if ext in (".docx", ".docm"):
        try:
            from doc_handlers.word_handler import docx_table_row_count_hint

            n = docx_table_row_count_hint(path)
        except Exception:
            n = -1
        planned = (n >= th) if n >= 0 else False
        _batch_log_and_web(
            "word precheck: ext=%s table_rows=%s threshold=%s planned_lite=%s file=%s",
            ext,
            n,
            th,
            planned,
            base,
        )
        return
    if ext == ".doc":
        _batch_log_and_web(
            "word precheck: ext=.doc threshold=%s planned_lite=unknown (先转 docx 后最终判定) file=%s",
            th,
            base,
        )
        return
    _batch_log_and_web(
        "word precheck: ext=%s threshold=%s planned_lite=unknown file=%s",
        ext or "?",
        th,
        base,
    )


# Word/Excel COM 非线程安全；多客户端同时批处理会并发 Dispatch，易触发 RPC 拒绝（被呼叫方拒绝接收呼叫）。
_OFFICE_COM_LOCK = threading.RLock()


def _word_modification_report_text(full_change_log, change_notes):
    """合并逐条修改明细与其它说明，供下载包「修改明细」使用。"""
    parts = []
    if full_change_log:
        parts.append("【删除线 / 水印等逐条】")
        parts.extend(full_change_log)
    if change_notes:
        if parts:
            parts.append("")
        parts.append("【其它处理说明】")
        parts.extend(change_notes)
    text = "\n".join(parts) if parts else ""
    return text if text.strip() else None


def _excel_modification_report_text(hints, *, check_formal, auto_fixed, dry_run):
    """Excel 仅保存：生成写入 ZIP 的步骤与提示说明（无 Word 式 XML 逐条日志）。"""
    parts = []
    if hints:
        parts.append("【处理提示】")
        parts.extend(str(h) for h in hints if str(h).strip())
    steps = []
    if check_formal:
        steps.append("正式性检查" + ("；已执行自动修复" if auto_fixed else ""))
    if not dry_run:
        steps.extend(
            [
                "接受工作簿更改/修订（若已启用跟踪修订）",
                "工作表字体改黑",
                "风险矩阵区域外：清除单元格底色、自适应行高",
            ]
        )
    if steps:
        if parts:
            parts.append("")
        parts.append("【已执行步骤】")
        for s in steps:
            parts.append("- " + s)
    text = "\n".join(parts) if parts else ""
    return text if text.strip() else None


# 写入「修改明细.txt」时去掉自我指涉套话（避免在明细文件里写「详见修改明细」）
_ZIP_POINTS_NOISE_SUBSTR = (
    "见下载包",
    "见修改明细",
    "详见",
    "完整逐条明细",
    "同包「修改明细",
    "根目录「修改明细",
)


def _zip_sanitize_points(points):
    out = []
    for p in points or []:
        s = str(p).strip()
        if not s or any(x in s for x in _ZIP_POINTS_NOISE_SUBSTR):
            continue
        out.append(s)
    return out


def build_batch_modification_zip_text(details):
    """
    将 run_batch 返回的 details（须已含 filename）拼成 ZIP 内「修改明细.txt」正文。
    每条输出【处理模式】【修改要点/步骤】；有 modification_detail 时再附【逐条修改明细】。
    """
    blocks = []
    for d in details:
        fn = (d.get("filename") or os.path.basename(d.get("path") or "")).strip()
        if not fn:
            continue
        blocks.append("=" * 64)
        blocks.append(fn)
        blocks.append("=" * 64)
        mode_label = (d.get("processing_mode_label") or "—").strip() or "—"
        blocks.append(f"【处理模式】{mode_label}")
        pts = _zip_sanitize_points(d.get("modification_points") or [])
        if pts:
            blocks.append("【修改要点与步骤】")
            for p in pts:
                blocks.append(f"- {p}")
        md = d.get("modification_detail")
        if md is not None and str(md).strip():
            blocks.append("【逐条修改明细】")
            blocks.append(str(md).strip())
        elif d.get("success") and pts:
            blocks.append(
                "【说明】上列为本次已记录的处理摘要；若与 Word/Excel 中实际显示不一致，请以成品文件为准。"
            )
        blocks.append("")

    # 需人工清单（集中汇总）：超时跳过 / 手动打印提示 / 其它需人工项
    manual = []
    for d in details:
        fn = (d.get("filename") or os.path.basename(d.get("path") or "")).strip()
        msg = str(d.get("message") or "").strip()
        if not fn or not msg:
            continue
        if (
            d.get("manual_attention")
            or d.get("timeout_skip")
            or d.get("manual_skip")
            or ("需人工" in msg)
            or ("手动打印" in msg)
            or ("请手动打印原文" in msg)
        ):
            manual.append((fn, msg))
    if manual:
        if blocks and blocks[-1] != "":
            blocks.append("")
        blocks.append("【需人工清单】")
        blocks.append("以下文件未能完全自动处理，请人工打开核对/处理：")
        if any(
            d.get("timeout_skip") or "【超时跳过】" in (d.get("message") or "")
            for d in details
        ):
            blocks.append(
                "（单文件总超时条目在 ZIP 内附带上传时的原文副本，包内文件名为「原名_【超时原文】.扩展名」，与同批处理成功的文件区分。）"
            )
        for fn, msg in manual:
            blocks.append(f"- {fn}：{msg}")
        blocks.append("")

    text = "\n".join(blocks).strip()
    return text if text else None


def _batch_meta_set(holder, mode, label, points=None):
    """写入单文件处理结果元数据（processing_mode / label / modification_points）。"""
    if holder is None:
        return
    holder.clear()
    holder["processing_mode"] = mode
    holder["processing_mode_label"] = label
    holder["modification_points"] = [str(x) for x in (points or []) if str(x).strip()]


def _batch_merge_entry_processing(
    entry,
    meta_holder,
    *,
    doc_type,
    raw_print,
    dry_run,
    skip_print,
    lite_diag_holder=None,
    watchdog=None,
):
    """把 meta_holder 与兜底推断合并进 entry；watchdog 覆盖「需人工」类标签。"""
    mh = meta_holder or {}
    mode = mh.get("processing_mode")
    label = mh.get("processing_mode_label")
    pts = list(mh.get("modification_points") or [])
    ld = lite_diag_holder or {}
    if not mode and doc_type == "word":
        if raw_print:
            mode, label = "word_raw_print", "Word·直接打印（无规范化）"
        elif dry_run:
            if ld.get("word_lite_mode"):
                mode, label = "word_lite_precheck", "Word·多表格轻量（预检）"
            else:
                mode, label = "word_full_precheck", "Word·全量规范化（预检）"
        elif ld.get("word_lite_mode") and entry.get("success"):
            mode, label = "word_lite", "Word·多表格轻量"
        elif entry.get("success"):
            mode, label = "word_full", "Word·全量规范化"
    if not mode and doc_type == "excel":
        if raw_print:
            mode, label = "excel_raw_print", "Excel·直接打印（无规范化）"
        elif dry_run:
            mode, label = "excel_precheck", "Excel·预检"
        elif skip_print and entry.get("success"):
            mode, label = "excel_full_save", "Excel·全量规范化（已保存）"
        elif entry.get("success"):
            mode, label = "excel_full_print", "Excel·规范化打印"
    if not mode and doc_type == "pdf":
        if skip_print:
            mode, label = "pdf_pack_only", "PDF·原样打包"
        elif dry_run:
            mode, label = "pdf_precheck", "PDF·预检"
        else:
            mode, label = "pdf_print", "PDF·打印"
    if not mode:
        mode = "unknown"
    if not label:
        label = "—"
    entry["processing_mode"] = mode
    entry["processing_mode_label"] = label
    entry["modification_points"] = pts
    wd = watchdog or {}
    if wd.get("manual_skip"):
        entry["processing_mode"] = "manual_skipped"
        entry["processing_mode_label"] = "需人工·跳过当前文件"
        entry["manual_attention"] = True
    elif wd.get("timed_out"):
        entry["processing_mode"] = "manual_timeout"
        entry["processing_mode_label"] = "需人工·单文件超时"
        entry["manual_attention"] = True
    elif entry.get("skipped_output_exists"):
        entry["processing_mode"] = "skipped_output_exists"
        entry["processing_mode_label"] = "已跳过·即时目录已有同名文件"
    elif not entry.get("success"):
        msg = entry.get("message") or ""
        if "需人工" in msg or entry.get("timeout_skip") or entry.get("manual_skip"):
            entry["manual_attention"] = True
            pm = entry.get("processing_mode") or ""
            if pm not in ("manual_timeout", "manual_skipped") and not str(
                pm
            ).startswith("manual_required"):
                entry["processing_mode"] = "manual_required"
                entry["processing_mode_label"] = "需人工·未完成自动处理"


def _format_error(e):
    """若为 COM 异常则返回友好提示；已含【步骤名】的说明原样返回。"""
    try:
        s = str(e)
        if "【" in s and "】" in s:
            return s
    except Exception:
        pass
    try:
        if getattr(e, "args", None) and len(e.args) > 0 and e.args[0] in _COM_ERR_CODES:
            return _COM_ERR_MSG
    except Exception:
        pass
    return str(e)


def _collect_files(paths, recursive=False):
    """从路径列表收集所有支持的文档文件。"""
    collected = []
    for p in paths:
        p = os.path.abspath(p)
        if not os.path.exists(p):
            continue
        if os.path.isfile(p):
            ext = os.path.splitext(p)[1].lower()
            if ext in WORD_EXT | EXCEL_EXT | PDF_EXT:
                collected.append(p)
            continue
        for root, _, files in os.walk(p):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in WORD_EXT | EXCEL_EXT | PDF_EXT:
                    collected.append(os.path.join(root, f))
            if not recursive:
                break
    return collected


def _get_type(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in WORD_EXT:
        return "word"
    if ext in EXCEL_EXT:
        return "excel"
    if ext in PDF_EXT:
        return "pdf"
    return None


def _next_processed_doc_type(files, from_idx, skip_indices):
    """下一待处理文件类型（跳过 output_exists 占位），无则 None。"""
    j = from_idx + 1
    n = len(files)
    while j < n:
        if j in skip_indices:
            j += 1
            continue
        return _get_type(files[j])
    return None


# 各类型单文件内「步骤」上界（用于进度条；实际 report 次数可少不可多，超出则按满步计）
_PROGRESS_MAX_STEPS_WORD = 10
_PROGRESS_MAX_STEPS_EXCEL = 12
_PROGRESS_MAX_STEPS_PDF = 4


def _format_duration_cn(seconds):
    """将秒数格式化为中文可读时长（用于日志与结果说明）。"""
    try:
        s = float(max(0.0, seconds))
    except Exception:
        return "—"
    if s < 60:
        return f"{int(round(s))} 秒"
    m, sec = divmod(int(round(s)), 60)
    if m < 60:
        return f"{m} 分 {sec} 秒"
    h, m = divmod(m, 60)
    return f"{h} 小时 {m} 分 {sec} 秒"


def _batch_eta_payload(completed_file_seconds, file_index, file_total):
    """
    根据已完成的单文件耗时列表，估算剩余时间（ optim：当前文件按整份平均耗时计）。
    file_index / file_total：1-based。
    """
    if not completed_file_seconds or file_total <= 0:
        return None
    try:
        avg = float(sum(completed_file_seconds)) / len(completed_file_seconds)
        fi = int(file_index)
        ft = int(file_total)
        n_left = max(0, ft - fi + 1)
        eta_rem = max(0.0, avg * n_left)
        est_total = max(0.0, avg * ft)
        return {
            "avg_sec_per_file": round(avg, 1),
            "eta_remaining_sec": int(round(eta_rem)),
            "estimated_total_sec": int(round(est_total)),
            "completed_for_eta": len(completed_file_seconds),
        }
    except Exception:
        return None


def compute_progress_percent(file_index, file_total, step_index, max_steps_in_file):
    """
    按「已完成文件数 + 当前文件内步骤」估算 0–99%（完成前不显示 100%，避免单文件任务一开始就满条）。
    file_index / file_total：当前第几个文件 / 总文件数（1-based）
    step_index：当前文件内第几步（1-based）
    max_steps_in_file：该类型单文件最多大约几步
    """
    try:
        ft = max(int(file_total), 1)
        ms = max(int(max_steps_in_file), 1)
        si = min(max(int(step_index), 0), ms)
        base = (file_index - 1) / ft
        cur = (si / ms) / ft
        return min(99, int(round((base + cur) * 100)))
    except Exception:
        return 0


def _batch_progress_stream_meta(
    meta_holder, lite_diag_holder, doc_type, raw_print, dry_run, skip_print
):
    """
    供 SSE 进度推送附带：优先用当前文件的 process_meta_holder（与最终结果一致），
    尚未写入时按类型给出「进行中」占位标签。
    """
    mh = meta_holder if isinstance(meta_holder, dict) else {}
    ld = lite_diag_holder if isinstance(lite_diag_holder, dict) else {}
    label = (mh.get("processing_mode_label") or "").strip()
    if label:
        out = {
            "processingMode": mh.get("processing_mode"),
            "processingModeLabel": label,
        }
        pts = mh.get("modification_points") or []
        if pts:
            out["modificationPoints"] = [
                str(x) for x in pts[:12] if str(x).strip()
            ]
        return out
    dt = doc_type or "unknown"
    if dt == "word":
        if raw_print:
            return {
                "processingMode": "word_raw_print",
                "processingModeLabel": "Word·直接打印（进行中）",
            }
        if dry_run:
            if ld.get("word_lite_mode"):
                return {
                    "processingMode": "word_lite_precheck",
                    "processingModeLabel": "Word·多表格轻量预检（进行中）",
                }
            return {
                "processingMode": "word_full_precheck",
                "processingModeLabel": "Word·全量预检（进行中）",
            }
        if ld.get("word_lite_mode"):
            return {
                "processingMode": "word_lite",
                "processingModeLabel": "Word·多表格轻量（进行中）",
            }
        return {
            "processingMode": "word_full",
            "processingModeLabel": "Word·全量规范化（进行中）",
        }
    if dt == "excel":
        if raw_print:
            return {
                "processingMode": "excel_raw_print",
                "processingModeLabel": "Excel·直接打印（进行中）",
            }
        if dry_run:
            return {
                "processingMode": "excel_precheck",
                "processingModeLabel": "Excel·预检（进行中）",
            }
        if skip_print:
            return {
                "processingMode": "excel_full_save",
                "processingModeLabel": "Excel·全量规范化保存（进行中）",
            }
        return {
            "processingMode": "excel_full_print",
            "processingModeLabel": "Excel·规范化打印（进行中）",
        }
    if dt == "pdf":
        if skip_print:
            return {
                "processingMode": "pdf_pack_only",
                "processingModeLabel": "PDF·原样打包（进行中）",
            }
        if dry_run:
            return {
                "processingMode": "pdf_precheck",
                "processingModeLabel": "PDF·预检（进行中）",
            }
        return {
            "processingMode": "pdf_print",
            "processingModeLabel": "PDF·打印（进行中）",
        }
    return {"processingMode": "unknown", "processingModeLabel": "处理中"}


def _word_lite_not_lite_hint_body(
    n_rows, threshold, *, word_lite_mode=False, raw_print=False
):
    """未进入多表格轻量时的提示正文（不含「【提示】」前缀）。"""
    if raw_print or word_lite_mode:
        return ""
    try:
        nr = int(n_rows) if n_rows is not None else -1
    except (TypeError, ValueError):
        nr = -1
    if nr < 0:
        return (
            f"未进入多表格轻量：表格行数统计失败（仅支持 .docx/.docm 的 XML 快速计数），"
            f"当前阈值={threshold}"
        )
    return f"未进入多表格轻量：文档表格行合计约 {nr} 行，阈值={threshold}"


def _append_word_lite_not_lite_hint(message, batch_path, raw_print, lite_diag_holder):
    """
    超时/中断等路径上补充「未进入轻量」时的表格计数与阈值，并写日志。
    lite_diag_holder：_process_word 写入的 dict（含 ready / n_table_rows / threshold / word_lite_mode）；
    若尚未写入（例如卡在转换 .doc 之前），则仅按 batch_path 做 XML 计数兜底。
    """
    if raw_print or not batch_path:
        return message
    if "未进入多表格轻量" in (message or ""):
        return message
    n_rows = None
    th = None
    word_lite_mode = None
    if isinstance(lite_diag_holder, dict) and lite_diag_holder.get("ready"):
        n_rows = lite_diag_holder.get("n_table_rows")
        if n_rows is None:
            n_rows = lite_diag_holder.get("n_tbl")
        th = lite_diag_holder.get("threshold")
        word_lite_mode = lite_diag_holder.get("word_lite_mode")
    if th is None:
        try:
            from runtime_settings.resolve import get_setting

            th = int(get_setting("WORD_MANY_TABLE_ROWS_LITE_THRESHOLD"))
        except Exception:
            th = 100
    if n_rows is None:
        from doc_handlers.word_handler import docx_table_row_count_hint

        n_rows = docx_table_row_count_hint(batch_path)
    if word_lite_mode is None:
        try:
            word_lite_mode = int(n_rows) >= int(th)
        except (TypeError, ValueError):
            word_lite_mode = False
    body = _word_lite_not_lite_hint_body(
        n_rows, th, word_lite_mode=word_lite_mode, raw_print=raw_print
    )
    if not body:
        return message
    _batch_log_and_web(
        "word not lite (timeout/interrupt path): table_rows=%s threshold=%s file=%s",
        n_rows,
        th,
        os.path.basename(batch_path),
    )
    if "【提示】" in message:
        return f"{message}；{body}"
    return f"{message} 【提示】{body}"


def _process_word(
    path,
    check_formal,
    check_signature,
    accept_revisions,
    printer_name,
    copies,
    dry_run,
    progress_callback=None,
    file_index=0,
    file_total=1,
    auto_fix_formal=True,
    skip_print=False,
    raw_print=False,
    checks_warn_only=False,
    _lite_diag_out=None,
    _effective_path_holder=None,
    _process_meta_holder=None,
):
    from doc_handlers.word_handler import (
        check_formal_document,
        check_cover_signature,
        accept_all_revisions_and_save,
        auto_fix_formal_word,
        print_word_document,
        print_word_with_basic_processing_no_save,
        docx_table_row_count_hint,
        convert_doc_to_docx,
        word_lite_repair_many_tables_save,
    )

    def _manual_print_hint():
        return "自动打印失败：请手动打印原文（本工具已尽量无人工确认）"
    _step_i = [0]

    def report(step):
        _step_i[0] += 1
        pct = compute_progress_percent(
            file_index, file_total, _step_i[0], _PROGRESS_MAX_STEPS_WORD
        )
        if progress_callback:
            progress_callback(
                step, file_index, file_total, os.path.basename(path), pct
            )
    errors = []
    hints = []
    content_notes = [] if checks_warn_only else None
    full_change_log = [] if checks_warn_only else None
    mh = _process_meta_holder

    def _word_pts():
        out = []
        if hints:
            out.extend([str(x) for x in hints])
        if content_notes:
            out.extend([str(x) for x in content_notes])
        return out

    def _meta(mode, label, points=None):
        _batch_meta_set(mh, mode, label, points if points is not None else _word_pts())

    def _fail_word(em, detail=None, *, mode="word_failed", label="Word·处理失败"):
        _batch_meta_set(mh, mode, label, _word_pts() + [str(em)])
        return False, em, detail
    t0 = time.perf_counter()
    auto_fixed = False
    image_risk_mode = False
    print_original_only = False
    risk_reasons = []

    try:
        from runtime_settings.resolve import get_setting

        _lite_th = int(get_setting("WORD_MANY_TABLE_ROWS_LITE_THRESHOLD"))
    except Exception:
        _lite_th = 100
    # .doc → .docx：全模式统一先转（含 raw_print），便于稳定 COM 与 OOXML 路径；输出与源同目录便于批处理临时目录清理。
    try:
        if os.path.splitext(path)[1].lower() == ".doc":
            new_path = convert_doc_to_docx(path)
            if new_path and os.path.isfile(new_path) and os.path.normcase(new_path) != os.path.normcase(path):
                hints.append("已将 .doc 自动转换为 .docx 后再处理（Open XML）")
                path = new_path
                if _effective_path_holder is not None:
                    _effective_path_holder[0] = path
    except Exception as e:
        hints.append("自动转 .docx 失败，继续按原 .doc 处理：" + _format_error(e))
    _n_rows = docx_table_row_count_hint(path) if not raw_print else -1
    word_lite_mode = (not raw_print) and (_n_rows >= _lite_th)
    word_lite_done = False
    cover_done_in_auto_fix = False
    if not raw_print and not word_lite_mode:
        if _n_rows < 0:
            hints.append(
                f"未进入多表格轻量：表格行数统计失败（仅支持 .docx/.docm 的 XML 快速计数），当前阈值={_lite_th}"
            )
        else:
            hints.append(
                f"未进入多表格轻量：文档表格行合计约 {_n_rows} 行，阈值={_lite_th}"
            )
    if raw_print:
        _batch_log_and_web(
            "word process path: raw_print=True ext=%s file=%s",
            os.path.splitext(path)[1].lower() or "?",
            os.path.basename(path),
        )
    else:
        _batch_log_and_web(
            "word process path: table_rows=%s threshold=%s lite=%s ext=%s file=%s",
            _n_rows,
            _lite_th,
            word_lite_mode,
            os.path.splitext(path)[1].lower() or "?",
            os.path.basename(path),
        )

    if _lite_diag_out is not None:
        _lite_diag_out.clear()
        _lite_diag_out.update(
            {
                "ready": True,
                "n_table_rows": _n_rows,
                "threshold": _lite_th,
                "word_lite_mode": word_lite_mode,
                "raw_print": raw_print,
            }
        )

    def _timeout_need_manual_return(em):
        _meta(
            "manual_required",
            "需人工（步骤超时/中断）",
            _word_pts() + [str(em)],
        )
        m = "需人工：" + em
        b = _word_lite_not_lite_hint_body(
            _n_rows, _lite_th, word_lite_mode=word_lite_mode, raw_print=raw_print
        )
        if b and "未进入多表格轻量" not in m:
            _batch_log_and_web(
                "word not lite (step-timeout in _process_word): table_rows=%s threshold=%s file=%s",
                _n_rows,
                _lite_th,
                os.path.basename(path),
            )
            m = m + " 【提示】" + b
        return False, m, None

    if raw_print:
        if skip_print:
            return _fail_word(
                "仅保存模式不支持「直接打印」流程",
                None,
                mode="invalid_mode",
                label="参数冲突·不支持该组合",
            )
        report("直接打印（跳过正式性检查、签字与修订处理）")
        if dry_run:
            _meta("word_raw_print_precheck", "Word·直接打印（预检）")
            return True, "预检通过（将不做任何处理后直接打印）", None
        try:
            out = print_word_document(path, printer_name=printer_name, copies=copies)
            if isinstance(out, str) and out.lower().endswith(".pdf"):
                _meta("word_raw_print", "Word·直接打印（导出 PDF）")
                return True, f"已导出 PDF：{out}", None
            _meta("word_raw_print", "Word·直接打印")
            return True, "已提交打印", None
        except Exception as e:
            em = _format_error(e)
            if "图片完整性风险" in em:
                _meta(
                    "manual_required_print",
                    "需人工·自动打印失败（图片风险）",
                    _word_pts() + [em],
                )
                return False, _manual_print_hint(), None
            return _fail_word(em)

    if word_lite_mode and dry_run:
        report("预检（多表格轻量模式）")
        _meta(
            "word_lite_precheck",
            "Word·多表格轻量（预检）",
            _word_pts()
            + [
                f"表格行合计约 {_n_rows} 行（阈值 {_lite_th}）：仅接受修订、去标黄、统一字体、目录与页码；跳过全文正式性/封面签字/深度排版"
            ],
        )
        return (
            True,
            f"预检通过（表格行合计约 {_n_rows} 行，≥{_lite_th} 行时将仅执行：接受修订、去标黄、统一字体、更新目录与页码；"
            "全文正式性检查、封面签字、深度排版与耗时表格处理将跳过）",
            None,
        )

    if word_lite_mode and not dry_run:
        report("多表格轻量修复（接受修订、去标黄、统一字体、目录与页码）")
        try:
            word_lite_repair_many_tables_save(
                path,
                save_path=path,
                change_notes=content_notes,
                full_change_log=full_change_log,
            )
            word_lite_done = True
            auto_fixed = True
            if hints is not None:
                hints.append(
                    f"多表格轻量模式（表格行合计≥{_lite_th} 行）：已跳过全文正式性检查、封面签字与深度排版步骤"
                )
        except Exception as e:
            em = _format_error(e)
            if "【超时跳过】" in em or "【超时】" in em:
                return _timeout_need_manual_return(em)
            if checks_warn_only:
                hints.append("多表格轻量修复失败：" + em)
            else:
                return _fail_word(em, None, label="Word·多表格轻量失败")

    if word_lite_done:
        cover_done_in_auto_fix = True

    if check_formal and not word_lite_done:
        # 仅保存且非预检：跳过一次全文正式性预扫，单会话直接修复+保存（封面检查可在会话内完成）
        if skip_print and not dry_run and not raw_print:
            report("规范化修复与保存（已跳过单独全文预检）")
            try:
                auto_fix_formal_word(
                    path,
                    change_notes=content_notes,
                    full_change_log=full_change_log,
                    run_cover_check=check_signature,
                )
                auto_fixed = True
                cover_done_in_auto_fix = bool(check_signature)
            except Exception as e:
                em = _format_error(e)
                if "【超时跳过】" in em or "【超时】" in em:
                    return _timeout_need_manual_return(em)
                if "图片完整性风险" in em:
                    if checks_warn_only:
                        hints.append("正式性自动修复因图片风险未完成：" + em)
                    else:
                        report("疑似图片风险：切换保真基础处理（不落盘）")
                        image_risk_mode = True
                        risk_reasons.append("auto_fix_formal_word: " + em)
                elif "页数变化" in em:
                    report("总页数保护：自动修复会改变总页数，已恢复原文件，将直接打印原稿")
                    print_original_only = True
                    if checks_warn_only:
                        hints.append("总页数保护：已跳过会改变页数的自动修复，保留当前文件状态")
                    logger.warning(
                        "page-count guard: skip auto_fix, print original file=%s",
                        os.path.basename(path),
                    )
                else:
                    if checks_warn_only:
                        hints.append("正式性自动修复失败：" + em)
                    else:
                        return _fail_word(em)
        else:
            report("正式性检查（标黄/修订/批注/非黑字体/隐藏文字/水印等）")
            try:
                passed, issues = check_formal_document(path)
            except Exception as e:
                em = _format_error(e)
                if checks_warn_only:
                    passed, issues = False, [em]
                else:
                    return _fail_word(em)
            if not passed:
                if dry_run:
                    joined = "；".join(issues)
                    _meta(
                        "word_formal_precheck_fail",
                        "Word·正式性预检未通过",
                        _word_pts() + [joined],
                    )
                    return False, "正式性检查未通过：" + joined, None
                if checks_warn_only:
                    hints.append(
                        "正式性检查未通过（仍继续尝试自动修复与保存）：" + "；".join(issues)
                    )
                report("自动修复正式性问题")
                try:
                    auto_fix_formal_word(
                        path,
                        change_notes=content_notes,
                        full_change_log=full_change_log,
                    )
                    auto_fixed = True
                except Exception as e:
                    em = _format_error(e)
                    if "【超时跳过】" in em or "【超时】" in em:
                        return _timeout_need_manual_return(em)
                    if "图片完整性风险" in em:
                        if checks_warn_only:
                            hints.append("正式性自动修复因图片风险未完成：" + em)
                        else:
                            report("疑似图片风险：切换保真基础处理（不落盘）")
                            image_risk_mode = True
                            risk_reasons.append("auto_fix_formal_word: " + em)
                    elif "页数变化" in em:
                        report("总页数保护：自动修复会改变总页数，已恢复原文件，将直接打印原稿")
                        print_original_only = True
                        if checks_warn_only:
                            hints.append("总页数保护：已跳过会改变页数的自动修复，保留当前文件状态")
                        logger.warning(
                            "page-count guard: skip auto_fix, print original file=%s",
                            os.path.basename(path),
                        )
                    else:
                        if checks_warn_only:
                            hints.append("正式性自动修复失败：" + em)
                        else:
                            return _fail_word(em)
        if skip_print and image_risk_mode and not checks_warn_only:
            _meta(
                "word_risk_blocked_saveonly",
                "Word·图片风险·仅保存不可用",
                _word_pts()
                + [
                    "正式性/修订处理触发图片风险保护，仅保存模式无法落盘，请改用规范化打印或在 Word 中处理"
                ],
            )
            return (
                False,
                "正式性修复触发图片风险保护，仅保存模式无法落盘修改稿；请改用「规范化打印」或在 Word 中处理",
                None,
            )
    if check_signature and not cover_done_in_auto_fix:
        report("封面签字与日期检查")
        ok, msg = True, ""
        try:
            ok, msg = check_cover_signature(path)
        except Exception as e:
            if checks_warn_only:
                hints.append("封面签字检查执行失败（仍继续）：" + _format_error(e))
                ok = True
            else:
                return _fail_word(
                    _format_error(e), label="Word·封面签字检查失败"
                )
        if not ok:
            if checks_warn_only:
                hints.append("封面签字未通过（仍继续保存）：" + msg)
            else:
                return _fail_word(
                    f"封面签字未通过: {msg}", label="Word·封面签字未通过"
                )
    if not dry_run and accept_revisions and not auto_fixed:
        report("接受修订与规范化")
        try:
            accept_all_revisions_and_save(
                path,
                change_notes=content_notes,
                full_change_log=full_change_log,
            )
        except Exception as e:
            em = _format_error(e)
            if "【超时跳过】" in em or "【超时】" in em:
                return _timeout_need_manual_return(em)
            if "图片完整性风险" in em:
                if checks_warn_only:
                    hints.append("接受修订与规范化因图片风险未完成：" + em)
                else:
                    report("疑似图片风险：切换保真基础处理（不落盘）")
                    image_risk_mode = True
                    risk_reasons.append("accept_all_revisions_and_save: " + em)
            elif "页数变化" in em:
                report("总页数保护：规范化会改变总页数，已恢复原文件，将直接打印原稿")
                print_original_only = True
                if checks_warn_only:
                    hints.append("总页数保护：已跳过会改变页数的接受修订步骤")
                logger.warning(
                    "page-count guard: skip accept/save, print original file=%s",
                    os.path.basename(path),
                )
            else:
                if checks_warn_only:
                    hints.append("接受修订与规范化失败：" + em)
                else:
                    return _fail_word(em, label="Word·接受修订失败")
        if skip_print and image_risk_mode and not checks_warn_only:
            _meta(
                "word_risk_blocked_saveonly",
                "Word·图片风险·仅保存不可用",
                _word_pts()
                + ["接受修订触发图片风险保护，仅保存模式无法完成，请改用规范化打印"],
            )
            return (
                False,
                "接受修订触发图片风险保护，仅保存模式无法完成；请改用「规范化打印」",
                None,
            )
    if skip_print:
        report("处理完成（已保存至文件，未送打印机）")
        base = "已保存修改后的 Word 文件（请从下载包取回）"
        if hints:
            base = base + " 【提示】" + "；".join(hints)
        mod_text = _word_modification_report_text(full_change_log, content_notes)
        if content_notes and not mod_text:
            base = base + " 【内容变更】" + "；".join(content_notes)
        pts = _word_pts()
        if word_lite_done:
            _meta(
                "word_lite_save",
                "Word·多表格轻量（已保存）",
                pts + [f"多表格轻量：表格行合计≥{_lite_th} 行"],
            )
        elif print_original_only:
            _meta(
                "word_page_guard_save",
                "Word·总页数保护（已保存）",
                pts + ["总页数保护：已跳过会改变页数的规范化步骤"],
            )
        else:
            _meta("word_full_save", "Word·全量规范化（已保存）", pts)
        return True, base, mod_text
    if dry_run:
        report("预检完成")
        if image_risk_mode:
            _meta(
                "word_full_precheck_risk",
                "Word·全量预检（图片风险打印路径）",
                _word_pts() + ["后续将走保真基础处理打印、不落盘"],
            )
            return True, "校验通过（图片风险文档将走保真基础处理打印，不落盘）", None
        _meta("word_full_precheck", "Word·全量规范化（预检）", _word_pts())
        return True, "校验通过（模拟，未打印）", None
    report("提交打印")
    try:
        if print_original_only:
            report("总页数保护：直接打印原文件")
            out = print_word_document(path, printer_name=printer_name, copies=copies)
        elif image_risk_mode:
            logger.warning(
                "word risk-mode enter idx=%s/%s file=%s reasons=%s",
                file_index,
                file_total,
                os.path.basename(path),
                " | ".join(risk_reasons) if risk_reasons else "n/a",
            )
            try:
                out = print_word_with_basic_processing_no_save(
                    path,
                    printer_name=printer_name,
                    copies=copies,
                    accept_revisions=accept_revisions,
                    remove_highlights=check_formal,
                )
            except Exception as e2:
                em2 = _format_error(e2)
                if "页数变化" in em2:
                    logger.warning(
                        "no-save basic page guard, fallback print original idx=%s/%s file=%s",
                        file_index,
                        file_total,
                        os.path.basename(path),
                    )
                    report("总页数保护：基础处理将改变总页数，改为直接打印原文件")
                    out = print_word_document(
                        path, printer_name=printer_name, copies=copies
                    )
                else:
                    raise
        else:
            out = print_word_document(path, printer_name=printer_name, copies=copies)
        logger.info(
            "word done idx=%s/%s file=%s auto_fixed=%s risk_mode=%s elapsed=%.2fs",
            file_index,
            file_total,
            os.path.basename(path),
            auto_fixed,
            image_risk_mode,
            time.perf_counter() - t0,
        )
        if isinstance(out, str) and out.lower().endswith(".pdf"):
            if print_original_only:
                _meta(
                    "word_page_guard_print_pdf",
                    "Word·总页数保护（导出 PDF）",
                    _word_pts(),
                )
                return True, f"总页数保护：已导出 PDF（原文件）：{out}", None
            if word_lite_done:
                _meta(
                    "word_lite_print_pdf",
                    "Word·多表格轻量（导出 PDF）",
                    _word_pts(),
                )
            else:
                _meta(
                    "word_full_print_pdf",
                    "Word·全量规范化（导出 PDF）",
                    _word_pts(),
                )
            return True, f"已导出 PDF：{out}", None
        if print_original_only:
            _meta(
                "word_page_guard_print",
                "Word·总页数保护（直接打印原稿）",
                _word_pts(),
            )
            return True, "总页数保护：已直接打印原文件（未做会改变页数的处理）", None
        if image_risk_mode:
            _meta(
                "word_risk_print",
                "Word·保真基础处理打印（不落盘）",
                _word_pts(),
            )
            return True, "已按保真基础处理自动打印（不落盘）", None
        if word_lite_done:
            _meta("word_lite_print", "Word·多表格轻量（打印）", _word_pts())
        else:
            _meta("word_full_print", "Word·全量规范化（打印）", _word_pts())
        return True, "已提交打印", None
    except Exception as e:
        em = _format_error(e)
        logger.error(
            "word print failed idx=%s/%s file=%s risk_mode=%s err=%s",
            file_index,
            file_total,
            os.path.basename(path),
            image_risk_mode,
            em,
        )
        if "图片完整性风险" in em:
            _meta(
                "manual_required_print",
                "需人工·自动打印失败（图片风险）",
                _word_pts() + [em],
            )
            return False, _manual_print_hint(), None
        return _fail_word(em, label="Word·打印失败")


def _process_excel(
    path,
    check_formal,
    accept_revisions,
    printer_name,
    copies,
    dry_run,
    progress_callback=None,
    file_index=0,
    file_total=1,
    auto_fix_formal=True,
    skip_print=False,
    raw_print=False,
    checks_warn_only=False,
    _effective_path_holder=None,
    _process_meta_holder=None,
):
    from doc_handlers.word_handler import FileAbortRequested
    from doc_handlers.excel_handler import (
        ExcelStepTimeout,
        accept_all_changes_and_save,
        auto_fix_formal_excel,
        check_formal_document,
        convert_xls_to_xlsx,
        ensure_font_black,
        excel_normalize_matrix_and_layout,
        print_excel_workbook,
    )

    def _reraise_excel_coop(e):
        if isinstance(e, (FileAbortRequested, ExcelStepTimeout)):
            raise

    _step_i = [0]

    def report(step):
        _step_i[0] += 1
        pct = compute_progress_percent(
            file_index, file_total, _step_i[0], _PROGRESS_MAX_STEPS_EXCEL
        )
        if progress_callback:
            progress_callback(
                step, file_index, file_total, os.path.basename(path), pct
            )
    t0 = time.perf_counter()
    auto_fixed = False
    hints = []
    mh = _process_meta_holder

    def _ex_pts():
        return [str(x) for x in hints] if hints else []

    def _emeta(mode, label, points=None):
        _batch_meta_set(mh, mode, label, points if points is not None else _ex_pts())

    def _fail_ex(em, detail=None, *, mode="excel_failed", label="Excel·处理失败"):
        _batch_meta_set(mh, mode, label, _ex_pts() + [str(em)])
        return False, em, detail

    try:
        if os.path.splitext(path)[1].lower() == ".xls":
            new_path = convert_xls_to_xlsx(path)
            if new_path and os.path.isfile(new_path) and os.path.normcase(new_path) != os.path.normcase(path):
                hints.append("已将 .xls 自动转换为 .xlsx 后再处理（Open XML）")
                path = new_path
                if _effective_path_holder is not None:
                    _effective_path_holder[0] = path
    except Exception as e:
        hints.append("自动将 .xls 转为 .xlsx 失败，继续按原 .xls 处理：" + _format_error(e))

    if raw_print:
        if skip_print:
            return _fail_ex(
                "仅保存模式不支持「直接打印」流程",
                None,
                mode="invalid_mode",
                label="参数冲突·不支持该组合",
            )
        report("直接打印（跳过正式性检查与规范化）")
        if dry_run:
            _emeta("excel_raw_print_precheck", "Excel·直接打印（预检）")
            return True, "预检通过（将不做任何处理后直接打印）", None
        try:
            out = print_excel_workbook(path, printer_name=printer_name, copies=copies)
            if isinstance(out, str) and out.lower().endswith(".pdf"):
                _emeta("excel_raw_print", "Excel·直接打印（导出 PDF）")
                return True, f"已导出 PDF：{out}", None
            _emeta("excel_raw_print", "Excel·直接打印")
            return True, "已提交打印", None
        except Exception as e:
            _reraise_excel_coop(e)
            return _fail_ex("打印失败 " + _format_error(e), label="Excel·直接打印失败")

    if check_formal:
        report("正式性检查（批注/非黑字体；保留表格底色）")
        try:
            passed, issues = check_formal_document(path, check_highlight=False)
        except Exception as e:
            _reraise_excel_coop(e)
            em = _format_error(e)
            if checks_warn_only:
                passed, issues = False, [em]
            else:
                return _fail_ex(em)
        if not passed:
            if dry_run:
                joined = "；".join(issues)
                _emeta(
                    "excel_formal_precheck_fail",
                    "Excel·正式性预检未通过",
                    _ex_pts() + [joined],
                )
                return False, "正式性检查未通过：" + joined, None
            if checks_warn_only:
                hints.append(
                    "正式性检查未通过（仍继续尝试修复与保存）：" + "；".join(issues)
                )
            report("自动修复正式性问题")
            try:
                auto_fix_formal_excel(path)
                auto_fixed = True
            except Exception as e:
                _reraise_excel_coop(e)
                if checks_warn_only:
                    hints.append("自动修复失败（仍继续后续规范化）：" + _format_error(e))
                else:
                    return _fail_ex("自动修复失败 " + _format_error(e), label="Excel·正式性自动修复失败")
    # 检查通过：接受修订后再做规范化；曾自动修复：修复内已含接受修订等，仍走同一套「字体改黑 + 矩阵外排版」收尾，与打印流程最终效果一致
    if not dry_run:
        if not auto_fixed:
            report("接受修订")
            try:
                accept_all_changes_and_save(path, accept_revisions=accept_revisions, remove_highlights=False)
            except Exception as e:
                _reraise_excel_coop(e)
        report("文档规范化")
        try:
            ensure_font_black(path)
        except Exception as e:
            _reraise_excel_coop(e)
            if checks_warn_only:
                hints.append("字体规范化失败：" + _format_error(e))
            else:
                return _fail_ex(
                    "字体规范化失败 " + _format_error(e), label="Excel·字体规范化失败"
                )
        report("风险矩阵外：清除底色、自适应行高（优先内容显示）")
        try:
            excel_normalize_matrix_and_layout(path)
        except Exception as e:
            _reraise_excel_coop(e)
            if checks_warn_only:
                hints.append("矩阵外排版失败：" + _format_error(e))
            else:
                return _fail_ex(
                    "矩阵外排版失败 " + _format_error(e), label="Excel·矩阵外排版失败"
                )
    if skip_print:
        report("处理完成（已保存至文件，未送打印机）")
        base = "已保存修改后的 Excel 文件（请从下载包取回）"
        if hints:
            base = base + " 【提示】" + "；".join(hints)
        pts = _ex_pts()
        pipe = []
        if check_formal:
            pipe.append("正式性检查" + ("；已执行自动修复" if auto_fixed else ""))
        if not dry_run:
            pipe.extend(
                [
                    "接受修订/更改",
                    "工作表字体改黑",
                    "风险矩阵外：清除底色、自适应行高",
                ]
            )
        _emeta("excel_full_save", "Excel·全量规范化（已保存）", pts + pipe)
        mod_x = _excel_modification_report_text(
            hints, check_formal=check_formal, auto_fixed=auto_fixed, dry_run=dry_run
        )
        return True, base, mod_x
    if dry_run:
        report("预检完成")
        _precheck_pts = []
        if check_formal:
            _precheck_pts.append("正式性检查（不通过时可自动修复）")
        _precheck_pts.extend(
            ["接受修订/更改", "字体改黑", "风险矩阵外排版（本步预检为模拟，未改文件）"]
        )
        _emeta(
            "excel_precheck",
            "Excel·全量规范化（预检）",
            _ex_pts() + ["；".join(_precheck_pts)],
        )
        return True, "校验通过（模拟，未打印）", None
    report("提交打印")
    try:
        out = print_excel_workbook(path, printer_name=printer_name, copies=copies)
        logger.info(
            "excel done idx=%s/%s file=%s auto_fixed=%s elapsed=%.2fs",
            file_index,
            file_total,
            os.path.basename(path),
            auto_fixed,
            time.perf_counter() - t0,
        )
        if isinstance(out, str) and out.lower().endswith(".pdf"):
            _emeta("excel_full_print_pdf", "Excel·全量规范化（导出 PDF）", _ex_pts())
            return True, f"已导出 PDF：{out}", None
        _emeta("excel_full_print", "Excel·全量规范化（打印）", _ex_pts())
        return True, "已提交打印", None
    except Exception as e:
        _reraise_excel_coop(e)
        return _fail_ex("打印失败 " + _format_error(e), label="Excel·打印失败")


def _process_pdf(
    path,
    printer_name,
    dry_run,
    progress_callback=None,
    file_index=0,
    file_total=1,
    skip_print=False,
    raw_print=False,
    _process_meta_holder=None,
):
    from doc_handlers.pdf_handler import print_pdf
    _step_i = [0]
    mh = _process_meta_holder

    def _pmeta(mode, label, points=None):
        _batch_meta_set(mh, mode, label, points or [])

    def report(step):
        _step_i[0] += 1
        pct = compute_progress_percent(
            file_index, file_total, _step_i[0], _PROGRESS_MAX_STEPS_PDF
        )
        if progress_callback:
            progress_callback(
                step, file_index, file_total, os.path.basename(path), pct
            )
    if skip_print:
        report("PDF 原样纳入打包（未做正式性处理）")
        _pmeta(
            "pdf_pack_only",
            "PDF·原样打包",
            ["未做 Word/Excel 类规范化，原样纳入 ZIP"],
        )
        return True, "PDF 已原样保留在下载包中", None
    if dry_run:
        report("预检完成")
        _pmeta("pdf_precheck", "PDF·预检", ["将直接送打印机或导出（按系统打印配置）"])
        return True, "校验通过（模拟，未打印）", None
    report("提交打印")
    try:
        out = print_pdf(path, printer_name=printer_name)
        logger.info(
            "pdf done idx=%s/%s file=%s",
            file_index,
            file_total,
            os.path.basename(path),
        )
        if isinstance(out, str) and out.lower().endswith(".pdf"):
            _pmeta("pdf_print_pdf", "PDF·导出/打印（PDF 输出）", [])
            return True, f"已导出 PDF：{out}", None
        _pmeta("pdf_print", "PDF·打印", [])
        return True, "已提交打印", None
    except Exception as e:
        _pmeta("pdf_failed", "PDF·失败", [str(e)])
        return False, str(e), None


def _run_batch_core(
    paths,
    *,
    recursive=False,
    check_formal=True,
    check_signature=True,
    accept_revisions=True,
    printer_name=None,
    copies=1,
    dry_run=False,
    skip_print=False,
    raw_print=False,
    checks_warn_only=None,
    word_content_preserve=True,
    word_preserve_page_count=True,
    word_image_risk_guard=False,
    word_font_profile="mixed",
    word_step_timeout_sec=None,
    word_skip_file_on_timeout=None,
    file_timeout_sec=None,
    progress_callback=None,
    cancel_event=None,
    pause_event=None,
    incremental_output_dir=None,
    relative_names=None,
    incremental_exists_action=None,
    skip_current_event=None,
):
    class _BatchCancelled(Exception):
        pass

    def _check_control():
        if cancel_event is not None:
            try:
                if cancel_event.is_set():
                    raise _BatchCancelled()
            except Exception:
                pass
        if pause_event is not None:
            try:
                while pause_event.is_set():
                    if cancel_event is not None and cancel_event.is_set():
                        raise _BatchCancelled()
                    time.sleep(0.2)
            except _BatchCancelled:
                raise
            except Exception:
                pass

    if checks_warn_only is None:
        checks_warn_only = bool(skip_print and not raw_print)
    # 仅保存模式：不按总页数变化中止/回滚，允许版式与页数随规范化变化
    effective_word_preserve_page_count = (
        word_preserve_page_count if not checks_warn_only else False
    )
    # 仅保存模式：完整非正式性清理（删删除线文本、去水印图形、清分页等），并在结果中汇总内容类变更
    effective_word_content_preserve = (
        word_content_preserve if not checks_warn_only else False
    )
    # 仅保存模式：不自动修补页眉页脚结构，避免改乱模板排版。
    effective_word_header_footer_layout_fix = (not checks_warn_only)
    # 仅保存模式：开启图片保全，避免域更新触发链接图片失效（红叉）。
    effective_word_preserve_linked_images = bool(checks_warn_only)
    _wfp = (word_font_profile or "mixed").strip().lower()
    if _wfp not in ("chinese", "english", "mixed"):
        _wfp = "mixed"
    t_start = time.perf_counter()
    files = _collect_files(paths, recursive=recursive)

    rel_for_each_file = []
    if relative_names is not None and len(relative_names) == len(paths):
        path_to_rel = {os.path.abspath(p): relative_names[i] for i, p in enumerate(paths)}
        for cf in files:
            rel_for_each_file.append(path_to_rel.get(os.path.abspath(cf)))
    else:
        if relative_names is not None and relative_names and len(relative_names) != len(paths):
            logger.warning(
                "incremental_output_dir: relative_names length %s != paths %s, use basenames only",
                len(relative_names),
                len(paths),
            )
        rel_for_each_file = [None] * len(files)

    def _rel_parts(rel):
        if not rel:
            return None
        parts = []
        for p in str(rel).replace("\\", "/").split("/"):
            p = p.strip()
            if not p or p == "." or p == "..":
                continue
            parts.append(p)
        return parts if parts else None

    if not files:
        return {
            "total": 0,
            "ok": 0,
            "failed": 0,
            "details": [],
            "batch_elapsed_sec": 0.0,
            "avg_sec_per_file": None,
            "eta_summary": None,
        }
    try:
        from doc_handlers.word_handler import (
            apply_resolved_word_base_settings,
            set_runtime_options as set_word_runtime_options,
        )

        apply_resolved_word_base_settings()
        set_word_runtime_options(
            word_content_preserve=effective_word_content_preserve,
            word_preserve_page_count=effective_word_preserve_page_count,
            word_image_risk_guard=word_image_risk_guard,
            word_preserve_linked_images=effective_word_preserve_linked_images,
            word_step_timeout_sec=word_step_timeout_sec,
            word_skip_file_on_timeout=word_skip_file_on_timeout,
            word_header_footer_layout_fix=effective_word_header_footer_layout_fix,
            word_font_profile=_wfp,
        )
    except Exception:
        pass
    try:
        from doc_handlers.excel_handler import set_runtime_options as set_excel_runtime_options

        set_excel_runtime_options(
            excel_font_profile=_wfp,
            excel_step_timeout_sec=word_step_timeout_sec,
            excel_skip_file_on_timeout=word_skip_file_on_timeout,
        )
    except Exception:
        pass
    try:
        from runtime_settings.resolve import get_setting

        _bud = str(get_setting("AIPRINTWORD_WORD_BACKUP_TEMP_DIR") or "").strip()
        if _bud:
            logger.info("Word 页数保护备份目录: %s", _bud)
    except Exception:
        pass

    results = []
    ok_count = 0
    total = len(files)
    logger.info(
        "run_batch start total=%s check_formal=%s check_signature=%s accept_revisions=%s dry_run=%s skip_print=%s raw_print=%s checks_warn_only=%s preserve=%s page_guard=%s image_guard=%s doc_font_profile=%s",
        total,
        check_formal,
        check_signature,
        accept_revisions,
        dry_run,
        skip_print,
        raw_print,
        checks_warn_only,
        effective_word_content_preserve,
        effective_word_preserve_page_count,
        word_image_risk_guard,
        _wfp,
    )
    completed_file_seconds = []
    progress_ctx = {"meta": None, "lite": None, "doc_type": None}

    def wrap_progress(step, file_index, file_total, file_name, percent=None):
        _check_control()
        eta = _batch_eta_payload(completed_file_seconds, file_index, file_total)
        if progress_callback:
            pm = _batch_progress_stream_meta(
                progress_ctx.get("meta"),
                progress_ctx.get("lite"),
                progress_ctx.get("doc_type"),
                raw_print,
                dry_run,
                skip_print,
            )
            progress_callback(
                step,
                file_index,
                file_total,
                file_name,
                percent,
                eta,
                pm,
            )

    effective_progress = wrap_progress if progress_callback else None

    _exists_act = (incremental_exists_action or "overwrite").strip().lower()
    if _exists_act not in ("overwrite", "skip"):
        _exists_act = "overwrite"

    # 「跳过」策略：在打开 Word/Excel 之前一次性扫完全批目标路径，不在循环内逐份判断
    skip_indices_output_exists = set()
    if incremental_output_dir and _exists_act == "skip" and not dry_run:
        for j, fp in enumerate(files):
            parts_chk = _rel_parts(
                rel_for_each_file[j] if j < len(rel_for_each_file) else None
            )
            if parts_chk:
                dest_chk = os.path.join(incremental_output_dir, *parts_chk)
            else:
                dest_chk = os.path.join(
                    incremental_output_dir, os.path.basename(fp)
                )
            if os.path.isfile(dest_chk):
                skip_indices_output_exists.add(j)
        if skip_indices_output_exists:
            logger.info(
                "run_batch output_exists prescan: skip %s/%s files (dir=%s)",
                len(skip_indices_output_exists),
                total,
                incremental_output_dir,
            )

    cancelled = False
    cancelled_at_index = None
    for idx, f in enumerate(files):
        try:
            _check_control()
        except _BatchCancelled:
            cancelled = True
            cancelled_at_index = idx + 1
            break
        if skip_current_event is not None:
            try:
                skip_current_event.clear()
            except Exception:
                pass
        f_start = time.perf_counter()
        if idx in skip_indices_output_exists:
            parts_chk = _rel_parts(
                rel_for_each_file[idx] if idx < len(rel_for_each_file) else None
            )
            if parts_chk:
                dest_chk = os.path.join(incremental_output_dir, *parts_chk)
            else:
                dest_chk = os.path.join(
                    incremental_output_dir, os.path.basename(f)
                )
            entry = {
                "path": f,
                "success": False,
                "message": "已跳过：批次开始时检测到即时保存目录已存在同名文件，未打开处理",
                "modification_detail": None,
                "skipped_output_exists": True,
                "skipped_output_dest": dest_chk,
            }
            _batch_merge_entry_processing(
                entry,
                {},
                doc_type=_get_type(f),
                raw_print=raw_print,
                dry_run=dry_run,
                skip_print=skip_print,
                lite_diag_holder={},
                watchdog={},
            )
            progress_ctx["doc_type"] = _get_type(f)
            progress_ctx["lite"] = {}
            progress_ctx["meta"] = {
                "processing_mode": "skipped_output_exists",
                "processing_mode_label": "已跳过·即时目录已有同名文件",
                "modification_points": [],
            }
            if effective_progress:
                effective_progress(
                    "已跳过：即时目录已有同名文件",
                    idx + 1,
                    total,
                    os.path.basename(f),
                    compute_progress_percent(idx + 1, total, 1, 1),
                )
            results.append(entry)
            f_elapsed = time.perf_counter() - f_start
            completed_file_seconds.append(f_elapsed)
            logger.info(
                "run_batch skip_output_exists idx=%s/%s file=%s dest=%s",
                idx + 1,
                total,
                os.path.basename(f),
                dest_chk,
            )
            continue
        try:
            if os.path.splitext(f)[1].lower() in WORD_EXT:
                _emit_word_lite_precheck(f, raw_print)
        except Exception as e:
            logger.warning(
                "word precheck failed: %s file=%s",
                e,
                os.path.basename(f),
            )
        doc_type = _get_type(f)
        if doc_type != "word":
            try:
                from doc_handlers.word_handler import word_batch_session_end

                word_batch_session_end()
            except Exception:
                pass
        else:
            try:
                from doc_handlers.word_handler import word_batch_session_begin

                word_batch_session_begin()
            except Exception:
                pass
        kwargs = dict(
            progress_callback=effective_progress,
            file_index=idx + 1,
            file_total=total,
        )
        watchdog_stop = threading.Event()
        watchdog = {"timed_out": False, "cancel_killed": False, "manual_skip": False}
        file_abort_event = threading.Event()

        def _watchdog_loop():
            # 单文件总时长超时 + 取消 + 手动跳过当前文件：先置 abort_event、再 taskkill
            t0 = time.perf_counter()
            paused_total = 0.0
            paused_at = None
            while not watchdog_stop.is_set():
                # pause 不计入超时
                if pause_event is not None and pause_event.is_set():
                    if paused_at is None:
                        paused_at = time.perf_counter()
                else:
                    if paused_at is not None:
                        paused_total += time.perf_counter() - paused_at
                        paused_at = None
                if cancel_event is not None and cancel_event.is_set():
                    if doc_type in ("word", "excel"):
                        try:
                            file_abort_event.set()
                        except Exception:
                            pass
                    try:
                        if doc_type == "word":
                            from doc_handlers.word_handler import kill_last_word_app

                            kill_last_word_app("cancel")
                        elif doc_type == "excel":
                            from doc_handlers.excel_handler import kill_last_excel_app

                            kill_last_excel_app("cancel")
                    except Exception:
                        pass
                    watchdog["cancel_killed"] = True
                    return
                if skip_current_event is not None and skip_current_event.is_set():
                    if doc_type in ("word", "excel"):
                        try:
                            file_abort_event.set()
                        except Exception:
                            pass
                    try:
                        if doc_type == "word":
                            from doc_handlers.word_handler import kill_last_word_app

                            kill_last_word_app("manual_skip")
                        elif doc_type == "excel":
                            from doc_handlers.excel_handler import kill_last_excel_app

                            kill_last_excel_app("manual_skip")
                    except Exception:
                        pass
                    watchdog["manual_skip"] = True
                    try:
                        skip_current_event.clear()
                    except Exception:
                        pass
                    return
                if file_timeout_sec is not None:
                    try:
                        limit = float(file_timeout_sec)
                    except Exception:
                        limit = None
                    if limit and limit > 0:
                        elapsed = (time.perf_counter() - t0) - paused_total
                        if elapsed >= limit:
                            watchdog["timed_out"] = True
                            if doc_type in ("word", "excel"):
                                try:
                                    file_abort_event.set()
                                except Exception:
                                    pass
                            try:
                                if doc_type == "word":
                                    from doc_handlers.word_handler import kill_last_word_app

                                    kill_last_word_app("file_timeout")
                                elif doc_type == "excel":
                                    from doc_handlers.excel_handler import kill_last_excel_app

                                    kill_last_excel_app("file_timeout")
                            except Exception:
                                pass
                            return
                time.sleep(0.5)

        t_watch = threading.Thread(target=_watchdog_loop, daemon=True)
        t_watch.start()

        lite_diag_holder = {}
        process_meta_holder = {}
        progress_ctx["meta"] = process_meta_holder
        progress_ctx["lite"] = lite_diag_holder
        progress_ctx["doc_type"] = doc_type
        result_holder = {}
        exc_holder = {}
        effective_path_holder = [f]

        def _run_doc_task():
            try:
                if doc_type == "word":
                    result_holder["v"] = _process_word(
                        f,
                        check_formal=check_formal,
                        check_signature=check_signature,
                        accept_revisions=accept_revisions,
                        printer_name=printer_name,
                        copies=copies,
                        dry_run=dry_run,
                        skip_print=skip_print,
                        raw_print=raw_print,
                        checks_warn_only=checks_warn_only,
                        _lite_diag_out=lite_diag_holder,
                        _effective_path_holder=effective_path_holder,
                        _process_meta_holder=process_meta_holder,
                        **kwargs,
                    )
                elif doc_type == "excel":
                    result_holder["v"] = _process_excel(
                        f,
                        check_formal=check_formal,
                        accept_revisions=accept_revisions,
                        printer_name=printer_name,
                        copies=copies,
                        dry_run=dry_run,
                        skip_print=skip_print,
                        raw_print=raw_print,
                        checks_warn_only=checks_warn_only,
                        _effective_path_holder=effective_path_holder,
                        _process_meta_holder=process_meta_holder,
                        **kwargs,
                    )
                elif doc_type == "pdf":
                    result_holder["v"] = _process_pdf(
                        f,
                        printer_name=printer_name,
                        dry_run=dry_run,
                        skip_print=skip_print,
                        raw_print=raw_print,
                        _process_meta_holder=process_meta_holder,
                        **kwargs,
                    )
                else:
                    result_holder["v"] = (False, "未知类型", None)
            except _BatchCancelled:
                exc_holder["batch_cancel"] = True
            except Exception as e:
                exc_holder["e"] = e

        try:
            if doc_type == "word":
                from doc_handlers.word_handler import (
                    FileAbortRequested,
                    clear_word_file_abort_event,
                    kill_last_word_app,
                    set_word_file_abort_event,
                )

                set_word_file_abort_event(file_abort_event)
                wt = threading.Thread(target=_run_doc_task, daemon=True)
                wt.start()
                try:
                    while wt.is_alive():
                        try:
                            _check_control()
                        except _BatchCancelled:
                            try:
                                file_abort_event.set()
                            except Exception:
                                pass
                            try:
                                kill_last_word_app("cancel")
                            except Exception:
                                pass
                            wt.join(timeout=120)
                            raise
                        if (
                            watchdog.get("timed_out")
                            or watchdog.get("cancel_killed")
                            or watchdog.get("manual_skip")
                        ):
                            break
                        wt.join(timeout=0.35)
                    if wt.is_alive():
                        try:
                            file_abort_event.set()
                        except Exception:
                            pass
                        try:
                            kill_last_word_app("stuck_after_signal")
                        except Exception:
                            pass
                        wt.join(timeout=90)
                finally:
                    clear_word_file_abort_event()
                if exc_holder.get("batch_cancel"):
                    raise _BatchCancelled()
                if "e" in exc_holder:
                    e = exc_holder["e"]
                    if isinstance(e, FileAbortRequested):
                        success, message, mod_detail = (
                            False,
                            f"需人工：【超时跳过】{e}",
                            None,
                        )
                    else:
                        success, message, mod_detail = False, _format_error(e), None
                elif "v" in result_holder:
                    success, message, mod_detail = result_holder["v"]
                else:
                    success, message, mod_detail = (
                        False,
                        "处理被中断（Word 未在预期时间内结束，可能仍卡住或已被结束进程）",
                        None,
                    )
            elif doc_type == "excel":
                from doc_handlers.word_handler import FileAbortRequested
                from doc_handlers.excel_handler import (
                    ExcelStepTimeout,
                    clear_excel_file_abort_event,
                    set_excel_file_abort_event,
                )

                set_excel_file_abort_event(file_abort_event)
                wt = threading.Thread(target=_run_doc_task, daemon=True)
                wt.start()
                while wt.is_alive():
                    try:
                        _check_control()
                    except _BatchCancelled:
                        try:
                            file_abort_event.set()
                        except Exception:
                            pass
                        try:
                            from doc_handlers.excel_handler import kill_last_excel_app

                            kill_last_excel_app("cancel")
                        except Exception:
                            pass
                        wt.join(timeout=120)
                        raise
                    if (
                        watchdog.get("timed_out")
                        or watchdog.get("cancel_killed")
                        or watchdog.get("manual_skip")
                    ):
                        break
                    wt.join(timeout=0.35)
                if wt.is_alive():
                    try:
                        file_abort_event.set()
                    except Exception:
                        pass
                    try:
                        from doc_handlers.excel_handler import kill_last_excel_app

                        kill_last_excel_app("stuck_after_signal")
                    except Exception:
                        pass
                    wt.join(timeout=90)
                clear_excel_file_abort_event()
                if exc_holder.get("batch_cancel"):
                    raise _BatchCancelled()
                if "e" in exc_holder:
                    e = exc_holder["e"]
                    if isinstance(e, FileAbortRequested):
                        success, message, mod_detail = (
                            False,
                            f"需人工：【超时跳过】{e}",
                            None,
                        )
                    elif isinstance(e, ExcelStepTimeout):
                        success, message, mod_detail = False, _format_error(e), None
                    else:
                        success, message, mod_detail = False, _format_error(e), None
                elif "v" in result_holder:
                    success, message, mod_detail = result_holder["v"]
                else:
                    success, message, mod_detail = (
                        False,
                        "处理被中断（Excel 线程未在预期时间内结束）",
                        None,
                    )
            elif doc_type == "pdf":
                _run_doc_task()
                if exc_holder.get("batch_cancel"):
                    raise _BatchCancelled()
                if "e" in exc_holder:
                    success, message, mod_detail = False, _format_error(exc_holder["e"]), None
                else:
                    success, message, mod_detail = result_holder.get(
                        "v", (False, "未知", None)
                    )
            else:
                success, message, mod_detail = False, "未知类型", None
        except _BatchCancelled:
            cancelled = True
            cancelled_at_index = idx + 1
            break
        finally:
            try:
                from doc_handlers.excel_handler import clear_excel_file_abort_event
                from doc_handlers.word_handler import clear_word_file_abort_event

                clear_word_file_abort_event()
                clear_excel_file_abort_event()
            except Exception:
                pass
            watchdog_stop.set()
            try:
                t_watch.join(timeout=1.0)
            except Exception:
                pass

        if watchdog.get("timed_out"):
            success = False
            message = f"需人工：【超时跳过】单文件总耗时超过 {int(file_timeout_sec)} 秒"
            mod_detail = None
        if watchdog.get("manual_skip"):
            success = False
            message = (
                "已手动跳过：已尝试结束 Word/Excel 进程（PDF 无法在后台打断，请等待该文件结束或使用取消整批）"
            )
            mod_detail = None
        if (
            doc_type == "word"
            and not raw_print
            and not success
            and (
                watchdog.get("timed_out")
                or "【超时跳过】" in (message or "")
                or "处理被中断（Word 未在预期时间内结束" in (message or "")
            )
        ):
            message = _append_word_lite_not_lite_hint(
                message or "", f, raw_print, lite_diag_holder
            )
        pmode = (process_meta_holder or {}).get("processing_mode")
        if (
            success
            and doc_type == "word"
            and skip_print
            and not raw_print
            and not dry_run
            and pmode != "word_page_guard_save"
        ):
            try:
                from doc_handlers.word_handler import verify_word_saved_no_pending_revisions

                vp, vm = verify_word_saved_no_pending_revisions(
                    os.path.abspath(effective_path_holder[0])
                )
                if not vp:
                    success = False
                    message = f"落盘校验未通过：{vm}"
                    if mod_detail:
                        mod_detail = str(mod_detail).strip() + "\n\n【校验】" + vm
                    else:
                        mod_detail = "【校验】" + vm
            except Exception as ex:
                logger.warning("word post-save verify error: %s", ex)
        entry = {
            "path": f,
            "success": success,
            "message": message,
            "modification_detail": mod_detail,
            "processed_path": os.path.abspath(effective_path_holder[0]),
        }
        _batch_merge_entry_processing(
            entry,
            process_meta_holder,
            doc_type=doc_type,
            raw_print=raw_print,
            dry_run=dry_run,
            skip_print=skip_print,
            lite_diag_holder=lite_diag_holder,
            watchdog=watchdog,
        )
        if watchdog.get("timed_out"):
            entry["timeout_skip"] = True
        if (not entry.get("timeout_skip")) and ("【超时跳过】" in (message or "")):
            entry["timeout_skip"] = True
        if watchdog.get("manual_skip"):
            entry["manual_skip"] = True
        if success and incremental_output_dir and not dry_run:
            try:
                src_e = effective_path_holder[0]
                parts = _rel_parts(
                    rel_for_each_file[idx] if idx < len(rel_for_each_file) else None
                )
                if parts:
                    stem, _ = os.path.splitext(parts[-1])
                    _, nxt = os.path.splitext(src_e)
                    parts = parts[:-1] + [stem + nxt if nxt else parts[-1]]
                    dest = os.path.join(incremental_output_dir, *parts)
                else:
                    stem, _ = os.path.splitext(os.path.basename(f))
                    _, nxt = os.path.splitext(src_e)
                    dest = os.path.join(
                        incremental_output_dir,
                        stem + (nxt or os.path.splitext(os.path.basename(f))[1]),
                    )
                dest_parent = os.path.dirname(dest)
                if dest_parent:
                    os.makedirs(dest_parent, exist_ok=True)
                # 即时保存目录下若已存在同名文件则覆盖
                if os.path.isfile(dest):
                    try:
                        os.remove(dest)
                    except OSError:
                        pass
                shutil.copy2(src_e, dest)
                entry["saved_to"] = dest
            except Exception as e:
                logger.warning("incremental save failed for %s: %s", f, e)
                entry["save_error"] = str(e)
        results.append(entry)
        if success:
            ok_count += 1
        f_elapsed = time.perf_counter() - f_start
        completed_file_seconds.append(f_elapsed)
        n_done = idx + 1
        eta_line = ""
        if n_done >= 1 and n_done < total:
            avg_so_far = sum(completed_file_seconds) / n_done
            rem_files = total - n_done
            eta_rem = avg_so_far * rem_files
            eta_line = " 预计剩余≈%s（按已完成 %s 份平均 %.1fs/份）" % (
                _format_duration_cn(eta_rem),
                n_done,
                avg_so_far,
            )
        logger.info(
            "run_batch file_done idx=%s/%s type=%s success=%s elapsed=%.2fs file=%s msg=%s%s",
            idx + 1,
            total,
            doc_type,
            success,
            f_elapsed,
            os.path.basename(f),
            message,
            eta_line,
        )
        if doc_type == "word":
            try:
                from doc_handlers.word_handler import word_batch_session_end

                nxt = _next_processed_doc_type(
                    files, idx, skip_indices_output_exists
                )
                if nxt != "word":
                    word_batch_session_end()
            except Exception:
                pass

    batch_elapsed = time.perf_counter() - t_start
    avg_all = None
    if completed_file_seconds:
        avg_all = round(sum(completed_file_seconds) / len(completed_file_seconds), 2)
    eta_summary = None
    if total > 0 and batch_elapsed > 0:
        eta_summary = (
            f"本批 {total} 个文件，总耗时 {_format_duration_cn(batch_elapsed)}"
        )
        if avg_all is not None:
            eta_summary += f"，平均每个文件约 {_format_duration_cn(avg_all)}"

    logger.info(
        "run_batch done total=%s ok=%s failed=%s elapsed=%.2fs avg_per_file=%s",
        len(files),
        ok_count,
        len(files) - ok_count,
        batch_elapsed,
        avg_all if avg_all is not None else "n/a",
    )
    return {
        "total": len(files),
        "ok": ok_count,
        "failed": len(files) - ok_count,
        "details": results,
        "dry_run": dry_run,
        "skip_print": skip_print,
        "raw_print": raw_print,
        "batch_elapsed_sec": round(batch_elapsed, 2),
        "avg_sec_per_file": avg_all,
        "eta_summary": eta_summary,
        "cancelled": cancelled,
        "cancelled_at_index": cancelled_at_index,
        "incremental_output_dir": incremental_output_dir,
    }


def run_batch(*args, **kwargs):
    """
    对 paths 中的文件/目录执行批量检查与打印。
    skip_print=True：完成 Word/Excel 规范化与保存，不送打印机（供打包下载）。
    raw_print=True：跳过检查与修订处理，直接打印（Word/Excel/PDF）。
    checks_warn_only=None 时与 skip_print 相同：检查未通过只记入提示，仍尽量完成后续处理并保存。
    progress_callback(step, file_index, file_total, file_name, percent=None, eta=None, processing_meta=None)
    用于上报当前任务与进度；percent 为 0–99；eta 为可选 dict；processing_meta 为可选 dict，
    含 processingMode、processingModeLabel、modificationPoints（供 SSE 显示当前处理模式）。

    与 Word/Excel 的 COM 会话串行化：多浏览器/多用户同时提交时，后到的任务会排队等待，避免「被呼叫方拒绝接收呼叫」。
    """
    logger.info("run_batch: 等待 Office 自动化锁（若有其它批任务正在运行将在此排队）…")
    with _OFFICE_COM_LOCK:
        logger.info("run_batch: 已获取 Office 锁，开始执行")
        try:
            return _run_batch_core(*args, **kwargs)
        finally:
            try:
                from doc_handlers.word_handler import word_batch_session_end

                word_batch_session_end()
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(
        description="批量打印 Word/Excel/PDF：检查封面签字、自动接受修订后打印"
    )
    parser.add_argument(
        "paths",
        nargs="+",
        help="文件或目录路径（可多个）",
    )
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="对目录递归查找文档",
    )
    parser.add_argument(
        "--no-formal-check",
        action="store_true",
        help="不做正式性检查（标黄、修订、批注、非黑字体等）",
    )
    parser.add_argument(
        "--no-signature-check",
        action="store_true",
        help="不检查 Word 封面电子签字",
    )
    parser.add_argument(
        "--no-accept-revisions",
        action="store_true",
        help="不自动接受修订（有修订时将跳过该文档）",
    )
    parser.add_argument(
        "-p", "--printer",
        default=None,
        help="打印机名称，不指定则使用默认打印机",
    )
    parser.add_argument(
        "-n", "--copies",
        type=int,
        default=1,
        help="打印份数（默认 1）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅检查并接受修订，不实际打印",
    )
    args = parser.parse_args()

    result = run_batch(
        args.paths,
        recursive=args.recursive,
        check_formal=not args.no_formal_check,
        check_signature=not args.no_signature_check,
        accept_revisions=not args.no_accept_revisions,
        printer_name=args.printer or None,
        copies=args.copies,
        dry_run=args.dry_run,
    )

    for d in result["details"]:
        status = "OK" if d["success"] else "FAIL"
        print(f"[{status}] {d['path']}")
        print(f"     {d['message']}")
    print("-" * 50)
    print(f"合计: {result['total']} 个文件, 成功 {result['ok']}, 失败 {result['failed']}")

    sys.exit(0 if result["failed"] == 0 else 1)


if __name__ == "__main__":
    main()
