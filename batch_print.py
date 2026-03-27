# -*- coding: utf-8 -*-
"""
批量打印工具：在打印前检查封面签字（Word）、修订记录，并自动接受所有修订后打印。
支持：.doc/.docx、.xls/.xlsx、.pdf
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
import time

# 支持的扩展名（小写）
WORD_EXT = {".doc", ".docx", ".docm"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm"}
PDF_EXT = {".pdf"}

# COM「发生意外」错误码，WPS/Office/Excel 易触发
_COM_ERR_CODES = (-2147352567, -2147467259, -2146827864)
_COM_ERR_MSG = "文档处理时发生意外，请确认 WPS/Excel 已安装、文档未被占用，或稍后重试。"

logger = logging.getLogger("aiprintword.batch")


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


def build_batch_modification_zip_text(details):
    """
    将 run_batch 返回的 details（须已含 filename）拼成 ZIP 内「修改明细.txt」正文。
    仅包含有 modification_detail 的条目。
    """
    blocks = []
    for d in details:
        fn = (d.get("filename") or os.path.basename(d.get("path") or "")).strip()
        md = d.get("modification_detail")
        if md is None or not str(md).strip():
            continue
        blocks.append("=" * 64)
        blocks.append(fn)
        blocks.append("=" * 64)
        blocks.append(str(md).strip())
        blocks.append("")
    text = "\n".join(blocks).strip()
    return text if text else None


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


# 各类型单文件内「步骤」上界（用于进度条；实际 report 次数可少不可多，超出则按满步计）
_PROGRESS_MAX_STEPS_WORD = 10
_PROGRESS_MAX_STEPS_EXCEL = 12
_PROGRESS_MAX_STEPS_PDF = 4


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
):
    from doc_handlers.word_handler import (
        check_formal_document,
        check_cover_signature,
        accept_all_revisions_and_save,
        auto_fix_formal_word,
        print_word_document,
        print_word_with_basic_processing_no_save,
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
    t0 = time.perf_counter()
    auto_fixed = False
    image_risk_mode = False
    print_original_only = False
    risk_reasons = []

    if raw_print:
        if skip_print:
            return False, "仅保存模式不支持「直接打印」流程", None
        report("直接打印（跳过正式性检查、签字与修订处理）")
        if dry_run:
            return True, "预检通过（将不做任何处理后直接打印）", None
        try:
            out = print_word_document(path, printer_name=printer_name, copies=copies)
            if isinstance(out, str) and out.lower().endswith(".pdf"):
                return True, f"已导出 PDF：{out}", None
            return True, "已提交打印", None
        except Exception as e:
            em = _format_error(e)
            if "图片完整性风险" in em:
                return False, _manual_print_hint(), None
            return False, em, None

    if check_formal:
        report("正式性检查（标黄/修订/批注/非黑字体/隐藏文字/水印等）")
        try:
            passed, issues = check_formal_document(path)
        except Exception as e:
            em = _format_error(e)
            if checks_warn_only:
                passed, issues = False, [em]
            else:
                return False, em, None
        if not passed:
            if dry_run:
                return False, "正式性检查未通过：" + "；".join(issues), None
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
                        return False, em, None
        if skip_print and image_risk_mode and not checks_warn_only:
            return (
                False,
                "正式性修复触发图片风险保护，仅保存模式无法落盘修改稿；请改用「规范化打印」或在 Word 中处理",
                None,
            )
    if check_signature:
        report("封面签字与日期检查")
        ok, msg = True, ""
        try:
            ok, msg = check_cover_signature(path)
        except Exception as e:
            if checks_warn_only:
                hints.append("封面签字检查执行失败（仍继续）：" + _format_error(e))
                ok = True
            else:
                return False, _format_error(e), None
        if not ok:
            if checks_warn_only:
                hints.append("封面签字未通过（仍继续保存）：" + msg)
            else:
                return False, f"封面签字未通过: {msg}", None
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
                    return False, em, None
        if skip_print and image_risk_mode and not checks_warn_only:
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
        if mod_text:
            base = base + " 【内容变更】完整逐条明细见下载包内「修改明细.txt」"
        elif content_notes:
            base = base + " 【内容变更】" + "；".join(content_notes)
        return True, base, mod_text
    if dry_run:
        report("预检完成")
        if image_risk_mode:
            return True, "校验通过（图片风险文档将走保真基础处理打印，不落盘）", None
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
                return True, f"总页数保护：已导出 PDF（原文件）：{out}", None
            return True, f"已导出 PDF：{out}", None
        if print_original_only:
            return True, "总页数保护：已直接打印原文件（未做会改变页数的处理）", None
        if image_risk_mode:
            return True, "已按保真基础处理自动打印（不落盘）", None
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
            return False, _manual_print_hint(), None
        return False, em, None


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
):
    from doc_handlers.excel_handler import (
        check_formal_document,
        accept_all_changes_and_save,
        ensure_font_black,
        auto_fix_formal_excel,
        excel_normalize_matrix_and_layout,
        print_excel_workbook,
    )
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

    if raw_print:
        if skip_print:
            return False, "仅保存模式不支持「直接打印」流程", None
        report("直接打印（跳过正式性检查与规范化）")
        if dry_run:
            return True, "预检通过（将不做任何处理后直接打印）", None
        try:
            out = print_excel_workbook(path, printer_name=printer_name, copies=copies)
            if isinstance(out, str) and out.lower().endswith(".pdf"):
                return True, f"已导出 PDF：{out}", None
            return True, "已提交打印", None
        except Exception as e:
            return False, "打印失败 " + _format_error(e), None

    if check_formal:
        report("正式性检查（批注/非黑字体；保留表格底色）")
        try:
            passed, issues = check_formal_document(path, check_highlight=False)
        except Exception as e:
            em = _format_error(e)
            if checks_warn_only:
                passed, issues = False, [em]
            else:
                return False, em, None
        if not passed:
            if dry_run:
                return False, "正式性检查未通过：" + "；".join(issues), None
            if checks_warn_only:
                hints.append(
                    "正式性检查未通过（仍继续尝试修复与保存）：" + "；".join(issues)
                )
            report("自动修复正式性问题")
            try:
                auto_fix_formal_excel(path)
                auto_fixed = True
            except Exception as e:
                if checks_warn_only:
                    hints.append("自动修复失败（仍继续后续规范化）：" + _format_error(e))
                else:
                    return False, "自动修复失败 " + _format_error(e), None
    # 检查通过：接受修订后再做规范化；曾自动修复：修复内已含接受修订等，仍走同一套「字体改黑 + 矩阵外排版」收尾，与打印流程最终效果一致
    if not dry_run:
        if not auto_fixed:
            report("接受修订")
            try:
                accept_all_changes_and_save(path, accept_revisions=accept_revisions, remove_highlights=False)
            except Exception:
                pass
        report("文档规范化")
        try:
            ensure_font_black(path)
        except Exception as e:
            if checks_warn_only:
                hints.append("字体规范化失败：" + _format_error(e))
            else:
                return False, "字体规范化失败 " + _format_error(e), None
        report("风险矩阵外：清除底色、自适应行高（优先内容显示）")
        try:
            excel_normalize_matrix_and_layout(path)
        except Exception as e:
            if checks_warn_only:
                hints.append("矩阵外排版失败：" + _format_error(e))
            else:
                return False, "矩阵外排版失败 " + _format_error(e), None
    if skip_print:
        report("处理完成（已保存至文件，未送打印机）")
        base = "已保存修改后的 Excel 文件（请从下载包取回）"
        if hints:
            base = base + " 【提示】" + "；".join(hints)
        return True, base, None
    if dry_run:
        report("预检完成")
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
            return True, f"已导出 PDF：{out}", None
        return True, "已提交打印", None
    except Exception as e:
        return False, "打印失败 " + _format_error(e), None


def _process_pdf(
    path,
    printer_name,
    dry_run,
    progress_callback=None,
    file_index=0,
    file_total=1,
    skip_print=False,
    raw_print=False,
):
    from doc_handlers.pdf_handler import print_pdf
    _step_i = [0]

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
        return True, "PDF 已原样保留在下载包中", None
    if dry_run:
        report("预检完成")
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
            return True, f"已导出 PDF：{out}", None
        return True, "已提交打印", None
    except Exception as e:
        return False, str(e), None


def run_batch(
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
    progress_callback=None,
):
    """
    对 paths 中的文件/目录执行批量检查与打印。
    skip_print=True：完成 Word/Excel 规范化与保存，不送打印机（供打包下载）。
    raw_print=True：跳过检查与修订处理，直接打印（Word/Excel/PDF）。
    checks_warn_only=None 时与 skip_print 相同：检查未通过只记入提示，仍尽量完成后续处理并保存。
    progress_callback(step, file_index, file_total, file_name, percent) 用于上报当前任务与进度；
    percent 为 0–99 的估算值（整本任务完成前不会为 100%）。
    """
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
    if not files:
        return {"total": 0, "ok": 0, "failed": 0, "details": []}
    try:
        from doc_handlers.word_handler import set_runtime_options as set_word_runtime_options

        set_word_runtime_options(
            word_content_preserve=effective_word_content_preserve,
            word_preserve_page_count=effective_word_preserve_page_count,
            word_image_risk_guard=word_image_risk_guard,
            word_preserve_linked_images=effective_word_preserve_linked_images,
            word_header_footer_layout_fix=effective_word_header_footer_layout_fix,
            word_font_profile=_wfp,
        )
    except Exception:
        pass
    try:
        from doc_handlers.excel_handler import set_runtime_options as set_excel_runtime_options

        set_excel_runtime_options(excel_font_profile=_wfp)
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
    for idx, f in enumerate(files):
        f_start = time.perf_counter()
        doc_type = _get_type(f)
        kwargs = dict(
            progress_callback=progress_callback,
            file_index=idx + 1,
            file_total=total,
        )
        if doc_type == "word":
            success, message, mod_detail = _process_word(
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
                **kwargs,
            )
        elif doc_type == "excel":
            success, message, mod_detail = _process_excel(
                f,
                check_formal=check_formal,
                accept_revisions=accept_revisions,
                printer_name=printer_name,
                copies=copies,
                dry_run=dry_run,
                skip_print=skip_print,
                raw_print=raw_print,
                checks_warn_only=checks_warn_only,
                **kwargs,
            )
        elif doc_type == "pdf":
            success, message, mod_detail = _process_pdf(
                f,
                printer_name=printer_name,
                dry_run=dry_run,
                skip_print=skip_print,
                raw_print=raw_print,
                **kwargs,
            )
        else:
            success, message, mod_detail = False, "未知类型", None
        results.append(
            {
                "path": f,
                "success": success,
                "message": message,
                "modification_detail": mod_detail,
            }
        )
        if success:
            ok_count += 1
        logger.info(
            "run_batch file_done idx=%s/%s type=%s success=%s elapsed=%.2fs file=%s msg=%s",
            idx + 1,
            total,
            doc_type,
            success,
            time.perf_counter() - f_start,
            os.path.basename(f),
            message,
        )

    logger.info(
        "run_batch done total=%s ok=%s failed=%s elapsed=%.2fs",
        len(files),
        ok_count,
        len(files) - ok_count,
        time.perf_counter() - t_start,
    )
    return {
        "total": len(files),
        "ok": ok_count,
        "failed": len(files) - ok_count,
        "details": results,
        "dry_run": dry_run,
        "skip_print": skip_print,
        "raw_print": raw_print,
    }


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
