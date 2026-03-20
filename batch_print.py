# -*- coding: utf-8 -*-
"""
批量打印工具：在打印前检查封面签字（Word）、修订记录，并自动接受所有修订后打印。
支持：.doc/.docx、.xls/.xlsx、.pdf
"""
from __future__ import annotations

import argparse
import os
import sys

# 支持的扩展名（小写）
WORD_EXT = {".doc", ".docx", ".docm"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm"}
PDF_EXT = {".pdf"}

# COM「发生意外」错误码，WPS/Office/Excel 易触发
_COM_ERR_CODES = (-2147352567, -2147467259, -2146827864)
_COM_ERR_MSG = "文档处理时发生意外，请确认 WPS/Excel 已安装、文档未被占用，或稍后重试。"


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


def _process_word(path, check_formal, check_signature, accept_revisions, printer_name, copies, dry_run, progress_callback=None, file_index=0, file_total=1, auto_fix_formal=True):
    from doc_handlers.word_handler import (
        check_formal_document,
        check_cover_signature,
        accept_all_revisions_and_save,
        auto_fix_formal_word,
        print_word_document,
    )
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
    if check_formal:
        report("正式性检查（标黄/修订/批注/非黑字体/隐藏文字/水印等）")
        try:
            passed, issues = check_formal_document(path)
        except Exception as e:
            return False, _format_error(e), []
        if not passed:
            if dry_run:
                return False, "正式性检查未通过：" + "；".join(issues), issues
            report("自动修复正式性问题")
            try:
                auto_fix_formal_word(path)
            except Exception as e:
                return False, _format_error(e), issues
    if check_signature:
        report("封面签字与日期检查")
        try:
            ok, msg = check_cover_signature(path)
        except Exception as e:
            return False, _format_error(e), errors
        if not ok:
            return False, f"封面签字未通过: {msg}", errors
    if not dry_run:
        report("接受修订与规范化")
        try:
            accept_all_revisions_and_save(path)
        except Exception as e:
            return False, _format_error(e), errors
    if dry_run:
        report("预检完成")
        return True, "校验通过（模拟，未打印）", errors
    report("提交打印")
    try:
        print_word_document(path, printer_name=printer_name, copies=copies)
        return True, "已提交打印", errors
    except Exception as e:
        return False, _format_error(e), errors


def _process_excel(path, check_formal, accept_revisions, printer_name, copies, dry_run, progress_callback=None, file_index=0, file_total=1, auto_fix_formal=True):
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
    if check_formal:
        report("正式性检查（批注/非黑字体；保留表格底色）")
        passed, issues = check_formal_document(path, check_highlight=False)
        if not passed:
            if dry_run:
                return False, "正式性检查未通过：" + "；".join(issues), issues
            report("自动修复正式性问题")
            try:
                auto_fix_formal_excel(path)
            except Exception as e:
                return False, "自动修复失败 " + _format_error(e), issues
    if not dry_run:
        report("接受修订")
        try:
            accept_all_changes_and_save(path, accept_revisions=accept_revisions, remove_highlights=False)
        except Exception:
            pass
        report("文档规范化")
        try:
            ensure_font_black(path)
        except Exception as e:
            return False, "字体规范化失败 " + _format_error(e), []
        report("风险矩阵外：清除底色、自适应行高（优先内容显示）")
        try:
            excel_normalize_matrix_and_layout(path)
        except Exception as e:
            return False, "矩阵外排版失败 " + _format_error(e), []
    if dry_run:
        report("预检完成")
        return True, "校验通过（模拟，未打印）", []
    report("提交打印")
    try:
        print_excel_workbook(path, printer_name=printer_name, copies=copies)
        return True, "已提交打印", []
    except Exception as e:
        return False, "打印失败 " + _format_error(e), []


def _process_pdf(path, printer_name, dry_run, progress_callback=None, file_index=0, file_total=1):
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
    if dry_run:
        report("预检完成")
        return True, "校验通过（模拟，未打印）", []
    report("提交打印")
    try:
        print_pdf(path, printer_name=printer_name)
        return True, "已提交打印", []
    except Exception as e:
        return False, str(e), []


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
    progress_callback=None,
):
    """
    对 paths 中的文件/目录执行批量检查与打印。
    progress_callback(step, file_index, file_total, file_name, percent) 用于上报当前任务与进度；
    percent 为 0–99 的估算值（整本任务完成前不会为 100%）。
    """
    files = _collect_files(paths, recursive=recursive)
    if not files:
        return {"total": 0, "ok": 0, "failed": 0, "details": []}

    results = []
    ok_count = 0
    total = len(files)
    for idx, f in enumerate(files):
        doc_type = _get_type(f)
        kwargs = dict(
            progress_callback=progress_callback,
            file_index=idx + 1,
            file_total=total,
        )
        if doc_type == "word":
            success, message, _ = _process_word(
                f,
                check_formal=check_formal,
                check_signature=check_signature,
                accept_revisions=accept_revisions,
                printer_name=printer_name,
                copies=copies,
                dry_run=dry_run,
                **kwargs,
            )
        elif doc_type == "excel":
            success, message, _ = _process_excel(
                f,
                check_formal=check_formal,
                accept_revisions=accept_revisions,
                printer_name=printer_name,
                copies=copies,
                dry_run=dry_run,
                **kwargs,
            )
        elif doc_type == "pdf":
            success, message, _ = _process_pdf(
                f,
                printer_name=printer_name,
                dry_run=dry_run,
                **kwargs,
            )
        else:
            success, message = False, "未知类型"
        results.append({"path": f, "success": success, "message": message})
        if success:
            ok_count += 1

    return {
        "total": len(files),
        "ok": ok_count,
        "failed": len(files) - ok_count,
        "details": results,
        "dry_run": dry_run,
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
