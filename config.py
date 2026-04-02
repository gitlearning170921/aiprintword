# -*- coding: utf-8 -*-
"""
选择文档处理后端：WPS 或 Microsoft Office。
默认使用 WPS；USE_OFFICE=1（或页面配置）可改回 Office。
运行时以 get_word_progid() / use_wps_runtime() 为准，便于读取库内配置。
"""
import os


def use_wps_runtime() -> bool:
    try:
        from runtime_settings.resolve import get_setting

        v = get_setting("USE_OFFICE")
        return str(v).strip().lower() not in ("1", "true", "yes", "on")
    except Exception:
        return os.environ.get("USE_OFFICE", "").strip().lower() not in (
            "1",
            "true",
            "yes",
            "on",
        )


def get_word_progid() -> str:
    return "KWPS.Application" if use_wps_runtime() else "Word.Application"


def get_excel_progid() -> str:
    return "KET.Application" if use_wps_runtime() else "Excel.Application"


USE_WPS = use_wps_runtime()
WORD_PROGID = get_word_progid()
EXCEL_PROGID = get_excel_progid()
