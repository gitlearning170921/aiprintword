# -*- coding: utf-8 -*-
"""
选择文档处理后端：WPS 或 Microsoft Office。
默认使用 WPS（避免 Office RPC 不可用）；设置环境变量 USE_OFFICE=1 可改回 Office。
"""
import os

# True = 使用 WPS 文字/表格（KWPS.Application, KET.Application）
# False = 使用 Microsoft Word/Excel
USE_WPS = os.environ.get("USE_OFFICE", "").strip().lower() not in ("1", "true", "yes")

# WPS 文字 / Microsoft Word 的 COM ProgID
WORD_PROGID = "KWPS.Application" if USE_WPS else "Word.Application"
# WPS 表格 / Microsoft Excel 的 COM ProgID
EXCEL_PROGID = "KET.Application" if USE_WPS else "Excel.Application"
