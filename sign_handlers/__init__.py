# -*- coding: utf-8 -*-
"""在线签名：将手写签名/日期 PNG 写入 Word/Excel 模板预留空白位。"""
from __future__ import annotations

import os

from sign_handlers.config import ROLE_ID_TO_KEYWORD
from sign_handlers.sign_docx import sign_docx
from sign_handlers.sign_xlsx import sign_xlsx

__all__ = ["sign_document", "ROLE_ID_TO_KEYWORD"]


def sign_document(
    file_path: str,
    role_to_signature_png: dict,
    role_to_date_png: dict,
    out_path: str | None = None,
) -> str:
    """
    根据扩展名写入签名与日期图片。
    role_* 的 key 为 ROLE_ID_TO_KEYWORD 中的 id（如 author）。
    返回输出文件路径。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        return sign_docx(file_path, role_to_signature_png, role_to_date_png, out_path)
    if ext == ".xlsx":
        return sign_xlsx(file_path, role_to_signature_png, role_to_date_png, out_path)
    raise ValueError(f"不支持的格式: {ext}")
