# -*- coding: utf-8 -*-
"""
Word 修订（跟踪更改）下的「有效正文」文本提取。

识别签字角色 / 签字落位应基于**修订接受后**读者所见内容：
- 忽略 w:del / w:moveFrom 内文字
- 保留 w:ins / w:moveTo 及普通 w:t
"""
from __future__ import annotations

import os
import zipfile
from typing import Any, Optional

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_SKIP_LOCAL = frozenset({"del", "moveFrom"})


def _local_tag(tag: str) -> str:
    if not tag:
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


def ooxml_effective_text(element: Any) -> str:
    """从 OOXML 节点提取修订接受后的可见文本。"""
    if element is None:
        return ""
    parts: list[str] = []

    def walk(node: Any, in_deleted: bool) -> None:
        local = _local_tag(getattr(node, "tag", "") or "")
        deleted = in_deleted or local in _SKIP_LOCAL
        if local == "t" and not deleted:
            parts.append(getattr(node, "text", None) or "")
        elif local == "tab" and not deleted:
            parts.append("\t")
        elif local in ("br", "cr") and not deleted:
            parts.append("\n")
        for ch in list(node):
            walk(ch, deleted)

    walk(element, False)
    return "".join(parts)


def paragraph_effective_text(paragraph: Any) -> str:
    try:
        el = paragraph._element
    except Exception:
        return str(getattr(paragraph, "text", None) or "")
    return ooxml_effective_text(el)


def cell_effective_text(cell: Any) -> str:
    try:
        paras = list(getattr(cell, "paragraphs", None) or [])
    except Exception:
        paras = []
    if paras:
        chunks = [paragraph_effective_text(p) for p in paras]
        return "\n".join(t for t in chunks if t)
    try:
        return ooxml_effective_text(cell._tc)
    except Exception:
        return str(getattr(cell, "text", None) or "")


def docx_has_track_changes(path: str) -> bool:
    """快速检测 docx 是否含跟踪修订标记（ins/del）。"""
    path = os.path.abspath(path or "")
    if not os.path.isfile(path):
        return False
    markers = (b"<w:ins ", b"<w:del ", b"<w:moveFrom ", b"<w:moveTo ")
    try:
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if not name.startswith("word/") or not name.endswith(".xml"):
                    continue
                if "document" not in name and "header" not in name and "footer" not in name:
                    continue
                try:
                    blob = zf.read(name)
                except Exception:
                    continue
                if any(m in blob for m in markers):
                    return True
    except Exception:
        return False
    return False
