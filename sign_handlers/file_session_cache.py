# -*- coding: utf-8 -*-
"""待签文件识别结果与工作台行状态持久化（与 file_id 关联）。"""
from __future__ import annotations

import json
from typing import Any, Dict, Optional

_DETECT_KEYS = (
    "ok",
    "file_id",
    "source_name",
    "roles",
    "blocks",
    "filename_rules",
    "role_evidence",
    "light",
    "error",
    "error_code",
    "content_sha256",
    "debug_summary",
    "document_role_rule",
)


def trim_detect_snapshot(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    out: Dict[str, Any] = {}
    for k in _DETECT_KEYS:
        if k in data:
            out[k] = data[k]
    return out


def trim_workbench_state(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    keys = (
        "status",
        "rolesLabel",
        "detectExplain",
        "detectWrongNote",
        "manualDetectWrong",
        "editor",
        "reviewer",
        "approver",
        "doc_date",
        "locale",
        "country",
        "selected",
    )
    return {k: data[k] for k in keys if k in data}


def _json_load(raw: Any) -> Any:
    if raw is None:
        return None
    if isinstance(raw, (dict, list)):
        return raw
    if isinstance(raw, (bytes, bytearray)):
        raw = raw.decode("utf-8", errors="replace")
    if not isinstance(raw, str) or not raw.strip():
        return None
    try:
        return json.loads(raw)
    except Exception:
        return None
