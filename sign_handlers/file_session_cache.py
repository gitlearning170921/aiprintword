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
    "signature_layout",
    "slot_probe",
)


def _layout_has_usable_slots(layout: Any) -> bool:
    if not isinstance(layout, dict):
        return False
    rl = layout.get("role_layouts")
    if not isinstance(rl, dict):
        return False
    for row in rl.values():
        if isinstance(row, dict) and (row.get("name_slot") or row.get("date_slot")):
            return True
    return False


def _probe_has_usable_results(probe: Any) -> bool:
    if not isinstance(probe, dict) or not probe.get("ok"):
        return False
    pr = probe.get("per_role_results")
    if not isinstance(pr, dict) or not pr:
        return False
    for row in pr.values():
        if isinstance(row, dict) and row.get("placed"):
            return True
    return False


def _trim_blocks_for_lite(blocks: Any) -> list:
    """lite 响应：仅保留签字位推断所需字段，体积远小于完整 blocks。"""
    if not isinstance(blocks, list):
        return []
    out: list = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        fields = []
        for f in b.get("fields") or []:
            if not isinstance(f, dict):
                continue
            ft = str(f.get("type") or "").strip().lower()
            if ft not in ("role_id", "date"):
                continue
            fields.append({"type": ft, "name": f.get("name")})
        row: Dict[str, Any] = {}
        if fields:
            row["fields"] = fields
        for k in ("source_hint", "label_preview", "table_hint"):
            if b.get(k) is not None:
                row[k] = b[k]
        if row:
            out.append(row)
    return out


def trim_detect_snapshot(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    out: Dict[str, Any] = {}
    for k in _DETECT_KEYS:
        if k in data:
            out[k] = data[k]
    return out


def trim_detect_snapshot_lite(data: Any) -> Dict[str, Any]:
    """首屏恢复用：保留工作台三列所需字段，去掉 blocks/role_evidence 等大字段。"""
    if not isinstance(data, dict):
        return {}
    out: Dict[str, Any] = {}
    for k in (
        "ok",
        "file_id",
        "source_name",
        "roles",
        "error",
        "error_code",
        "content_sha256",
        "document_role_rule",
        "slot_probe",
        "signature_layout",
        "debug_summary",
    ):
        if k in data:
            out[k] = data[k]
    layout = out.get("signature_layout")
    if isinstance(layout, dict) and layout.get("role_layouts"):
        slim = dict(layout)
        rl = slim.get("role_layouts")
        if isinstance(rl, dict):
            slim["role_layouts"] = {
                rid: row for rid, row in rl.items() if isinstance(row, dict)
            }
        out["signature_layout"] = slim
    # 三列判定常依赖 blocks 作为补充证据（尤其日期位）；lite 保留极简 blocks，避免误报缺位。
    slim_blocks = _trim_blocks_for_lite(data.get("blocks"))
    if slim_blocks:
        out["blocks"] = slim_blocks
    re = data.get("role_evidence")
    if isinstance(re, dict) and re:
        out["role_evidence"] = re
    return out


def trim_detect_correction(data: Any) -> Dict[str, Any]:
    from sign_handlers.detect_correction import trim_detect_correction as _trim

    return _trim(data)


def trim_workbench_state(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    keys = (
        "status",
        "rolesLabel",
        "slotLabel",
        "slotRolesLine",
        "slotLayoutLine",
        "slotBadgeClass",
        "slotExplain",
        "slotTags",
        "detectedRoleIds",
        "slotProbeOk",
        "slotSignable",
        "slotMissingName",
        "slotMissingDate",
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
