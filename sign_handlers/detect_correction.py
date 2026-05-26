# -*- coding: utf-8 -*-
"""人工识别纠正登记：重新识别时带入期望角色与标签提示。"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from sign_handlers.config import ROLE_ID_TO_KEYWORD

_CORRECTION_KEYS = (
    "wrong_description",
    "expected_roles",
    "expected_note",
    "wrong_roles_detected",
    "label_keywords",
    "reference_images",
    "expected_slot_layout",
    "correction_save",
    "updated_at",
)

_SLOT_LAYOUT_ALLOWED = {
    "arrangement": frozenset(
        {"horizontal", "vertical", "mixed", "single", "inline", "unknown"}
    ),
    "date_relation": frozenset(
        {"same_cell", "different_cell", "paragraph_inline", "none"}
    ),
    "date_position": frozenset({"right", "below", "inline", "left", "none"}),
    "separator": frozenset(
        {
            "slash",
            "space",
            "cell",
            "empty_cell",
            "newline",
            "adjacent",
            "none",
            "unknown",
        }
    ),
}


def _trim_expected_slot_layout(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    out: Dict[str, Any] = {}
    for k, allowed in _SLOT_LAYOUT_ALLOWED.items():
        v = str(data.get(k) or "").strip()
        if v and v in allowed:
            out[k] = v
    if "same_cell" in data:
        if data.get("same_cell") is True:
            out["same_cell"] = True
        elif data.get("same_cell") is False:
            out["same_cell"] = False
    if out.get("date_relation") == "same_cell":
        out["same_cell"] = True
    elif out.get("date_relation") == "different_cell":
        out["same_cell"] = False
    return out


def trim_detect_correction(data: Any) -> Dict[str, Any]:
    if not isinstance(data, dict):
        return {}
    out: Dict[str, Any] = {}
    for k in _CORRECTION_KEYS:
        if k not in data:
            continue
        v = data[k]
        if k == "expected_roles" and isinstance(v, list):
            roles = []
            for rid in v:
                rid_s = str(rid or "").strip()
                if rid_s in ROLE_ID_TO_KEYWORD and rid_s not in roles:
                    roles.append(rid_s)
            out[k] = roles
        elif k == "wrong_roles_detected" and isinstance(v, list):
            out[k] = [
                str(x).strip()
                for x in v
                if str(x).strip() in ROLE_ID_TO_KEYWORD
            ]
        elif k == "label_keywords" and isinstance(v, list):
            kws = []
            seen = set()
            for item in v:
                s = str(item or "").strip()
                if not s or s in seen:
                    continue
                seen.add(s)
                kws.append(s[:64])
                if len(kws) >= 24:
                    break
            out[k] = kws
        elif k == "reference_images" and isinstance(v, list):
            imgs = []
            for item in v[:6]:
                if not isinstance(item, dict):
                    continue
                iid = str(item.get("id") or "").strip()
                if not iid:
                    continue
                ftp_p = str(item.get("ftp_path") or "").strip()
                entry = {
                    "id": iid[:64],
                    "filename": str(item.get("filename") or "")[:200],
                    "uploaded_at": str(item.get("uploaded_at") or "")[:40],
                }
                if ftp_p:
                    entry["ftp_path"] = ftp_p[:768]
                imgs.append(entry)
            out[k] = imgs
        elif k == "expected_slot_layout" and isinstance(v, dict):
            trimmed = _trim_expected_slot_layout(v)
            if trimmed:
                out[k] = trimmed
        elif k == "correction_save" and isinstance(v, dict):
            out[k] = {
                "roles": bool(v.get("roles")),
                "slot": bool(v.get("slot")),
            }
        elif k in ("wrong_description", "expected_note", "updated_at"):
            out[k] = str(v or "")[:2000]
        else:
            out[k] = v
    return out


def _signature_layout_from_correction(
    correction: Dict[str, Any], role_ids: List[str]
) -> Optional[Dict[str, Any]]:
    esl = _trim_expected_slot_layout(correction.get("expected_slot_layout"))
    if not esl:
        return None
    roles = [r for r in (role_ids or []) if r in ROLE_ID_TO_KEYWORD]
    if not roles:
        return None
    rel = esl.get("date_relation") or "none"
    pos = esl.get("date_position") or "none"
    sep = esl.get("separator") or "none"
    has_date = rel != "none"
    role_layouts: Dict[str, Any] = {}
    for rid in roles:
        role_layouts[rid] = {
            "name_slot": True,
            "date_slot": has_date,
            "date_relation": rel,
            "date_position": pos,
            "separator": sep,
            "name_loc": "detect_correction",
            "date_loc": "detect_correction" if has_date else None,
        }
    return {
        "ok": True,
        "kind": "detect_correction",
        "arrangement": esl.get("arrangement") or "unknown",
        "role_layouts": role_layouts,
        "source": "manual_correction",
    }


def _filter_blocks_by_roles(blocks: Any, allowed_roles: List[str]) -> List[Dict[str, Any]]:
    if not isinstance(blocks, list):
        return []
    allow = set(allowed_roles or [])
    if not allow:
        return []
    out: List[Dict[str, Any]] = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        fields = b.get("fields")
        if not isinstance(fields, list):
            continue
        kept_fields = []
        has_role = False
        for f in fields:
            if not isinstance(f, dict):
                continue
            if str(f.get("type") or "").strip() == "role_id":
                rid = str(f.get("name") or "").strip()
                if rid not in allow:
                    continue
                has_role = True
            kept_fields.append(f)
        if not has_role:
            continue
        b2 = dict(b)
        b2["fields"] = kept_fields
        out.append(b2)
    return out


def apply_detect_correction(
    result: Dict[str, Any],
    correction: Optional[Dict[str, Any]],
    *,
    source_name: str = "",
) -> Dict[str, Any]:
    """
    将人工登记纠正作为“识别提示”应用到结果（在文件名规则之后调用）。
    重要：不强行覆盖文档实识别结果，避免出现“看起来改对了，但实际不可签”的假阳性。
    """
    if not isinstance(result, dict):
        return result
    corr = trim_detect_correction(correction)
    if not corr:
        return result
    result = dict(result)
    result["detect_correction"] = corr

    expected = [r for r in (corr.get("expected_roles") or []) if r in ROLE_ID_TO_KEYWORD]
    label_kws = corr.get("label_keywords") or []
    wrong_desc = str(corr.get("wrong_description") or "").strip()
    expected_note = str(corr.get("expected_note") or "").strip()

    if expected:
        # 仅补充提示证据，不覆盖 roles/blocks，保证“展示=文档真实识别”。
        role_evidence = result.get("role_evidence")
        if not isinstance(role_evidence, dict):
            role_evidence = {}
        for rid in expected:
            arr = role_evidence.get(rid) if isinstance(role_evidence.get(rid), list) else []
            if not arr:
                preview = expected_note or wrong_desc or source_name
                arr = [
                    {
                        "confidence": 0.92,
                        "source_hint": "detect_correction_hint",
                        "matched_rules": ["manual_correction_hint"],
                        "label_preview": str(preview or rid)[:120],
                    }
                ]
            if label_kws:
                for kw in label_kws[:8]:
                    arr.append(
                        {
                            "confidence": 0.9,
                            "source_hint": "detect_correction_keyword",
                            "matched_rules": ["manual_label_keyword_hint"],
                            "label_preview": str(kw)[:120],
                        }
                    )
            role_evidence[rid] = arr[:8]
        result["role_evidence"] = role_evidence
        ds = dict(result.get("debug_summary") or {})
        ds["correction_hint_only"] = True
        ds["correction_expected_roles"] = expected
        if wrong_desc:
            ds["correction_wrong_description"] = wrong_desc[:200]
        result["debug_summary"] = ds
    elif label_kws:
        role_evidence = result.get("role_evidence")
        if isinstance(role_evidence, dict):
            for rid in list(role_evidence.keys()):
                if rid not in ROLE_ID_TO_KEYWORD:
                    continue
                arr = role_evidence.get(rid)
                if not isinstance(arr, list):
                    arr = []
                for kw in label_kws[:6]:
                    arr.append(
                        {
                            "confidence": 0.92,
                            "source_hint": "detect_correction_keyword",
                            "matched_rules": ["manual_label_keyword"],
                            "label_preview": str(kw)[:120],
                        }
                    )
                role_evidence[rid] = arr[:6]
            result["role_evidence"] = role_evidence
        ds = dict(result.get("debug_summary") or {})
        ds["correction_keywords_added"] = label_kws
        result["debug_summary"] = ds

    # 保留标误版式作为 hint，不覆盖 signature_layout（由真实文档分析产生）。
    esl = _trim_expected_slot_layout(corr.get("expected_slot_layout"))
    if esl:
        ds = dict(result.get("debug_summary") or {})
        ds["correction_slot_layout_hint"] = esl
        result["debug_summary"] = ds

    return result
