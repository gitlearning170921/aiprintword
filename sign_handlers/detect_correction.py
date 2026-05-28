# -*- coding: utf-8 -*-
"""人工识别纠正登记：重新识别时带入期望角色与标签提示。"""
from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

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


def _get_setting_bool(key: str, default: bool) -> bool:
    try:
        from runtime_settings.resolve import get_setting

        return bool(get_setting(key))
    except Exception:
        return bool(default)


def _get_setting_int(key: str, default: int, low: int, high: int) -> int:
    try:
        from runtime_settings.resolve import get_setting

        v = int(get_setting(key))
    except Exception:
        v = int(default)
    if v < low:
        v = low
    if v > high:
        v = high
    return v


def _ocr_text_from_image_bytes(data: bytes) -> str:
    """轻量 OCR：优先 pytesseract（可选依赖），失败时返回空串。"""
    if not data:
        return ""
    try:
        from PIL import Image
    except Exception:
        return ""
    try:
        import pytesseract
    except Exception:
        return ""
    try:
        im = Image.open(io.BytesIO(data)).convert("L")
    except Exception:
        return ""
    txt = ""
    for lang in ("chi_sim+eng", "eng"):
        try:
            t = pytesseract.image_to_string(im, lang=lang)
            if t and len(t.strip()) >= len(txt.strip()):
                txt = t
        except Exception:
            continue
    return str(txt or "").strip()


def _extract_role_keywords_from_ocr_text(text: str) -> List[str]:
    t = str(text or "")
    if not t:
        return []
    t_norm = re.sub(r"\s+", "", t).lower()
    out: List[str] = []
    seen = set()
    try:
        from sign_handlers.config import role_keywords

        for rid in ROLE_ID_TO_KEYWORD:
            for kw in role_keywords(rid):
                ks = str(kw or "").strip()
                if not ks:
                    continue
                if re.sub(r"\s+", "", ks).lower() in t_norm:
                    if ks not in seen:
                        seen.add(ks)
                        out.append(ks[:64])
                    break
    except Exception:
        for rid, kw in (ROLE_ID_TO_KEYWORD or {}).items():
            ks = str(kw or "").strip()
            if not ks:
                continue
            if re.sub(r"\s+", "", ks).lower() in t_norm and ks not in seen:
                seen.add(ks)
                out.append(ks[:64])
    return out[:24]


def _ocr_keywords_from_reference_images(
    corr: Dict[str, Any],
) -> Tuple[List[str], Dict[str, Any]]:
    """从参考图 OCR 提取角色词；返回 (关键词列表, 统计信息供前端展示)。"""
    stats: Dict[str, Any] = {
        "enabled": _get_setting_bool("SIGN_DETECT_HINT_OCR_REF_IMAGES", True),
        "images_configured": 0,
        "images_tried": 0,
        "images_read_ok": 0,
        "keywords_found": [],
        "skipped_reason": "",
    }
    if not stats["enabled"]:
        stats["skipped_reason"] = "ocr_disabled"
        return [], stats
    imgs = list(corr.get("reference_images") or [])
    stats["images_configured"] = len(imgs)
    if not imgs:
        stats["skipped_reason"] = "no_reference_images"
        return [], stats
    max_imgs = _get_setting_int("SIGN_DETECT_HINT_OCR_MAX_IMAGES", 2, 0, 6)
    if max_imgs <= 0:
        stats["skipped_reason"] = "ocr_max_images_zero"
        return [], stats
    merged: List[str] = []
    seen = set()
    try:
        from sign_handlers.detect_correction_storage import download_reference_image
    except Exception:
        stats["skipped_reason"] = "storage_unavailable"
        return [], stats
    for meta in imgs[:max_imgs]:
        if not isinstance(meta, dict):
            continue
        stats["images_tried"] += 1
        try:
            blob, _ = download_reference_image(meta)
        except Exception:
            continue
        stats["images_read_ok"] += 1
        txt = _ocr_text_from_image_bytes(blob)
        kws = _extract_role_keywords_from_ocr_text(txt)
        for kw in kws:
            if kw in seen:
                continue
            seen.add(kw)
            merged.append(kw)
            if len(merged) >= 24:
                stats["keywords_found"] = merged[:24]
                return merged, stats
    stats["keywords_found"] = merged
    if stats["images_read_ok"] and not merged:
        stats["skipped_reason"] = "ocr_no_role_keywords"
    elif not stats["images_read_ok"]:
        stats["skipped_reason"] = "ftp_read_failed"
    return merged, stats


def build_detect_hint(correction: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """把人工纠正整理为识别 hint（供 detect_file 软引导）。"""
    corr = trim_detect_correction(correction)
    if not corr:
        return {}
    manual_kws: List[str] = []
    seen = set()
    for kw in list(corr.get("label_keywords") or []):
        s = str(kw or "").strip()
        if not s or s in seen:
            continue
        seen.add(s)
        manual_kws.append(s[:64])
    ocr_kws, ocr_stats = _ocr_keywords_from_reference_images(corr)
    kws = list(manual_kws)
    for kw in ocr_kws:
        s = str(kw or "").strip()
        if not s or s in seen:
            continue
        seen.add(s)
        kws.append(s[:64])
        if len(kws) >= 24:
            break
    return {
        "expected_roles": [r for r in (corr.get("expected_roles") or []) if r in ROLE_ID_TO_KEYWORD],
        "label_keywords": kws,
        "manual_keywords": manual_kws,
        "ocr_keywords": ocr_kws,
        "ocr_hint_stats": ocr_stats,
        "wrong_description": str(corr.get("wrong_description") or "").strip(),
        "expected_note": str(corr.get("expected_note") or "").strip(),
        "expected_slot_layout": _trim_expected_slot_layout(corr.get("expected_slot_layout")),
    }


def apply_detect_correction(
    result: Dict[str, Any],
    correction: Optional[Dict[str, Any]],
    *,
    source_name: str = "",
) -> Dict[str, Any]:
    """将人工登记作为“提示证据”挂载到识别结果（不直接覆盖识别输出）。"""
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

    # 版式纠正仅挂在 debug_summary 作为 hint，不直接替换 signature_layout。
    esl = _trim_expected_slot_layout(corr.get("expected_slot_layout"))
    if esl:
        ds = dict(result.get("debug_summary") or {})
        ds["correction_slot_layout_hint"] = esl
        result["debug_summary"] = ds

    return result
