# -*- coding: utf-8 -*-
"""将签字位版式标误同步到 sign_slot_layout_rules.json 并导出 MD。"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from sign_handlers.detect_correction import _trim_expected_slot_layout, trim_detect_correction
from sign_handlers.detect_correction_rules import infer_pattern_from_source_name

_ROOT = Path(__file__).resolve().parents[1]
_SLOT_JSON_PATH = _ROOT / "sign_handlers" / "sign_slot_layout_rules.json"
_SLOT_MD_PATH = _ROOT / "signature_slot_layout_document_rules_T2.md"


def _esl_to_vocabulary(esl: Dict[str, Any]) -> Dict[str, Any]:
    """把登记字段映射为规则/MD 用的中文描述与表单名。"""
    arr = str(esl.get("arrangement") or "").strip()
    rel = str(esl.get("date_relation") or "").strip()
    pos = str(esl.get("date_position") or "").strip()
    sep = str(esl.get("separator") or "").strip()

    arrangement_labels = {
        "horizontal": "角色从左到右排列",
        "vertical": "角色从上到下排列",
        "mixed": "角色混合排列",
    }
    signature_slot_forms: List[str] = []
    date_layout_forms: List[str] = []
    layout_types: List[str] = []

    if rel == "same_cell":
        date_layout_forms.append("与角色同一单元格/字段标记")
        if sep == "slash":
            signature_slot_forms.append("与角色同一单元格/字段标记")
        if sep in ("space", "adjacent"):
            signature_slot_forms.append("同字段长空格占位")
        layout_types.append("same_cell_inline")
    elif rel == "different_cell":
        if pos == "below":
            date_layout_forms.append("日期在下方相邻/偏右单元格")
            layout_types.append("two_row_signoff_table")
        elif pos == "right":
            date_layout_forms.append("日期在右相邻单元格")
            layout_types.append("adjacent_right_cell")
        if sep in ("cell", "empty_cell"):
            signature_slot_forms.append("右侧空白单元格")
        if sep == "empty_cell":
            date_layout_forms.append("日期在右侧隔列单元格")
    elif rel == "paragraph_inline":
        date_layout_forms.append("与角色同一单元格/字段标记")

    if sep == "slash" and "与角色同一单元格/字段标记" not in date_layout_forms:
        date_layout_forms.append("与角色同一单元格/字段标记")

    return {
        "arrangement": arr,
        "arrangement_label": arrangement_labels.get(arr, arr or "未指定"),
        "date_relation": rel,
        "date_position": pos,
        "separator": sep,
        "signature_slot_forms": signature_slot_forms,
        "date_layout_forms": date_layout_forms,
        "layout_types": layout_types,
    }


def _load_slot_rules_raw() -> Dict[str, Any]:
    if not _SLOT_JSON_PATH.is_file():
        return {"schema_version": 2, "document_layout_rules": []}
    with open(_SLOT_JSON_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        return {"schema_version": 2, "document_layout_rules": []}
    if not isinstance(raw.get("document_layout_rules"), list):
        raw["document_layout_rules"] = []
    return raw


def _save_slot_rules_raw(raw: Dict[str, Any]) -> None:
    raw["updated"] = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    _SLOT_JSON_PATH.write_text(
        json.dumps(raw, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    try:
        from sign_handlers import sign_slot_layout_rules as slr

        slr.load_sign_slot_layout_rules(force=True)
    except Exception:
        pass


def _merge_unique_str_list(base: List[str], extra: List[str]) -> List[str]:
    out = list(base or [])
    seen = {str(x).strip() for x in out if str(x).strip()}
    for item in extra or []:
        s = str(item or "").strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
    return out


def _find_document_layout_index(
    rules: List[Dict[str, Any]], pattern: str
) -> Optional[int]:
    pat = str(pattern or "").strip()
    if not pat:
        return None
    best_i = None
    best_len = -1
    for i, item in enumerate(rules):
        if not isinstance(item, dict):
            continue
        ip = str(item.get("pattern") or "").strip()
        if ip == pat or (ip and ip in pat) or (pat and pat in ip):
            if len(ip) > best_len:
                best_len = len(ip)
                best_i = i
    return best_i


def correction_to_slot_document_rule(
    correction: Dict[str, Any],
    *,
    source_name: str,
    pattern: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    corr = trim_detect_correction(correction)
    wrong = str(corr.get("wrong_description") or "").strip()
    esl = _trim_expected_slot_layout(corr.get("expected_slot_layout"))
    if not esl and not wrong:
        return None
    pat = (pattern or infer_pattern_from_source_name(source_name) or "").strip()
    if len(pat) < 4:
        return None
    vocab = _esl_to_vocabulary(esl) if esl else {}
    note_parts = []
    if wrong:
        note_parts.append(f"人工签字位标误：{wrong[:400]}")
    exp_note = str(corr.get("expected_note") or "").strip()
    if exp_note:
        note_parts.append(exp_note[:200])
    if vocab.get("arrangement_label"):
        note_parts.append(str(vocab["arrangement_label"]))
    entry: Dict[str, Any] = {
        "pattern": pat,
        "match": "contains",
        "note": "；".join(note_parts)[:600],
        "source_example": str(source_name or "")[:300],
        "learned_from_correction": True,
    }
    if esl:
        entry.update(esl)
        entry["arrangement_label"] = vocab.get("arrangement_label", "")
        entry["signature_slot_forms"] = vocab.get("signature_slot_forms") or []
        entry["date_layout_forms"] = vocab.get("date_layout_forms") or []
        entry["layout_types"] = vocab.get("layout_types") or []
    return entry


def upsert_slot_layout_from_correction(
    source_name: str,
    correction: Dict[str, Any],
    *,
    pattern: Optional[str] = None,
) -> Dict[str, Any]:
    """写入 sign_slot_layout_rules.json（含 document_layout_rules）。"""
    entry = correction_to_slot_document_rule(
        correction, source_name=source_name, pattern=pattern
    )
    if not entry:
        return {
            "ok": False,
            "error": "缺少「错在哪」或签字位版式登记，或无法推断 pattern",
        }
    raw = _load_slot_rules_raw()
    doc_rules: List[Dict[str, Any]] = [
        r for r in (raw.get("document_layout_rules") or []) if isinstance(r, dict)
    ]
    pat = str(entry.get("pattern") or "")
    idx = _find_document_layout_index(doc_rules, pat)
    action = "updated"
    if idx is None:
        doc_rules.append(entry)
        action = "created"
    else:
        prev = dict(doc_rules[idx])
        prev.update(entry)
        doc_rules[idx] = prev
    raw["document_layout_rules"] = doc_rules

    vocab = _esl_to_vocabulary(_trim_expected_slot_layout(entry))
    if vocab.get("signature_slot_forms"):
        raw["signature_slot_forms"] = _merge_unique_str_list(
            list(raw.get("signature_slot_forms") or []),
            vocab["signature_slot_forms"],
        )
    if vocab.get("date_layout_forms"):
        raw["date_layout_forms"] = _merge_unique_str_list(
            list(raw.get("date_layout_forms") or []),
            vocab["date_layout_forms"],
        )

    lpr = list(raw.get("layout_priority_rules") or [])
    for lt in vocab.get("layout_types") or []:
        hints = {
            "two_row_signoff_table": ["两行签批表"],
            "same_cell_inline": ["同字段长空格占位"],
            "adjacent_right_cell": ["日期在右相邻单元格", "右侧空白单元格"],
        }.get(lt, [])
        found = False
        for rule in lpr:
            if not isinstance(rule, dict):
                continue
            if str(rule.get("layout_type") or "") == lt:
                rule["match_contains"] = _merge_unique_str_list(
                    list(rule.get("match_contains") or []), hints
                )
                found = True
                break
        if not found and hints:
            lpr.append(
                {
                    "layout_type": lt,
                    "priority": 70,
                    "match_contains": hints,
                    "description": "人工标误同步",
                }
            )
    raw["layout_priority_rules"] = lpr
    _save_slot_rules_raw(raw)
    return {
        "ok": True,
        "action": action,
        "pattern": pat,
        "json_path": str(_SLOT_JSON_PATH),
    }


def export_slot_layout_markdown() -> Dict[str, Any]:
    """从 sign_slot_layout_rules.json 导出签字位文档规则 MD。"""
    try:
        raw = _load_slot_rules_raw()
    except Exception as e:
        return {"ok": False, "error": str(e)[:500]}
    doc_rules = [
        r for r in (raw.get("document_layout_rules") or []) if isinstance(r, dict)
    ]
    lines = [
        "# T2 签字位版式文档规则（人工标误同步）",
        "",
        f"- 配置文件: `sign_handlers/sign_slot_layout_rules.json`",
        f"- 更新: {raw.get('updated', '')}",
        f"- 文档规则条数: {len(doc_rules)}",
        "",
        "| pattern | 角色排列 | 名日关系 | 日期位置 | 分隔 | 说明 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    arr_cn = {
        "horizontal": "左右",
        "vertical": "上下",
        "mixed": "混合",
    }
    rel_cn = {
        "same_cell": "同格",
        "different_cell": "分格",
        "paragraph_inline": "正文",
    }
    pos_cn = {"right": "右方", "below": "下方", "inline": "同行"}
    sep_cn = {
        "slash": "/",
        "space": "空格",
        "cell": "单元格",
        "empty_cell": "空单元格",
        "newline": "换行",
    }
    for item in sorted(doc_rules, key=lambda x: str(x.get("pattern") or "")):
        pat = str(item.get("pattern") or "")
        lines.append(
            "| `{pat}` | {arr} | {rel} | {pos} | {sep} | {note} |".format(
                pat=pat.replace("|", "\\|"),
                arr=arr_cn.get(str(item.get("arrangement") or ""), "—"),
                rel=rel_cn.get(str(item.get("date_relation") or ""), "—"),
                pos=pos_cn.get(str(item.get("date_position") or ""), "—"),
                sep=sep_cn.get(str(item.get("separator") or ""), "—"),
                note=str(item.get("note") or "")[:80].replace("|", "\\|"),
            )
        )
    lines.extend(["", "## 全局签字位表单（JSON 摘录）", ""])
    for key in ("signature_slot_forms", "date_layout_forms"):
        vals = raw.get(key) or []
        if vals:
            lines.append(f"### {key}")
            for v in vals[:30]:
                lines.append(f"- {v}")
            lines.append("")
    _SLOT_MD_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"ok": True, "md_path": str(_SLOT_MD_PATH)}


def sync_slot_rules_from_correction(
    source_name: str,
    correction: Dict[str, Any],
    *,
    export_md: bool = False,
) -> Dict[str, Any]:
    corr = trim_detect_correction(correction)
    esl = _trim_expected_slot_layout(corr.get("expected_slot_layout"))
    wrong = str(corr.get("wrong_description") or "").strip()
    if not esl and not wrong:
        return {"ok": False, "skipped": True, "reason": "无签字位版式登记"}
    up = upsert_slot_layout_from_correction(source_name, correction)
    if not up.get("ok"):
        return up
    out = dict(up)
    if export_md:
        ex = export_slot_layout_markdown()
        out["md_exported"] = bool(ex.get("ok"))
        out["md_path"] = ex.get("md_path") or str(_SLOT_MD_PATH)
        if not ex.get("ok"):
            out["md_warning"] = ex.get("error") or "导出签字位 MD 失败"
    else:
        out["md_exported"] = False
    return out


def match_document_layout_rule(source_name: str) -> Optional[Dict[str, Any]]:
    """按文件名匹配 document_layout_rules（与角色规则同样 contains 逻辑）。"""
    try:
        raw = _load_slot_rules_raw()
    except Exception:
        return None
    name = str(source_name or "").replace("\\", "/")
    base = name.split("/")[-1] if name else ""
    best: Optional[Dict[str, Any]] = None
    best_len = -1
    for item in raw.get("document_layout_rules") or []:
        if not isinstance(item, dict):
            continue
        pat = str(item.get("pattern") or "").strip()
        if not pat or len(pat) < 2:
            continue
        if pat in name or pat in base:
            if len(pat) > best_len:
                best_len = len(pat)
                best = item
    return best
