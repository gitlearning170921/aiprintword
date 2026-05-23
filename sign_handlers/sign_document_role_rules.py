# -*- coding: utf-8 -*-
"""按文档名补全签字角色的可维护规则。"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

from sign_handlers.config import ROLE_ID_TO_KEYWORD

_JSON_NAME = "sign_document_role_rules.json"
_RULES_CACHE: Optional[Dict[str, Any]] = None


def _json_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), _JSON_NAME)


def _norm_name(s: str) -> str:
    return (
        str(s or "")
        .replace("\\", "/")
        .replace("（", "(")
        .replace("）", ")")
        .strip()
        .lower()
    )


def load_sign_document_role_rules(force: bool = False) -> Dict[str, Any]:
    global _RULES_CACHE
    if _RULES_CACHE is not None and not force:
        return _RULES_CACHE
    path = _json_path()
    if not os.path.isfile(path):
        _RULES_CACHE = {"schema_version": 1, "rules": []}
        return _RULES_CACHE
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        raw = {}
    rules = raw.get("rules")
    if not isinstance(rules, list):
        rules = []
    clean = []
    for item in rules:
        if not isinstance(item, dict):
            continue
        pattern = str(item.get("pattern") or item.get("name") or "").strip()
        if not pattern:
            continue
        match = str(item.get("match") or "endswith").strip().lower()
        if match not in {"exact", "endswith", "contains"}:
            match = "endswith"
        roles = []
        for rid in item.get("roles") or []:
            rid_s = str(rid or "").strip()
            if rid_s in ROLE_ID_TO_KEYWORD and rid_s not in roles:
                roles.append(rid_s)
        entry: Dict[str, Any] = {
            "pattern": pattern.replace("\\", "/"),
            "match": match,
            "roles": roles,
            "note": str(item.get("note") or ""),
        }
        sp = str(item.get("sign_policy") or "").strip()
        if sp in {"no_sign", "detect_roles"}:
            entry["sign_policy"] = sp
        if not roles:
            entry["no_sign_required"] = True
            entry.setdefault("sign_policy", "no_sign")
        cat = str(item.get("category") or "").strip()
        if cat:
            entry["category"] = cat
        lbl = str(item.get("label") or "").strip()
        if lbl:
            entry["label"] = lbl
        clean.append(entry)
    _RULES_CACHE = {
        "schema_version": int(raw.get("schema_version", 1) or 1),
        "source": raw.get("source") or "",
        "rules": clean,
    }
    return _RULES_CACHE


def match_document_role_rule(source_name: str) -> Optional[Dict[str, Any]]:
    name = _norm_name(source_name)
    if not name:
        return None
    best = None
    best_len = -1
    for rule in load_sign_document_role_rules().get("rules", []):
        pattern = _norm_name(rule.get("pattern") or "")
        if not pattern:
            continue
        mode = rule.get("match") or "endswith"
        hit = False
        if mode == "exact":
            hit = name == pattern
        elif mode == "contains":
            hit = pattern in name
        else:
            hit = name.endswith(pattern)
        if hit and len(pattern) > best_len:
            best = rule
            best_len = len(pattern)
    return best


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
            ftype = str(f.get("type") or "").strip()
            if ftype == "role_id":
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


def apply_document_role_rules(result: Dict[str, Any], source_name: str) -> Dict[str, Any]:
    """按文件名规则修正文档角色；人工维护规则命中后以规则为准。"""
    if not isinstance(result, dict):
        return result
    rule = match_document_role_rule(source_name)
    if not rule:
        return result
    roles_from_rule = [r for r in (rule.get("roles") or []) if r in ROLE_ID_TO_KEYWORD]
    result["document_role_rule"] = {
        "matched": True,
        "pattern": rule.get("pattern"),
        "roles": roles_from_rule,
        "override": True,
        "no_sign_required": not bool(roles_from_rule),
        "sign_policy": rule.get("sign_policy")
        or ("no_sign" if not roles_from_rule else "detect_roles"),
        "category": rule.get("category") or "",
        "label": rule.get("label") or "",
    }
    if not roles_from_rule:
        # 空角色也是明确识别结果：用例表、规范评审报告等不应进入签字流程的文档。
        result["ok"] = True
        result.pop("error", None)
        result.pop("error_code", None)
        result["roles"] = []
        result["blocks"] = []
        result["role_evidence"] = {}
        result["debug_summary"] = {
            "rule_override": True,
            "override_pattern": rule.get("pattern"),
            "override_roles": [],
        }
        return result
    original_roles = [str((x or {}).get("id") or "") for x in (result.get("roles") or []) if isinstance(x, dict)]
    result["roles"] = [
        {"id": rid, "confidence": 0.99, "source": "document_role_rule"}
        for rid in roles_from_rule
    ]
    result["blocks"] = _filter_blocks_by_roles(result.get("blocks"), roles_from_rule)
    role_evidence = result.get("role_evidence")
    if isinstance(role_evidence, dict):
        result["role_evidence"] = {
            rid: role_evidence.get(rid, [])
            for rid in roles_from_rule
            if isinstance(role_evidence.get(rid), list)
        }
    else:
        result["role_evidence"] = {}
    for rid in roles_from_rule:
        ev = result["role_evidence"].setdefault(rid, [])
        if not ev:
            ev.append(
                {
                    "confidence": 0.99,
                    "source_hint": "document_name_rule",
                    "matched_rules": ["document_role_rule_override"],
                    "label_preview": str(rule.get("pattern") or ""),
                }
            )
    result["debug_summary"] = {
        "rule_override": True,
        "override_pattern": rule.get("pattern"),
        "override_roles": roles_from_rule,
        "dropped_roles": [rid for rid in original_roles if rid and rid not in set(roles_from_rule)],
    }
    return result

