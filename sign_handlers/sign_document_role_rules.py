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
        clean.append(
            {
                "pattern": pattern.replace("\\", "/"),
                "match": match,
                "roles": roles,
                "note": str(item.get("note") or ""),
            }
        )
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


def apply_document_role_rules(result: Dict[str, Any], source_name: str) -> Dict[str, Any]:
    """把文件名规则识别到的角色并入 detect 结果，正文识别不够时兜底补全。"""
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
    }
    if not roles_from_rule:
        # 空角色也是明确识别结果：用于规范评审报告、调查问卷等不应进入签字流程的文档。
        result["roles"] = []
        result["blocks"] = []
        return result
    existing = []
    for r in result.get("roles") or []:
        rid = str((r or {}).get("id") or "")
        if rid:
            existing.append(rid)
    roles = list(result.get("roles") or [])
    for rid in roles_from_rule:
        if rid not in existing:
            roles.append({"id": rid, "confidence": 0.99, "source": "document_role_rule"})
            existing.append(rid)
    result["roles"] = roles
    return result

