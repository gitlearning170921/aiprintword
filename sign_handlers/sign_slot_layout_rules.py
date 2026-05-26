# -*- coding: utf-8 -*-
"""签字位版式配置（来自 sign_slot_layout_rules.json）。"""
from __future__ import annotations

import json
import os
import warnings
from typing import Any, Dict, List, Pattern
import re

_JSON_NAME = "sign_slot_layout_rules.json"

_DEFAULT_RULES: Dict[str, Any] = {
    "schema_version": 2,
    "layout_priority_rules": [],
    "replace_prefilled_slot": {
        "enabled": True,
        "max_text_len": 48,
        "fullmatch_patterns": [
            r"^[\u4e00-\u9fff]{2,4}$",
            r"^[A-Za-z][A-Za-z\s]{1,30}$",
            r"^\d{4}[./-]\d{1,2}(?:[./-]\d{1,2})?$",
            r"^[\u4e00-\u9fff]{2,4}\s*[/｜|]\s*\d{4}[./-]\d{1,2}(?:[./-]\d{1,2})?$",
        ],
        "search_patterns": [
            r"\d{4}[./-]\d{1,2}(?:[./-]\d{1,2})?",
            r"[\u4e00-\u9fff]{2,4}\s*[/｜|]\s*\d{4}[./-]\d{1,2}",
        ],
    },
}


def _json_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), _JSON_NAME)


def _load_rules_from_disk() -> Dict[str, Any]:
    path = _json_path()
    if not os.path.isfile(path):
        return dict(_DEFAULT_RULES)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("sign_slot_layout_rules.json 顶层必须为对象")
    out = dict(_DEFAULT_RULES)
    out.update(data)
    return out


def _compile_patterns(patterns: List[str]) -> List[Pattern[str]]:
    out: List[Pattern[str]] = []
    for p in patterns:
        try:
            out.append(re.compile(str(p), re.IGNORECASE))
        except Exception as e:
            warnings.warn(f"忽略非法签字位匹配正则: {p} ({e})", UserWarning, stacklevel=2)
    return out


def _prepare_rules(raw: Dict[str, Any]) -> Dict[str, Any]:
    slot = raw.get("replace_prefilled_slot")
    if not isinstance(slot, dict):
        slot = {}
    enabled = bool(slot.get("enabled", True))
    max_len = int(slot.get("max_text_len", 48) or 48)
    fullmatch_raw = slot.get("fullmatch_patterns") or []
    search_raw = slot.get("search_patterns") or []
    fullmatch_pats = _compile_patterns(
        [str(x) for x in fullmatch_raw if str(x).strip()]
    )
    search_pats = _compile_patterns([str(x) for x in search_raw if str(x).strip()])
    return {
        "schema_version": int(raw.get("schema_version", 2) or 2),
        "layout_priority_rules": list(raw.get("layout_priority_rules") or []),
        "replace_prefilled_slot": {
            "enabled": enabled,
            "max_text_len": max(8, min(max_len, 200)),
            "fullmatch_patterns": fullmatch_pats,
            "search_patterns": search_pats,
        },
    }


def _load_rules_safe() -> Dict[str, Any]:
    try:
        return _prepare_rules(_load_rules_from_disk())
    except Exception as e:
        warnings.warn(
            "sign_slot_layout_rules.json 加载失败，回退默认配置：" + str(e),
            UserWarning,
            stacklevel=2,
        )
        return _prepare_rules(dict(_DEFAULT_RULES))


SIGN_SLOT_LAYOUT_RULES: Dict[str, Any] = _load_rules_safe()


def validate_sign_slot_layout_rules_payload(raw: Dict[str, Any]) -> Dict[str, Any]:
    """
    校验并规范化上传的 JSON 配置（用于管理接口）。
    返回可直接落盘的 JSON 对象（仅基础类型，不含 regex 对象）。
    """
    if not isinstance(raw, dict):
        raise ValueError("规则文件顶层必须为 JSON 对象")
    out: Dict[str, Any] = {
        "schema_version": int(raw.get("schema_version", 2) or 2),
    }
    slot = raw.get("replace_prefilled_slot")
    if slot is None:
        slot = {}
    if not isinstance(slot, dict):
        raise ValueError("replace_prefilled_slot 必须为对象")
    enabled = bool(slot.get("enabled", True))
    max_len = int(slot.get("max_text_len", 48) or 48)
    max_len = max(8, min(max_len, 200))
    fullmatch_raw = slot.get("fullmatch_patterns") or []
    search_raw = slot.get("search_patterns") or []
    if not isinstance(fullmatch_raw, list):
        raise ValueError("replace_prefilled_slot.fullmatch_patterns 必须为数组")
    if not isinstance(search_raw, list):
        raise ValueError("replace_prefilled_slot.search_patterns 必须为数组")
    fullmatch = [str(x).strip() for x in fullmatch_raw if str(x).strip()]
    search = [str(x).strip() for x in search_raw if str(x).strip()]
    # 严格校验正则，防止上传后热加载才失败。
    for p in fullmatch + search:
        try:
            re.compile(p, re.IGNORECASE)
        except Exception as e:
            raise ValueError(f"非法正则: {p} ({e})")
    out["replace_prefilled_slot"] = {
        "enabled": enabled,
        "max_text_len": max_len,
        "fullmatch_patterns": fullmatch,
        "search_patterns": search,
    }
    if "layout_priority_rules" in raw:
        if not isinstance(raw["layout_priority_rules"], list):
            raise ValueError("layout_priority_rules 必须为数组")
        norm_rules = []
        for item in raw["layout_priority_rules"]:
            if not isinstance(item, dict):
                continue
            lt = str(item.get("layout_type") or "").strip()
            if not lt:
                continue
            norm_rules.append(
                {
                    "layout_type": lt,
                    "priority": int(item.get("priority", 0) or 0),
                    "match_contains": [
                        str(x).strip()
                        for x in (item.get("match_contains") or [])
                        if str(x).strip()
                    ],
                    "description": str(item.get("description") or "").strip(),
                }
            )
        out["layout_priority_rules"] = norm_rules
    # 透传文档化字段，便于后续维护（可选）
    if "signature_slot_forms" in raw:
        if not isinstance(raw["signature_slot_forms"], list):
            raise ValueError("signature_slot_forms 必须为数组")
        out["signature_slot_forms"] = [str(x) for x in raw["signature_slot_forms"]]
    if "date_layout_forms" in raw:
        if not isinstance(raw["date_layout_forms"], list):
            raise ValueError("date_layout_forms 必须为数组")
        out["date_layout_forms"] = [str(x) for x in raw["date_layout_forms"]]
    return out


def reload_sign_slot_layout_rules_from_disk() -> None:
    global SIGN_SLOT_LAYOUT_RULES
    SIGN_SLOT_LAYOUT_RULES = _load_rules_safe()


def is_replaceable_prefilled_slot_text(text: str) -> bool:
    """
    是否可视为“可替换的电脑输入签字位”（先清除再贴手写签名/日期）。
    只对短文本生效，避免把正文段落误判为签字位。
    """
    conf = SIGN_SLOT_LAYOUT_RULES["replace_prefilled_slot"]
    if not conf.get("enabled", True):
        return False
    t = str(text or "").strip()
    if not t:
        return False
    if len(t) > int(conf.get("max_text_len", 48) or 48):
        return False
    for pat in conf.get("fullmatch_patterns", []):
        if pat.fullmatch(t):
            return True
    for pat in conf.get("search_patterns", []):
        if pat.search(t):
            return True
    return False

