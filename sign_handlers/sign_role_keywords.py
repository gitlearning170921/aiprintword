# -*- coding: utf-8 -*-
"""
从 sign_role_keywords.json 加载签字角色同义词；便于手工扩展而无需改 Python 代码。

匹配用词 = 各角色下 synonyms ∪ zh ∪ en 去重（顺序保留）。
"""
from __future__ import annotations

import json
import os
import warnings
from typing import Any, Dict, List, Tuple, Union

_JSON_NAME = "sign_role_keywords.json"


def _json_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), _JSON_NAME)


def _dedupe_preserve(words: List[str]) -> Tuple[str, ...]:
    seen = set()
    out: List[str] = []
    for w in words:
        s = str(w).strip()
        if not s or s in seen:
            continue
        seen.add(s)
        out.append(s)
    return tuple(out)


def _parse_role_entry(spec: Union[list, dict, str, None]) -> Tuple[str, ...]:
    if spec is None:
        return tuple()
    if isinstance(spec, str):
        return _dedupe_preserve([spec])
    if isinstance(spec, list):
        return _dedupe_preserve([str(x) for x in spec])
    if isinstance(spec, dict):
        merged: List[str] = []
        for key in ("synonyms", "zh", "en"):
            part = spec.get(key)
            if isinstance(part, str):
                merged.append(part)
            elif isinstance(part, list):
                merged.extend(str(x) for x in part)
        return _dedupe_preserve(merged)
    return tuple()


def load_role_id_to_keyword() -> Dict[str, Tuple[str, ...]]:
    path = _json_path()
    if not os.path.isfile(path):
        raise FileNotFoundError(f"缺少签字同义词规则文件: {path}")
    with open(path, encoding="utf-8") as f:
        data: Any = json.load(f)
    roles = data.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("sign_role_keywords.json: 根节点需包含 roles 对象")
    out: Dict[str, Tuple[str, ...]] = {}
    for rid, spec in roles.items():
        rid = str(rid).strip()
        if not rid:
            continue
        out[rid] = _parse_role_entry(spec)
    return out


_FALLBACK_ROLE_KEYWORDS: Dict[str, Tuple[str, ...]] = {
    "author": ("作者", "Author"),
    "reviewer": ("审核", "Reviewer"),
    "approver": ("批准", "Approver"),
    "executor": ("执行人", "Executor"),
    "reviewer_tail": ("审核人员", "QA"),
}


def _load_role_keywords_safe() -> Dict[str, Tuple[str, ...]]:
    try:
        return load_role_id_to_keyword()
    except Exception as e:
        warnings.warn(
            "sign_role_keywords.json 加载失败，使用内置最小兜底（请检查 JSON）：" + str(e),
            UserWarning,
            stacklevel=2,
        )
        return dict(_FALLBACK_ROLE_KEYWORDS)


ROLE_ID_TO_KEYWORD: Dict[str, Tuple[str, ...]] = _load_role_keywords_safe()


def reload_role_keywords_from_disk() -> None:
    """测试或热重载用（生产一般需重启服务）。"""
    global ROLE_ID_TO_KEYWORD
    ROLE_ID_TO_KEYWORD = _load_role_keywords_safe()


def role_keywords(role_id: str) -> Tuple[str, ...]:
    v = ROLE_ID_TO_KEYWORD.get(role_id)
    if v is None:
        raise KeyError(f"未知 role_id: {role_id}")
    return v


# 文末「审核人员/复核人员」与 reviewer 同义；QA/会签等仍用 reviewer_tail 单独识别
_REVIEWER_TAIL_CANONICAL_TO_REVIEWER = frozenset(
    {
        "审核人员",
        "复核人员",
        "审核组员",
        "审核成员",
    }
)

# 签字落位时：为 reviewer / executor 合并额外同义词（长词优先由调用方排序）
_ROLE_APPLY_EXTRA_IDS: Dict[str, Tuple[str, ...]] = {
    "reviewer": ("reviewer_tail",),
    "executor": (),
}


def canonical_sign_role_id(role_id: str, matched_keyword: str | None = None) -> str:
    """将 detect/映射中的 role_id 规范为签字用的 author/reviewer/approver/executor。"""
    rid = str(role_id or "").strip()
    if not rid:
        return rid
    if rid == "reviewer_tail":
        kw = str(matched_keyword or "").strip()
        if not kw:
            return "reviewer_tail"
        if kw in _REVIEWER_TAIL_CANONICAL_TO_REVIEWER:
            return "reviewer"
        for x in _REVIEWER_TAIL_CANONICAL_TO_REVIEWER:
            if kw.startswith(x):
                return "reviewer"
        return "reviewer_tail"
    return rid


def normalize_role_signer_map(mapping: dict | None) -> Dict[str, Any]:
    """读库后合并 reviewer_tail→reviewer，避免旧映射缺 executor/reviewer 落位。"""
    if not isinstance(mapping, dict):
        return {}
    out: Dict[str, Any] = {}
    for rid, pair in mapping.items():
        rid2 = canonical_sign_role_id(str(rid))
        if rid2 not in ROLE_ID_TO_KEYWORD:
            continue
        if not isinstance(pair, dict):
            continue
        if rid2 not in out:
            out[rid2] = dict(pair)
            continue
        cur = out[rid2]
        for k, v in pair.items():
            if v and not cur.get(k):
                cur[k] = v
    return out


def role_keywords_for_apply(role_id: str) -> Tuple[str, ...]:
    """生成文档时使用的同义词（含 reviewer_tail 中 QA/会签等，便于表格多标签落位）。"""
    rid = canonical_sign_role_id(role_id)
    merged: List[str] = list(role_keywords(rid))
    seen = set(merged)
    for extra_rid in _ROLE_APPLY_EXTRA_IDS.get(rid, ()):
        for kw in role_keywords(extra_rid):
            if kw not in seen:
                seen.add(kw)
                merged.append(kw)
    return tuple(merged)
