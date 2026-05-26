# -*- coding: utf-8 -*-
"""将人工标误登记同步到 sign_document_role_rules.json 并导出 MD。"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from sign_handlers.config import ROLE_ID_TO_KEYWORD
from sign_handlers.detect_correction import trim_detect_correction

_ROOT = Path(__file__).resolve().parents[1]
_RULES_PATH = _ROOT / "sign_handlers" / "sign_document_role_rules.json"
_MD_PATH = _ROOT / "signature_role_results_T2.md"

_VERSION_SUFFIX_RE = re.compile(
    r"[\s_]*(?:[（(]\s*[Aa]?\d+(?:\.\d+)?\s*[）)]|[Vv]\d+(?:\.\d+)?)\s*$"
)


def _strip_ext(name: str) -> str:
    s = str(name or "").strip()
    for ext in (".docx", ".docm", ".xlsx", ".xls", ".xlsm", ".pdf"):
        if s.lower().endswith(ext):
            return s[: -len(ext)].strip()
    return s


def infer_pattern_from_source_name(source_name: str) -> str:
    """从展示文件名推断 contains 用 pattern（与 T2 规则风格一致）。"""
    full = str(source_name or "").replace("\\", "/").strip()
    parts = [p.strip() for p in full.split("/") if p.strip() and p not in (".", "..")]
    candidates: List[str] = []
    if parts:
        last = _VERSION_SUFFIX_RE.sub("", _strip_ext(parts[-1])).strip()
        last = re.sub(r"^附件\s*", "", last).strip()
        if len(last) >= 4:
            candidates.append(last)
    base = _VERSION_SUFFIX_RE.sub("", _strip_ext(os.path.basename(full))).strip()
    base = re.sub(r"^附件\s*", "", base).strip()
    if len(base) >= 4 and base not in candidates:
        candidates.append(base)
    if not candidates:
        return (base or last or "")[:80]
    candidates.sort(key=len, reverse=True)
    return candidates[0][:80]


def _infer_category(pattern: str, roles: List[str]) -> str:
    if roles:
        return ""
    pat = str(pattern or "")
    if "测试任务" in pat:
        return "test_task_no_sign"
    if "用例表" in pat and "用例执行" not in pat:
        return "use_case_spec_table"
    return ""


def _roles_label(roles: List[str]) -> str:
    labels = {
        "author": "编写/编制",
        "executor": "执行/测试",
        "reviewer": "审核/复核",
        "approver": "批准",
    }
    if not roles:
        return "无需签字"
    return "、".join(labels.get(r, r) for r in roles)


def correction_to_rule_entry(
    correction: Dict[str, Any],
    *,
    source_name: str,
    pattern: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    corr = trim_detect_correction(correction)
    wrong = str(corr.get("wrong_description") or "").strip()
    if not wrong:
        return None
    pat = (pattern or infer_pattern_from_source_name(source_name) or "").strip()
    if len(pat) < 4:
        return None
    roles = [r for r in (corr.get("expected_roles") or []) if r in ROLE_ID_TO_KEYWORD]
    note_parts = [f"人工标误同步：{wrong[:400]}"]
    exp_note = str(corr.get("expected_note") or "").strip()
    if exp_note:
        note_parts.append(exp_note[:200])
    esl = corr.get("expected_slot_layout")
    if isinstance(esl, dict) and esl:
        slot_bits = []
        arr = str(esl.get("arrangement") or "").strip()
        if arr == "horizontal":
            slot_bits.append("角色左右排")
        elif arr == "vertical":
            slot_bits.append("角色上下排")
        rel = str(esl.get("date_relation") or "").strip()
        if rel == "same_cell":
            slot_bits.append("名日同格")
        elif rel == "different_cell":
            slot_bits.append("名日分格")
        pos = str(esl.get("date_position") or "").strip()
        if pos == "right":
            slot_bits.append("日期在右")
        elif pos == "below":
            slot_bits.append("日期在下")
        sep = str(esl.get("separator") or "").strip()
        if sep == "slash":
            slot_bits.append("分隔/")
        elif sep == "cell":
            slot_bits.append("分隔单元格")
        elif sep == "space":
            slot_bits.append("分隔空格")
        if slot_bits:
            note_parts.append("签字位：" + "、".join(slot_bits))
    entry: Dict[str, Any] = {
        "pattern": pat,
        "match": "contains",
        "roles": roles,
        "note": "；".join(note_parts)[:600],
        "source_example": str(source_name or "")[:300],
        "learned_from_correction": True,
    }
    if roles:
        entry["sign_policy"] = "detect_roles"
        entry["label"] = _roles_label(roles)
        entry.pop("no_sign_required", None)
        entry.pop("category", None)
    else:
        entry["sign_policy"] = "no_sign"
        entry["no_sign_required"] = True
        entry["label"] = "无需签字"
        cat = _infer_category(pat, roles)
        if cat:
            entry["category"] = cat
    return entry


def _load_rules_raw() -> Dict[str, Any]:
    if not _RULES_PATH.is_file():
        return {"schema_version": 2, "rules": []}
    with open(_RULES_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)
    if not isinstance(raw, dict):
        return {"schema_version": 2, "rules": []}
    if not isinstance(raw.get("rules"), list):
        raw["rules"] = []
    return raw


def _save_rules_raw(raw: Dict[str, Any]) -> None:
    raw["updated"] = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    _RULES_PATH.write_text(
        json.dumps(raw, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    try:
        from sign_handlers import sign_document_role_rules as sdr

        sdr.load_sign_document_role_rules(force=True)
    except Exception:
        pass


def _find_rule_index(
    rules: List[Dict[str, Any]],
    *,
    source_name: str,
    pattern: str,
) -> Optional[int]:
    from sign_handlers.sign_document_role_rules import match_document_role_rule

    matched = match_document_role_rule(source_name)
    if matched:
        mp = str(matched.get("pattern") or "").strip()
        for i, item in enumerate(rules):
            if str(item.get("pattern") or "").strip() == mp:
                return i
    pat = str(pattern or "").strip()
    if pat:
        best_i = None
        best_len = -1
        for i, item in enumerate(rules):
            ip = str(item.get("pattern") or "").strip()
            if ip == pat or (ip and ip in pat) or (pat and pat in ip):
                if len(ip) > best_len:
                    best_len = len(ip)
                    best_i = i
        if best_i is not None:
            return best_i
    return None


def upsert_rule_from_correction(
    source_name: str,
    correction: Dict[str, Any],
    *,
    pattern: Optional[str] = None,
) -> Dict[str, Any]:
    """
    将标误登记写入 sign_document_role_rules.json。
    返回 {ok, action, pattern, sign_policy, roles, error?}
    """
    entry = correction_to_rule_entry(
        correction, source_name=source_name, pattern=pattern
    )
    if not entry:
        return {
            "ok": False,
            "error": "缺少「错在哪」或无法从文件名推断 pattern",
        }
    raw = _load_rules_raw()
    rules: List[Dict[str, Any]] = [
        r for r in (raw.get("rules") or []) if isinstance(r, dict)
    ]
    pat = str(entry.get("pattern") or "")
    idx = _find_rule_index(rules, source_name=source_name, pattern=pat)
    action = "updated"
    if idx is None:
        rules.append(entry)
        action = "created"
    else:
        prev = dict(rules[idx])
        prev.update(entry)
        rules[idx] = prev
    raw["rules"] = rules
    _save_rules_raw(raw)
    return {
        "ok": True,
        "action": action,
        "pattern": pat,
        "sign_policy": entry.get("sign_policy"),
        "roles": list(entry.get("roles") or []),
        "category": entry.get("category") or "",
    }


def _run_script(rel_path: str) -> Tuple[bool, str]:
    script = _ROOT / rel_path
    if not script.is_file():
        return False, f"脚本不存在: {rel_path}"
    try:
        proc = subprocess.run(
            [sys.executable, str(script)],
            cwd=str(_ROOT),
            capture_output=True,
            text=True,
            timeout=120,
        )
        if proc.returncode != 0:
            err = (proc.stderr or proc.stdout or "").strip()
            return False, err[:500] or f"exit {proc.returncode}"
        return True, ""
    except Exception as e:
        return False, str(e)[:500]


def export_rules_markdown() -> Dict[str, Any]:
    ok, err = _run_script("scripts/sync_sign_role_rules_metadata.py")
    if not ok:
        return {"ok": False, "error": "sync_metadata: " + err}
    ok2, err2 = _run_script("scripts/export_sign_role_rules_md.py")
    if not ok2:
        return {"ok": False, "error": "export_md: " + err2}
    return {"ok": True, "md_path": str(_MD_PATH)}


def sync_rules_from_correction(
    source_name: str,
    correction: Dict[str, Any],
    *,
    export_md: bool = False,
) -> Dict[str, Any]:
    """标误 → 更新角色/签字位 JSON；可选导出对应 MD。"""
    from sign_handlers.detect_correction import _trim_expected_slot_layout

    corr = trim_detect_correction(correction)
    wrong = str(corr.get("wrong_description") or "").strip()
    esl = _trim_expected_slot_layout(corr.get("expected_slot_layout"))
    has_slot = bool(esl)
    scopes = corr.get("correction_save") if isinstance(corr.get("correction_save"), dict) else {}
    save_roles = scopes.get("roles", True) if scopes else True
    save_slot = scopes.get("slot", True) if scopes else True

    if not wrong and not (has_slot and save_slot):
        return {"ok": False, "error": "缺少「错在哪」或签字位版式登记"}

    out: Dict[str, Any] = {"ok": True}
    if save_roles and wrong:
        up = upsert_rule_from_correction(source_name, correction)
        if not up.get("ok"):
            return up
        out.update(up)
        if export_md:
            ex = export_rules_markdown()
            out["md_exported"] = bool(ex.get("ok"))
            if not ex.get("ok"):
                out["md_warning"] = ex.get("error") or "导出角色 MD 失败"
        else:
            out["md_exported"] = False

    if save_slot and (has_slot or wrong):
        try:
            from sign_handlers.detect_correction_slot_rules import (
                sync_slot_rules_from_correction,
            )

            slot_up = sync_slot_rules_from_correction(
                source_name, correction, export_md=export_md
            )
            out["slot_rule_sync"] = slot_up
            if slot_up.get("ok") and slot_up.get("md_exported"):
                out["slot_md_exported"] = True
                out["slot_md_path"] = slot_up.get("md_path")
        except Exception as e:
            out["slot_rule_sync"] = {"ok": False, "error": str(e)[:500]}

    if not wrong and has_slot:
        out.setdefault("action", out.get("slot_rule_sync", {}).get("action"))
        out.setdefault("pattern", out.get("slot_rule_sync", {}).get("pattern"))

    return out
