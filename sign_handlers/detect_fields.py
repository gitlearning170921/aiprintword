# -*- coding: utf-8 -*-
"""
自动识别签名位/签字位（用于前端自动勾选角色）。

实现遵循 .cursor/skills/signature-field-identification 的规则要点：
- 同一区域出现「角色 + 日期」 => 高置信度
- 出现「编制/审核/批准」组合 => 高置信度
- 仅单个 token（如“签字/盖章”）=> 低置信度候选
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from sign_handlers.config import ROLE_ID_TO_KEYWORD, role_keywords
from sign_handlers.label_match import cell_text_matches_keyword

_DATE_TOKENS = ("日期", "签署日期", "Date", "Signed Date")
_ACTION_TOKENS = ("签字", "签名", "签章", "盖章")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _contains_any(text: str, tokens: Sequence[str]) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    low = t.lower()
    for x in tokens:
        if not x:
            continue
        if all(ord(c) < 128 for c in x):
            if x.lower() in low:
                return True
        else:
            if x in t:
                return True
    return False


def _role_ids_for_text_label(label: str) -> List[str]:
    """把命中的标签映射成 role_id（按 ROLE_ID_TO_KEYWORD 反查）。"""
    out: List[str] = []
    for rid in ROLE_ID_TO_KEYWORD:
        for kw in role_keywords(rid):
            if cell_text_matches_keyword(label, kw):
                out.append(rid)
                break
    return out


@dataclass
class DetectedBlock:
    block_id: str
    confidence: float
    matched_rules: List[str]
    fields: List[Dict[str, str]]
    source_hint: str

    def to_dict(self) -> dict:
        return {
            "block_id": self.block_id,
            "confidence": self.confidence,
            "matched_rules": self.matched_rules,
            "fields": self.fields,
            "source_hint": self.source_hint,
        }


def _score_block(role_labels: List[str], has_date: bool, has_action: bool) -> Tuple[float, List[str]]:
    matched: List[str] = []
    roles = set(role_labels)
    if roles and has_date:
        matched.append("strong_block_rule")
    if {"author", "reviewer", "approver"}.issubset(roles) or _contains_any(" ".join(role_labels), ("编制人", "审核人", "批准人")):
        matched.append("triad_rule")
    if _contains_any(" ".join(role_labels), ("企业负责人", "法定代表人")) and (has_date or has_action):
        matched.append("org_seal_rule")
    if has_action and has_date and roles:
        matched.append("table_header_rule")

    # 置信度：强规则优先
    if "triad_rule" in matched or "strong_block_rule" in matched:
        return 0.93, matched
    if roles and (has_date or has_action):
        return 0.75, matched or ["medium_block_rule"]
    if has_action:
        return 0.45, ["single_token_fallback"]
    return 0.0, []


def detect_xlsx(path: str, max_scan_rows: int = 300, max_scan_cols: int = 60) -> dict:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=False)
    blocks: List[DetectedBlock] = []
    role_ids_found: Dict[str, float] = {}
    bi = 1

    for ws in wb.worksheets:
        # 按行扫描：同一行出现 role label + 日期 label => 一个候选 block
        for r in range(1, min((ws.max_row or 0), max_scan_rows) + 1):
            role_labels: List[str] = []
            role_ids: List[str] = []
            has_date = False
            has_action = False
            for c in range(1, min((ws.max_column or 0), max_scan_cols) + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                s = _norm(str(v))
                if not s or len(s) < 2:
                    continue
                # 角色标签（整格匹配）
                for rid in ROLE_ID_TO_KEYWORD:
                    for kw in role_keywords(rid):
                        if cell_text_matches_keyword(s, kw):
                            role_labels.append(kw)
                            role_ids.append(rid)
                            break
                if any(cell_text_matches_keyword(s, d) for d in _DATE_TOKENS):
                    has_date = True
                if any(cell_text_matches_keyword(s, a) for a in _ACTION_TOKENS):
                    has_action = True

            if not role_ids and not has_action:
                continue

            conf, matched = _score_block(role_ids, has_date, has_action)
            if conf <= 0:
                continue
            # 过滤明显噪音：无角色且仅动作词 => 候选但不作为“确认块”
            if conf < 0.6:
                continue

            fields: List[Dict[str, str]] = []
            for rid in sorted(set(role_ids)):
                fields.append({"name": rid, "type": "role_id"})
            if has_date:
                fields.append({"name": "日期", "type": "date"})
            if has_action:
                fields.append({"name": "签字/盖章", "type": "action"})

            b = DetectedBlock(
                block_id=f"xlsx_{bi:03d}",
                confidence=conf,
                matched_rules=matched,
                fields=fields,
                source_hint=f"{ws.title}!row{r}",
            )
            blocks.append(b)
            bi += 1
            for rid in set(role_ids):
                role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)

    return {
        "ok": True,
        "kind": "xlsx",
        "roles": [{"id": rid, "confidence": role_ids_found[rid]} for rid in sorted(role_ids_found)],
        "blocks": [b.to_dict() for b in blocks],
    }


def detect_docx(path: str, max_paragraphs: int = 1200) -> dict:
    from docx import Document

    doc = Document(path)
    blocks: List[DetectedBlock] = []
    role_ids_found: Dict[str, float] = {}
    bi = 1

    # 1) 表格行：强场景
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            texts = [_norm(c.text or "") for c in row.cells]
            joined = " | ".join(t for t in texts if t)
            if not joined:
                continue
            role_ids: List[str] = []
            for t in texts:
                if not t:
                    continue
                role_ids.extend(_role_ids_for_text_label(t))
            has_date = _contains_any(joined, _DATE_TOKENS)
            has_action = _contains_any(joined, _ACTION_TOKENS)

            conf, matched = _score_block(role_ids, has_date, has_action)
            if conf < 0.6:
                continue
            fields: List[Dict[str, str]] = []
            for rid in sorted(set(role_ids)):
                fields.append({"name": rid, "type": "role_id"})
            if has_date:
                fields.append({"name": "日期", "type": "date"})
            if has_action:
                fields.append({"name": "签字/盖章", "type": "action"})
            b = DetectedBlock(
                block_id=f"docx_t{ti+1}_r{ri+1}_{bi:03d}",
                confidence=conf,
                matched_rules=matched,
                fields=fields,
                source_hint=f"table{ti+1}.row{ri+1}",
            )
            blocks.append(b)
            bi += 1
            for rid in set(role_ids):
                role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)

    # 2) 段落：角色+日期同段
    for pi, p in enumerate(doc.paragraphs[:max_paragraphs]):
        t = _norm(p.text or "")
        if not t:
            continue
        # 先粗过滤：含任一角色/日期/动作
        if not (_contains_any(t, _DATE_TOKENS) or _contains_any(t, _ACTION_TOKENS) or any(_contains_any(t, role_keywords(rid)) for rid in ROLE_ID_TO_KEYWORD)):
            continue
        role_ids: List[str] = []
        for rid in ROLE_ID_TO_KEYWORD:
            for kw in role_keywords(rid):
                if kw and (kw in t or (all(ord(c) < 128 for c in kw) and kw.lower() in t.lower())):
                    role_ids.append(rid)
                    break
        has_date = _contains_any(t, _DATE_TOKENS)
        has_action = _contains_any(t, _ACTION_TOKENS)
        conf, matched = _score_block(role_ids, has_date, has_action)
        if conf < 0.85:
            continue
        fields: List[Dict[str, str]] = [{"name": rid, "type": "role_id"} for rid in sorted(set(role_ids))]
        fields.append({"name": "日期", "type": "date"})
        if has_action:
            fields.append({"name": "签字/盖章", "type": "action"})
        b = DetectedBlock(
            block_id=f"docx_p{pi+1}_{bi:03d}",
            confidence=conf,
            matched_rules=matched,
            fields=fields,
            source_hint=f"paragraph{pi+1}",
        )
        blocks.append(b)
        bi += 1
        for rid in set(role_ids):
            role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)

    return {
        "ok": True,
        "kind": "docx",
        "roles": [{"id": rid, "confidence": role_ids_found[rid]} for rid in sorted(role_ids_found)],
        "blocks": [b.to_dict() for b in blocks],
    }


def detect_file(path: str) -> dict:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return detect_xlsx(path)
    if ext == ".docx":
        return detect_docx(path)
    return {"ok": False, "error": f"不支持的格式: {ext}"}

