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
from sign_handlers.label_match import cell_text_matches_keyword, xlsx_cell_has_leading_role_keyword

_DATE_TOKENS = (
    "日期",
    "签署日期",
    "签字日期",
    "填报日期",
    "年月日",
    "年 月 日",
    "Date",
    "Signed Date",
    "Sign Date",
)
_ACTION_TOKENS = ("签字", "签名", "签章", "盖章", "签 字", "签 名")


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


def _paragraph_role_ids(text: str) -> List[str]:
    """段落/短文本中的角色标签（避免在长句中子串误匹配）。"""
    t = _norm(text or "")
    if not t:
        return []
    out: List[str] = []
    if len(t) <= 56:
        for rid in ROLE_ID_TO_KEYWORD:
            for kw in role_keywords(rid):
                if cell_text_matches_keyword(t, kw):
                    out.append(rid)
                    break
        return out
    for line in re.split(r"[\r\n]+", t):
        line = _norm(line)
        if not line or len(line) > 72:
            continue
        for rid in ROLE_ID_TO_KEYWORD:
            for kw in role_keywords(rid):
                if cell_text_matches_keyword(line, kw):
                    out.append(rid)
                    break
    return out


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


def _score_block(
    role_ids: List[str], joined_text: str, has_date: bool, has_action: bool
) -> Tuple[float, List[str]]:
    """role_ids 为英文 id（author/reviewer/…）；joined_text 为单元格拼接后的原文，用于中文关键词规则。"""
    matched: List[str] = []
    rset = set(role_ids)
    jt = joined_text or ""
    if rset and has_date:
        matched.append("strong_block_rule")
    if {"author", "reviewer", "approver"}.issubset(rset):
        matched.append("triad_rule")
    elif _contains_any(jt, ("编制人", "编制", "审核人", "审核", "批准人", "批准")) and len(rset) >= 2:
        matched.append("triad_text_hint")
    if _contains_any(jt, ("企业负责人", "法定代表人", "法人代表")) and (has_date or has_action):
        matched.append("org_seal_rule")
    if has_action and has_date and rset:
        matched.append("table_header_rule")
    if len(rset) >= 2 and has_action:
        matched.append("multi_role_action_rule")
    if len(rset) >= 1 and has_date and "strong_block_rule" not in matched:
        matched.append("role_date_rule")

    # 置信度：强规则优先
    if "triad_rule" in matched or "strong_block_rule" in matched:
        return 0.93, matched
    if "triad_text_hint" in matched:
        return 0.88, matched
    if "org_seal_rule" in matched:
        return 0.86, matched
    if "table_header_rule" in matched:
        return 0.82, matched
    if "multi_role_action_rule" in matched:
        return 0.78, matched
    if "role_date_rule" in matched:
        return 0.74, matched
    if rset and (has_date or has_action):
        return 0.72, matched or ["medium_block_rule"]
    if has_action and rset:
        return 0.55, matched or ["action_with_role"]
    if has_action:
        return 0.45, ["single_token_fallback"]
    return 0.0, []


def _table_cell_role_ids(cell_text: str) -> List[str]:
    """Word/Excel 单元格内匹配到的 role_id（整格或格首同义词）；长关键词优先。"""
    s = _norm(cell_text or "")
    if not s or len(s) > 96:
        return []
    pairs: List[Tuple[str, str]] = []
    for rid in ROLE_ID_TO_KEYWORD:
        for kw in role_keywords(rid):
            pairs.append((rid, kw))
    pairs.sort(key=lambda p: len(p[1]), reverse=True)
    out: List[str] = []
    seen: set = set()
    for rid, kw in pairs:
        if rid in seen:
            continue
        if xlsx_cell_has_leading_role_keyword(s, kw):
            seen.add(rid)
            out.append(rid)
    return out


def _xlsx_effective_scan_dims(ws, cap_rows: int, cap_cols: int) -> Tuple[int, int]:
    """工作表真实占用范围（多页/长表时 max_row 可能偏小，用 calculate_dimension 补全）。"""
    eff_r = int(ws.max_row or 1)
    eff_c = int(ws.max_column or 1)
    try:
        from openpyxl.utils import range_boundaries

        dim = ws.calculate_dimension()
        if dim:
            _min_c, _min_r, max_c, max_r = range_boundaries(dim)
            eff_r = max(eff_r, int(max_r))
            eff_c = max(eff_c, int(max_c))
    except Exception:
        pass
    return min(max(eff_r, 1), cap_rows), min(max(eff_c, 1), cap_cols)


def _harvest_roles_xlsx_cell(ws, max_scan_rows: int, max_scan_cols: int) -> Dict[str, float]:
    """逐格扫描：整格或格首为角色标签即记入（补漏单行检测不到的表头）。"""
    found: Dict[str, float] = {}
    for r in range(1, max_scan_rows + 1):
        for c in range(1, max_scan_cols + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = _norm(str(v))
            if not s or len(s) > 96:
                continue
            for rid in _table_cell_role_ids(s):
                found[rid] = max(found.get(rid, 0.0), 0.58)
    return found


def _xlsx_row_sign_signals(
    ws,
    r: int,
    max_c: int,
    scan_last_row: int,
    _dtoks: Sequence[str],
    _atoks: Sequence[str],
) -> Tuple[List[str], bool, bool, List[str]]:
    """扫描第 r 行（并合并 r+1 行常见日期/签字格）得到角色、日期/动作标记与拼接文本。"""
    role_ids: List[str] = []
    has_date = False
    has_action = False
    cell_texts: List[str] = []
    for c in range(1, max_c + 1):
        v = ws.cell(row=r, column=c).value
        if v is None:
            continue
        s = _norm(str(v))
        if not s or len(s) < 2:
            continue
        cell_texts.append(s)
        role_ids.extend(_table_cell_role_ids(s))
        if any(cell_text_matches_keyword(s, d) for d in _DATE_TOKENS) or any(
            xlsx_cell_has_leading_role_keyword(s, d) for d in _dtoks
        ):
            has_date = True
        if any(cell_text_matches_keyword(s, a) for a in _ACTION_TOKENS) or any(
            xlsx_cell_has_leading_role_keyword(s, a) for a in _atoks
        ):
            has_action = True

    if r < scan_last_row:
        for c in range(1, max_c + 1):
            v2 = ws.cell(row=r + 1, column=c).value
            if v2 is None:
                continue
            s2 = _norm(str(v2))
            if not s2:
                continue
            cell_texts.append(s2)
            if any(cell_text_matches_keyword(s2, d) for d in _DATE_TOKENS) or any(
                xlsx_cell_has_leading_role_keyword(s2, d) for d in _dtoks
            ):
                has_date = True
            if any(cell_text_matches_keyword(s2, a) for a in _ACTION_TOKENS) or any(
                xlsx_cell_has_leading_role_keyword(s2, a) for a in _atoks
            ):
                has_action = True

    return role_ids, has_date, has_action, cell_texts


def _append_xlsx_block(
    blocks: List[DetectedBlock],
    bi_holder: List[int],
    role_ids: List[str],
    joined: str,
    has_date: bool,
    has_action: bool,
    conf: float,
    matched: List[str],
    sheet_title: str,
    row_index: int,
    role_ids_found: Dict[str, float],
) -> None:
    fields: List[Dict[str, str]] = []
    for rid in sorted(set(role_ids)):
        fields.append({"name": rid, "type": "role_id"})
    if has_date:
        fields.append({"name": "日期", "type": "date"})
    if has_action:
        fields.append({"name": "签字/盖章", "type": "action"})
    bi = bi_holder[0]
    bi_holder[0] = bi + 1
    b = DetectedBlock(
        block_id=f"xlsx_{bi:03d}",
        confidence=conf,
        matched_rules=matched,
        fields=fields,
        source_hint=f"{sheet_title}!row{row_index}",
    )
    blocks.append(b)
    for rid in set(role_ids):
        role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)


def detect_xlsx(path: str, max_scan_rows: int = 8000, max_scan_cols: int = 128) -> dict:
    from openpyxl import load_workbook

    # data_only=True 读取公式格缓存结果，避免格内为「=」开头公式时识别不到显示文字
    wb = load_workbook(path, data_only=True)
    blocks: List[DetectedBlock] = []
    role_ids_found: Dict[str, float] = {}
    bi_holder = [1]

    for ws in wb.worksheets:
        eff_rows, eff_cols = _xlsx_effective_scan_dims(ws, max_scan_rows, max_scan_cols)

        # 逐格补漏（与有效扫描范围一致，避免只扫到「第一页」）
        for rid, conf in _harvest_roles_xlsx_cell(ws, eff_rows, eff_cols).items():
            role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)

        _dtoks = sorted(_DATE_TOKENS, key=len, reverse=True)
        _atoks = sorted(_ACTION_TOKENS, key=len, reverse=True)
        rows_with_block = set()

        for r in range(1, eff_rows + 1):
            role_ids, has_date, has_action, cell_texts = _xlsx_row_sign_signals(
                ws, r, eff_cols, eff_rows, _dtoks, _atoks
            )
            joined = " | ".join(cell_texts)
            if not role_ids and not has_action:
                continue

            conf, matched = _score_block(role_ids, joined, has_date, has_action)
            if conf <= 0:
                continue
            if conf < 0.5:
                continue

            _append_xlsx_block(
                blocks,
                bi_holder,
                role_ids,
                joined,
                has_date,
                has_action,
                conf,
                matched,
                ws.title or "?",
                r,
                role_ids_found,
            )
            rows_with_block.add(r)

        # 补漏：后面几页/另一块签字区常因无日期列等导致分数<0.5 被整行丢弃，这里按行再收一遍
        for r in range(1, eff_rows + 1):
            if r in rows_with_block:
                continue
            role_ids, has_date, has_action, cell_texts = _xlsx_row_sign_signals(
                ws, r, eff_cols, eff_rows, _dtoks, _atoks
            )
            if not role_ids:
                continue
            joined = " | ".join(cell_texts)
            conf, matched = _score_block(role_ids, joined, has_date, has_action)
            if conf < 0.5:
                if len(set(role_ids)) >= 2 or has_date or has_action:
                    conf = 0.56
                    matched = list(matched or []) + ["xlsx_multi_page_supplement"]
                elif len(set(role_ids)) == 1 and has_date:
                    conf = 0.56
                    matched = list(matched or []) + ["xlsx_single_role_date_supplement"]
                else:
                    continue
            if conf <= 0:
                continue

            _append_xlsx_block(
                blocks,
                bi_holder,
                role_ids,
                joined,
                has_date,
                has_action,
                conf,
                matched,
                ws.title or "?",
                r,
                role_ids_found,
            )

    for b in blocks:
        for f in b.fields:
            if f.get("type") == "role_id" and f.get("name") in ROLE_ID_TO_KEYWORD:
                rid = f["name"]
                role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), float(b.confidence) * 0.95)

    return {
        "ok": True,
        "kind": "xlsx",
        "roles": [{"id": rid, "confidence": role_ids_found[rid]} for rid in sorted(role_ids_found)],
        "blocks": [b.to_dict() for b in blocks],
    }


def _harvest_roles_docx_tables(doc) -> Dict[str, float]:
    found: Dict[str, float] = {}
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                s = _norm(cell.text or "")
                if not s or len(s) > 96:
                    continue
                for rid in _table_cell_role_ids(s):
                    found[rid] = max(found.get(rid, 0.0), 0.58)
    return found


def detect_docx(path: str, max_paragraphs: int = 1200) -> dict:
    from docx import Document

    doc = Document(path)
    blocks: List[DetectedBlock] = []
    role_ids_found: Dict[str, float] = {}
    bi = 1

    for rid, conf in _harvest_roles_docx_tables(doc).items():
        role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), conf)

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
                role_ids.extend(_table_cell_role_ids(t))
            has_date = _contains_any(joined, _DATE_TOKENS)
            has_action = _contains_any(joined, _ACTION_TOKENS)

            conf, matched = _score_block(role_ids, joined, has_date, has_action)
            if conf < 0.5:
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
        if not (
            _contains_any(t, _DATE_TOKENS)
            or _contains_any(t, _ACTION_TOKENS)
            or any(_contains_any(t, role_keywords(rid)) for rid in ROLE_ID_TO_KEYWORD)
        ):
            continue
        role_ids = _paragraph_role_ids(t)
        has_date = _contains_any(t, _DATE_TOKENS)
        has_action = _contains_any(t, _ACTION_TOKENS)
        conf, matched = _score_block(role_ids, t, has_date, has_action)
        if conf < 0.68:
            continue
        fields: List[Dict[str, str]] = [{"name": rid, "type": "role_id"} for rid in sorted(set(role_ids))]
        if has_date:
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

    for b in blocks:
        for f in b.fields:
            if f.get("type") == "role_id" and f.get("name") in ROLE_ID_TO_KEYWORD:
                rid = f["name"]
                role_ids_found[rid] = max(role_ids_found.get(rid, 0.0), float(b.confidence) * 0.95)

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

