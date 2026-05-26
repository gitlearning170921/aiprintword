# -*- coding: utf-8 -*-
"""从 T2 的 4 份 Markdown 候选文件生成可合并规则草稿。"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path


ROLE_CATEGORY_TO_ID = {
    "编写": "author",
    "审核": "reviewer",
    "批准": "approver",
    "执行人员位": "executor",
}


def _read_text(path: Path) -> str:
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="ignore")


def _parse_role_keywords(md_text: str) -> dict[str, list[str]]:
    """
    解析 signature_label_candidates_T2.md 中的汇总表：
    | 归类 | 文案标记 | 命中次数 | 来源文件数 |
    """
    out: dict[str, list[str]] = defaultdict(list)
    for raw in md_text.splitlines():
        s = raw.strip()
        if not s.startswith("|"):
            continue
        cols = [x.strip() for x in s.strip("|").split("|")]
        if len(cols) < 4:
            continue
        if cols[0] in ("归类", "---"):
            continue
        cat = cols[0]
        kw = cols[1]
        if not kw or kw in ("文案标记", "???"):
            continue
        rid = ROLE_CATEGORY_TO_ID.get(cat)
        if not rid:
            continue
        if kw not in out[rid]:
            out[rid].append(kw)
    return dict(out)


def _parse_layout_forms(md_text: str) -> list[dict]:
    """
    解析 signature_slot_layout_full_rescan_T2.md 中的版式统计表：
    | 形式 | 命中次数 | 来源文件数 |
    """
    forms: list[dict] = []
    for raw in md_text.splitlines():
        s = raw.strip()
        if not s.startswith("|"):
            continue
        cols = [x.strip() for x in s.strip("|").split("|")]
        if len(cols) < 3:
            continue
        name = cols[0]
        if name in ("形式", "---", "??", "????"):
            continue
        if not name:
            continue
        m = re.search(r"(\d+)", cols[1] or "")
        count = int(m.group(1)) if m else 0
        forms.append({"name": name, "count": count})
    # 按命中次数降序，便于人工优先固化高频版式
    forms.sort(key=lambda x: (-int(x.get("count", 0)), str(x.get("name", ""))))
    return forms


def _parse_no_sign_patterns(md_text: str) -> list[str]:
    """
    解析 signature_role_results_T2.md 中 no_sign 小节的 pattern 列表：
    - `XXX` — ...
    """
    out: list[str] = []
    in_no_sign = False
    for raw in md_text.splitlines():
        s = raw.strip()
        if s.startswith("## 无需签字"):
            in_no_sign = True
            continue
        if in_no_sign and s.startswith("## "):
            break
        if not in_no_sign:
            continue
        m = re.match(r"^- `([^`]+)`", s)
        if not m:
            continue
        pat = m.group(1).strip()
        if pat and pat not in out:
            out.append(pat)
    return out


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    md_label = root / "signature_label_candidates_T2.md"
    md_role = root / "signature_role_results_T2.md"
    md_layout = root / "signature_slot_layout_candidates_T2.md"
    md_layout_full = root / "signature_slot_layout_full_rescan_T2.md"

    txt_label = _read_text(md_label)
    txt_role = _read_text(md_role)
    # 优先用 full_rescan，若缺失则回退 candidates。
    txt_layout = _read_text(md_layout_full) or _read_text(md_layout)

    role_keywords = _parse_role_keywords(txt_label)
    layout_forms = _parse_layout_forms(txt_layout)
    no_sign_patterns = _parse_no_sign_patterns(txt_role)

    drafts_dir = root / "sign_handlers" / "drafts"
    drafts_dir.mkdir(parents=True, exist_ok=True)

    role_draft = {
        "schema_version": 1,
        "source": "signature_label_candidates_T2.md + signature_role_results_T2.md",
        "generated_by": "scripts/generate_sign_rule_drafts_from_md.py",
        "role_keywords_incremental": role_keywords,
        "no_sign_pattern_candidates": no_sign_patterns,
    }
    slot_draft = {
        "schema_version": 1,
        "source": "signature_slot_layout_full_rescan_T2.md",
        "generated_by": "scripts/generate_sign_rule_drafts_from_md.py",
        "layout_forms_ranked": layout_forms,
        "layout_priority_rules": [
            {
                "layout_type": "two_row_signoff_table",
                "description": "两行签批表：优先角色行右侧签名格 + 日期行右侧日期格",
                "priority": 100,
                "match_contains": ["两行签批表"],
            },
            {
                "layout_type": "same_cell_inline",
                "description": "同字段长空格/下划线占位，优先同格内联插入",
                "priority": 80,
                "match_contains": ["同字段长空格占位", "同字段下划线/点线占位"],
            },
            {
                "layout_type": "adjacent_right_cell",
                "description": "右侧相邻/隔列空白格落位",
                "priority": 60,
                "match_contains": ["右侧空白单元格", "日期在右相邻单元格", "日期在右侧隔列单元格"],
            },
        ],
    }

    role_path = drafts_dir / "sign_role_keywords.incremental.json"
    slot_path = drafts_dir / "sign_slot_layout.incremental.json"
    role_path.write_text(json.dumps(role_draft, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    slot_path.write_text(json.dumps(slot_draft, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("wrote", role_path)
    print("wrote", slot_path)


if __name__ == "__main__":
    main()
