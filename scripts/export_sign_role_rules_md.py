# -*- coding: utf-8 -*-
"""从 sign_document_role_rules.json 导出 signature_role_results_T2.md。"""
import json
from collections import defaultdict
from pathlib import Path

ROLE_LABELS = {
    "author": "编写/编制",
    "executor": "执行/测试",
    "reviewer": "审核/复核",
    "approver": "批准",
}


def roles_label(roles):
    if not roles:
        return "无需签字"
    return "、".join(ROLE_LABELS.get(r, r) for r in roles)


def main():
    root = Path(__file__).resolve().parents[1]
    jpath = root / "sign_handlers" / "sign_document_role_rules.json"
    mpath = root / "signature_role_results_T2.md"
    raw = json.loads(jpath.read_text(encoding="utf-8"))
    rules = [r for r in (raw.get("rules") or []) if isinstance(r, dict)]
    rp = raw.get("runtime_policy") or {}

    by_key = defaultdict(list)
    for item in rules:
        roles = tuple(item.get("roles") or [])
        by_key[(roles, item.get("sign_policy") or ("no_sign" if not roles else "detect_roles"))].append(item)

    lines = [
        "# T2 每份文件需签字角色识别结果清单",
        "",
        f"- 配置文件: `sign_handlers/sign_document_role_rules.json`（schema_version={raw.get('schema_version', 1)}）",
        f"- 来源: {raw.get('source', '')}",
        f"- 更新: {raw.get('updated', '')}",
        f"- 规则条数: {len(rules)}（与 T2 压缩包 107 份文件主题一一对应）",
        "- 角色 id: `author`=编写/编制，`executor`=执行/测试，`reviewer`=审核/复核，`approver`=批准",
        "",
        "## 系统运行时策略（与 JSON `runtime_policy` 一致）",
        "",
    ]
    for k, v in rp.items():
        lines.append(f"- **{k}**: {v}")
    lines.extend(
        [
            "",
            "## 分组统计",
            "",
            "| 策略 | 角色组合 | roles | 规则条数 |",
            "| --- | --- | --- | ---: |",
        ]
    )

    groups_sorted = sorted(
        by_key.items(),
        key=lambda x: (-len(x[1]), x[0][1], x[0][0]),
    )
    for (roles, policy), items in groups_sorted:
        lines.append(
            f"| `{policy}` | {roles_label(list(roles))} | "
            f"`{', '.join(roles) if roles else '(none)'}` | {len(items)} |"
        )

    lines.extend(["", "## 用例表 / 用例执行表（人工约定，核对重点）", ""])
    lines.append("| 类型 | pattern | sign_policy | 说明 |")
    lines.append("| --- | --- | --- | --- |")
    for item in rules:
        pat = str(item.get("pattern") or "")
        if "用例" not in pat:
            continue
        sp = item.get("sign_policy") or ""
        cat = item.get("category") or ""
        note = str(item.get("note") or "")[:80]
        lines.append(f"| {cat or '—'} | `{pat}` | `{sp}` | {note} |")

    lines.extend(["", "## 无需签字（sign_policy=no_sign）", ""])
    no_sign = [r for r in rules if not (r.get("roles") or [])]
    lines.append(f"共 {len(no_sign)} 条规则（含用例表 7 类主题 + 评审/问卷等附件）。")
    lines.append("")
    for item in sorted(no_sign, key=lambda x: str(x.get("pattern") or "")):
        pat = item.get("pattern") or ""
        ex = item.get("source_example") or ""
        cat = item.get("category") or ""
        note = item.get("note") or ""
        extra = f"；category=`{cat}`" if cat else ""
        note_s = (note or "").rstrip("。.")
        lines.append(f"- `{pat}` — {note_s}{extra}")
        if ex:
            lines.append(f"  - 示例: `{ex}`")

    lines.extend(["", "## 需识别签字角色（sign_policy=detect_roles）", ""])
    detect = [r for r in rules if r.get("roles")]
    lines.append(f"共 {len(detect)} 条规则。")
    lines.append("")
    by_roles = defaultdict(list)
    for item in detect:
        by_roles[tuple(item.get("roles") or [])].append(item)
    for roles, items in sorted(by_roles.items(), key=lambda x: (-len(x[1]), x[0])):
        lines.append(f"### {roles_label(list(roles))}（{len(items)} 条）")
        lines.append("")
        lines.append(f"- roles: `{', '.join(roles)}`")
        lines.append("")
        for item in sorted(items, key=lambda x: str(x.get("pattern") or "")):
            pat = item.get("pattern") or ""
            ex = item.get("source_example") or ""
            note = item.get("note") or ""
            suffix = f" — {note}" if note else ""
            lines.append(f"- `{pat}`{suffix}")
            if ex:
                lines.append(f"  - 示例: `{ex}`")
        lines.append("")

    lines.extend(
        [
            "## 本次人工修正（摘要）",
            "",
            "- **用例表**（文件名含「用例表」且不含「用例执行表」）：`sign_policy=no_sign`，工作台 **无需签字**，批量可跳过签字位识别。",
            "- **用例执行表 / 用例执行记录**：`executor` + `reviewer`，正常识别与匹配。",
            "- **兼容性/系统测试方案/报告**等：仅 `author`+`reviewer`+`approver`，去除误识别执行人员。",
            "- **配置管理计划 / 配置状态报告**：正文表格作者/复核 → `author`+`reviewer`+`approver`。",
            "- **代码评审报告**：作者=编写，研发负责人=审核 → `author`+`reviewer`。",
            "",
            "> 修改规则请编辑 JSON 后运行 `python scripts/sync_sign_role_rules_metadata.py` 与 `python scripts/export_sign_role_rules_md.py` 同步本文件。",
            "",
        ]
    )

    mpath.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("wrote", mpath)


if __name__ == "__main__":
    main()
