# -*- coding: utf-8 -*-
"""为 sign_document_role_rules.json 补全 sign_policy / label / runtime_policy。"""
import json
from pathlib import Path

ROLE_LABELS = {
    "author": "编写/编制",
    "executor": "执行/测试",
    "reviewer": "审核/复核",
    "approver": "批准",
}

USE_CASE_SPEC_PATTERNS = {
    "单元测试用例表",
    "总结性可用性测试用例表",
    "兼容性测试用例表",
    "系统测试用例表",
    "网络安全测试用例表",
    "传统算法测试用例表",
}


def roles_label(roles):
    if not roles:
        return "无需签字"
    return "、".join(ROLE_LABELS.get(r, r) for r in roles)


def use_case_spec_category(pattern: str) -> str:
    if pattern in USE_CASE_SPEC_PATTERNS:
        return "use_case_spec_table"
    if "用例表" in pattern and "用例执行" not in pattern:
        return "use_case_spec_table"
    return ""


def main():
    path = Path(__file__).resolve().parents[1] / "sign_handlers" / "sign_document_role_rules.json"
    raw = json.loads(path.read_text(encoding="utf-8"))
    rules = raw.get("rules") or []
    for item in rules:
        if not isinstance(item, dict):
            continue
        roles = item.get("roles") or []
        pat = str(item.get("pattern") or "")
        if roles:
            item["sign_policy"] = "detect_roles"
            item["label"] = roles_label(roles)
            item.pop("no_sign_required", None)
            item.pop("category", None)
        else:
            item["sign_policy"] = "no_sign"
            item["no_sign_required"] = True
            item["label"] = "无需签字"
            cat = use_case_spec_category(pat)
            if cat:
                item["category"] = cat
                if not item.get("note"):
                    item["note"] = (
                        "人工约定：用例表不签字；对应用例执行表需 executor+reviewer。"
                    )
            elif not item.get("note"):
                item["note"] = (
                    "本类附件无签批栏，不进签字流程；仍会轻量解析文档（非用例表快速跳过）。"
                )

    rs = raw.get("review_summary")
    if isinstance(rs, dict):
        for g in rs.get("groups") or []:
            if not isinstance(g, dict):
                continue
            r = g.get("roles") or []
            g["label"] = roles_label(r)
            g["sign_policy"] = "no_sign" if not r else "detect_roles"

    raw["schema_version"] = 2
    raw["updated"] = "2026-05-22"
    raw["runtime_policy"] = {
        "match": "文件名 contains 主题片段（pattern），不含项目编号/版本号/扩展名",
        "detect_roles": "sign_policy=detect_roles：正常识别签字角色并进入匹配/签字流程",
        "no_sign": "sign_policy=no_sign：工作台「无需签字」，roles 强制为空",
        "use_case_spec_table": "category=use_case_spec_table：用例表；批量可跳过签字位识别",
        "use_case_execution": "含「用例执行表/用例执行记录」：识别 executor+reviewer",
        "other_files": "未命中规则：按正文签批栏自动识别",
    }
    path.write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("ok rules=", len(rules))


if __name__ == "__main__":
    main()
