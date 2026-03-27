# -*- coding: utf-8 -*-
"""在线签名：角色 id 与文档内关键词对应（前端用英文 id 提交）。"""

# 每个角色可对应多个模板用词（中英混排表格常见英文标签）
ROLE_ID_TO_KEYWORD = {
    "author": ("作者", "Author"),
    "reviewer": ("审核", "Reviewer", "Review"),
    "approver": ("批准", "Approver", "Approval"),
    "executor": ("执行人员", "Executor"),
    "reviewer_tail": ("审核人员", "Reviewer"),
}


def role_keywords(role_id: str) -> tuple[str, ...]:
    v = ROLE_ID_TO_KEYWORD[role_id]
    return (v,) if isinstance(v, str) else tuple(v)


def role_display_name(role_id: str) -> str:
    """错误提示等：优先取中文关键词。"""
    for k in role_keywords(role_id):
        if any("\u4e00" <= ch <= "\u9fff" for ch in k):
            return k
    return role_keywords(role_id)[0]
