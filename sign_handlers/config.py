# -*- coding: utf-8 -*-
"""在线签名：角色 id 与同义词（自 sign_role_keywords.json 加载）。"""

from sign_handlers.sign_role_keywords import ROLE_ID_TO_KEYWORD, role_keywords

__all__ = ["ROLE_ID_TO_KEYWORD", "role_keywords", "role_display_name"]


def role_display_name(role_id: str) -> str:
    """错误提示等：优先取中文关键词。"""
    for k in role_keywords(role_id):
        if any("\u4e00" <= ch <= "\u9fff" for ch in k):
            return k
    return role_keywords(role_id)[0]
