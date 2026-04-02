# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from typing import Any, Dict, List

from runtime_settings import db as settings_db
from runtime_settings.registry import REGISTRY, coerce_value, ordered_keys


def invalidate_cache() -> None:
    """保留接口；当前每次均解析，无进程内缓存。"""


def get_setting(key: str) -> Any:
    """数据库 > 环境变量 > 注册表默认值。"""
    meta = REGISTRY.get(key)
    if meta is None:
        return os.environ.get(key, "")

    db_s = settings_db.get_value(key)
    if db_s is not None:
        return coerce_value(meta, db_s)
    env_s = os.environ.get(key)
    if env_s is not None and env_s != "":
        return coerce_value(meta, env_s)
    return meta.default


def list_all_settings(*, mask_secrets: bool = True) -> List[dict]:
    """合并 DB + env + 默认，供管理页展示。"""
    rows = {k: v for k, v, _ in settings_db.get_all_rows()}
    out = []
    for key in ordered_keys():
        meta = REGISTRY[key]
        src = "default"
        raw = None
        if key in rows:
            raw = rows[key]
            src = "database"
        else:
            env_s = os.environ.get(key)
            if env_s is not None and env_s != "":
                raw = env_s
                src = "environment"
        if raw is None:
            if meta.value_type == "bool":
                raw = "1" if meta.default else "0"
            elif meta.value_type in ("int", "float"):
                raw = str(meta.default)
            else:
                raw = str(meta.default)

        masked = mask_secrets and meta.is_secret and bool(raw)
        out.append(
            {
                "key": key,
                "group": meta.group,
                "label": meta.label,
                "description": meta.description,
                "value_type": meta.value_type,
                "is_secret": meta.is_secret,
                "source": src,
                "raw": "********" if masked else raw,
                "has_secret_value": bool(meta.is_secret and raw),
            }
        )
    return out


def set_settings(updates: Dict[str, str]) -> List[str]:
    """校验键名后写入 MySQL。"""
    from runtime_settings import db as _db

    if not _db.mysql_settings_enabled():
        raise RuntimeError(
            "未配置 MYSQL_HOST：运行时配置需写入 MySQL，请在 .env 中设置 MYSQL_HOST 等连接参数"
        )
    bad = [k for k in updates if k not in REGISTRY]
    if bad:
        raise ValueError("未知配置项: " + ", ".join(bad))
    to_write = {}
    for k, v in updates.items():
        meta = REGISTRY[k]
        if meta.is_secret and (v is None or str(v).strip() == ""):
            continue
        to_write[k] = str(v).strip()
    if to_write:
        settings_db.set_values(to_write)
    invalidate_cache()
    return list(to_write.keys())


def reset_keys_to_env_or_default(keys: List[str]) -> None:
    for k in keys:
        if k in REGISTRY:
            settings_db.delete_key(k)
    invalidate_cache()
