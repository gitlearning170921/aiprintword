# -*- coding: utf-8 -*-
"""
运行时配置持久化：MySQL（与在线签名共用 MYSQL_* 环境变量）。
连接参数仅从 os.environ 读取，避免与 get_setting() 循环依赖（库内覆盖的 MYSQL_* 不改变连接目标）。
"""
from __future__ import annotations

import os
import re
import threading
import time
from typing import Any, Dict, List, Optional, Tuple

import pymysql
from pymysql.cursors import DictCursor

_lock = threading.RLock()
_DB_NAME_RE = re.compile(r"^[a-zA-Z0-9_]{1,64}$")

# 与签名等业务表同库，独立表名
SETTINGS_TABLE = "app_runtime_settings"


def _config_env_only() -> Dict[str, Any]:
    port = os.environ.get("MYSQL_PORT", "3306").strip() or "3306"
    try:
        port_i = int(port)
    except ValueError:
        port_i = 3306
    db = (os.environ.get("MYSQL_DATABASE") or "aiprintword_sign").strip()
    if not _DB_NAME_RE.match(db):
        db = "aiprintword_sign"
    return {
        "host": os.environ.get("MYSQL_HOST", "").strip(),
        "port": port_i,
        "user": (os.environ.get("MYSQL_USER") or "root").strip(),
        "password": os.environ.get("MYSQL_PASSWORD") or "",
        "charset": (os.environ.get("MYSQL_CHARSET") or "utf8mb4").strip(),
        "database": db,
    }


def mysql_settings_enabled() -> bool:
    return bool(_config_env_only()["host"])


def _connect_server():
    c = _config_env_only()
    return pymysql.connect(
        host=c["host"],
        port=c["port"],
        user=c["user"],
        password=c["password"],
        charset=c["charset"],
        cursorclass=DictCursor,
    )


def _connect_db():
    c = _config_env_only()
    return pymysql.connect(
        host=c["host"],
        port=c["port"],
        user=c["user"],
        password=c["password"],
        database=c["database"],
        charset=c["charset"],
        cursorclass=DictCursor,
    )


def init_db() -> None:
    if not mysql_settings_enabled():
        return
    with _lock:
        c = _config_env_only()
        conn = _connect_server()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"CREATE DATABASE IF NOT EXISTS `{c['database']}` "
                    "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
                )
                cur.execute(f"USE `{c['database']}`")
                cur.execute(
                    f"""
                    CREATE TABLE IF NOT EXISTS `{SETTINGS_TABLE}` (
                        setting_key VARCHAR(255) NOT NULL PRIMARY KEY,
                        setting_value TEXT NOT NULL,
                        updated_at DOUBLE NOT NULL
                    ) ENGINE=InnoDB DEFAULT CHARACTER SET utf8mb4
                    COLLATE utf8mb4_unicode_ci
                    """
                )
            conn.commit()
        finally:
            conn.close()


def get_all_rows() -> List[Tuple[str, str, float]]:
    if not mysql_settings_enabled():
        return []
    init_db()
    with _lock:
        conn = _connect_db()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT setting_key, setting_value, updated_at "
                    f"FROM `{SETTINGS_TABLE}` ORDER BY setting_key"
                )
                rows = cur.fetchall()
                return [
                    (str(r["setting_key"]), str(r["setting_value"]), float(r["updated_at"]))
                    for r in rows
                ]
        finally:
            conn.close()


def get_value(key: str) -> Optional[str]:
    if not mysql_settings_enabled():
        return None
    init_db()
    with _lock:
        conn = _connect_db()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT setting_value FROM `{SETTINGS_TABLE}` "
                    "WHERE setting_key = %s",
                    (key,),
                )
                row = cur.fetchone()
                return str(row["setting_value"]) if row else None
        finally:
            conn.close()


def set_values(updates: Dict[str, str]) -> None:
    if not updates or not mysql_settings_enabled():
        return
    init_db()
    now = time.time()
    with _lock:
        conn = _connect_db()
        try:
            with conn.cursor() as cur:
                sql = (
                    f"INSERT INTO `{SETTINGS_TABLE}` "
                    "(setting_key, setting_value, updated_at) VALUES (%s, %s, %s) "
                    "ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value), "
                    "updated_at = VALUES(updated_at)"
                )
                for k, v in updates.items():
                    cur.execute(sql, (k, v, now))
            conn.commit()
        finally:
            conn.close()


def delete_key(key: str) -> None:
    if not mysql_settings_enabled():
        return
    init_db()
    with _lock:
        conn = _connect_db()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"DELETE FROM `{SETTINGS_TABLE}` WHERE setting_key = %s", (key,)
                )
            conn.commit()
        finally:
            conn.close()
