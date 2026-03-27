# -*- coding: utf-8 -*-
"""
在线签名：待签文件列表与文件二进制存入 MySQL，多机/多浏览器共享。
连接参数仅从环境变量读取（可在项目根目录放置 .env，由 app 加载，勿提交 .env）。

环境变量：
  MYSQL_HOST     非空则启用 MySQL 存储（为空则回退 session+本地目录）
  MYSQL_PORT     默认 3306
  MYSQL_USER     默认 root
  MYSQL_PASSWORD
  MYSQL_CHARSET  默认 utf8mb4
  MYSQL_DATABASE 库名，默认 aiprintword_sign（不存在则自动创建）

表 sign_signed_output：每次「生成已签名文档」成功时由 app 写入（需配置 MYSQL_HOST），供多机下载。
"""
from __future__ import annotations

import os
import re
import threading
from contextlib import contextmanager
from typing import Any, Dict, List, Optional

import pymysql
from pymysql.cursors import DictCursor

_DB_NAME_RE = re.compile(r"^[a-zA-Z0-9_]{1,64}$")

_mysql_init_lock = threading.Lock()
_mysql_inited = False


def mysql_sign_enabled() -> bool:
    return bool(os.environ.get("MYSQL_HOST", "").strip())


def _config() -> Dict[str, Any]:
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


def _connect_server():
    """连接 MySQL 服务器（不指定库），用于 CREATE DATABASE。"""
    c = _config()
    return pymysql.connect(
        host=c["host"],
        port=c["port"],
        user=c["user"],
        password=c["password"],
        charset=c["charset"],
        cursorclass=DictCursor,
    )


def _connect_db():
    c = _config()
    return pymysql.connect(
        host=c["host"],
        port=c["port"],
        user=c["user"],
        password=c["password"],
        database=c["database"],
        charset=c["charset"],
        cursorclass=DictCursor,
    )


def init_schema() -> None:
    """创建库、表（若不存在）。"""
    c = _config()
    if not c["host"]:
        return
    conn = _connect_server()
    try:
        with conn.cursor() as cur:
            cur.execute(
                f"CREATE DATABASE IF NOT EXISTS `{c['database']}` "
                "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            cur.execute(f"USE `{c['database']}`")
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_uploaded_file (
                    id VARCHAR(32) NOT NULL PRIMARY KEY,
                    original_name VARCHAR(512) NOT NULL,
                    ext VARCHAR(16) NOT NULL,
                    file_data LONGBLOB NOT NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    KEY idx_created_at (created_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_signed_output (
                    id VARCHAR(32) NOT NULL PRIMARY KEY,
                    source_file_id VARCHAR(32) NULL,
                    source_name VARCHAR(512) NOT NULL,
                    output_name VARCHAR(512) NOT NULL,
                    ext VARCHAR(16) NOT NULL,
                    roles_json TEXT NOT NULL,
                    file_data LONGBLOB NOT NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    KEY idx_signed_created (created_at),
                    KEY idx_signed_source (source_file_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
        conn.commit()
    finally:
        conn.close()


def ensure_sign_mysql() -> None:
    """首次调用时初始化库表（线程安全）。"""
    global _mysql_inited
    if not mysql_sign_enabled():
        return
    with _mysql_init_lock:
        if _mysql_inited:
            return
        init_schema()
        _mysql_inited = True


@contextmanager
def _conn_commit():
    conn = _connect_db()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def list_files() -> List[dict]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, original_name AS name, ext, created_at "
                "FROM sign_uploaded_file ORDER BY created_at DESC"
            )
            rows = cur.fetchall()
    out = []
    for r in rows:
        ts = r.get("created_at")
        out.append(
            {
                "id": r["id"],
                "name": r["name"],
                "ext": r["ext"],
                "created_at": ts.isoformat(sep=" ") if ts is not None else None,
            }
        )
    return out


def count_files() -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) AS c FROM sign_uploaded_file")
            row = cur.fetchone()
            return int(row["c"]) if row else 0


def insert_file(file_id: str, original_name: str, ext: str, file_data: bytes) -> None:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO sign_uploaded_file (id, original_name, ext, file_data) "
                "VALUES (%s, %s, %s, %s)",
                (file_id, original_name, ext, file_data),
            )


def delete_file(file_id: str) -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_uploaded_file WHERE id=%s", (file_id,))
            return cur.rowcount


def get_file_row(file_id: str) -> Optional[dict]:
    """返回 id, name, ext, file_data(bytes)；不存在则 None。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, original_name AS name, ext, file_data "
                "FROM sign_uploaded_file WHERE id=%s",
                (file_id,),
            )
            return cur.fetchone()


def _ensure_signed_output_table(conn) -> None:
    """幂等创建已签名结果表（兼容已部署实例在未重启前缺表的情况）。"""
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_signed_output (
                id VARCHAR(32) NOT NULL PRIMARY KEY,
                source_file_id VARCHAR(32) NULL,
                source_name VARCHAR(512) NOT NULL,
                output_name VARCHAR(512) NOT NULL,
                ext VARCHAR(16) NOT NULL,
                roles_json TEXT NOT NULL,
                file_data LONGBLOB NOT NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_signed_created (created_at),
                KEY idx_signed_source (source_file_id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )


def count_signed_outputs() -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) AS c FROM sign_signed_output")
            row = cur.fetchone()
            return int(row["c"]) if row else 0


def list_signed_outputs() -> List[dict]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, source_file_id, source_name, output_name, ext, roles_json, created_at "
                "FROM sign_signed_output ORDER BY created_at DESC"
            )
            rows = cur.fetchall()
    out: List[dict] = []
    for r in rows:
        ts = r.get("created_at")
        out.append(
            {
                "id": r["id"],
                "source_file_id": r.get("source_file_id"),
                "source_name": r.get("source_name"),
                "name": r.get("output_name"),
                "ext": r.get("ext"),
                "roles_json": r.get("roles_json"),
                "created_at": ts.isoformat(sep=" ") if ts is not None else None,
            }
        )
    return out


def insert_signed_output(
    signed_id: str,
    source_file_id: Optional[str],
    source_name: str,
    output_name: str,
    ext: str,
    roles_json: str,
    file_data: bytes,
) -> None:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO sign_signed_output "
                "(id, source_file_id, source_name, output_name, ext, roles_json, file_data) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s)",
                (
                    signed_id,
                    source_file_id,
                    source_name[:512],
                    output_name[:512],
                    ext[:16],
                    roles_json,
                    file_data,
                ),
            )


def get_signed_row(signed_id: str) -> Optional[dict]:
    """返回含 file_data(bytes)；不存在则 None。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, output_name AS name, ext, file_data "
                "FROM sign_signed_output WHERE id=%s",
                (signed_id,),
            )
            return cur.fetchone()


def delete_signed_output(signed_id: str) -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_signed_output WHERE id=%s", (signed_id,))
            return cur.rowcount
