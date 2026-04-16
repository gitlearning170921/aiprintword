# -*- coding: utf-8 -*-
"""
在线签名：待签文件列表与记录存入 MySQL，多机/多浏览器共享；文件内容存入 FTP（主动模式）。
连接参数仅从环境变量读取（可在项目根目录放置 .env，由 app 加载，勿提交 .env）。

环境变量：
  MYSQL_HOST     非空则启用 MySQL 存储（为空则回退 session+本地目录）
  MYSQL_PORT     默认 3306
  MYSQL_USER     默认 root
  MYSQL_PASSWORD
  MYSQL_CHARSET  默认 utf8mb4
  MYSQL_DATABASE 库名，默认 aiprintword_sign（不存在则自动创建）

表 sign_signed_output：每次「生成已签名文档」成功时由 app 写入（需配置 MYSQL_HOST），供多机下载。
表 sign_signer / sign_signer_stroke：签署人及可复用的签名、日期笔迹（PNG）；legacy 每人一行，与最新一套笔迹同步。
表 sign_stroke_item：签名/日期素材分开存储；同一签署人同语言同 kind 同内容（sha256）重复提交则覆盖。
表 sign_stroke_set：历史兼容（旧版「签名+日期」成对存储）。
表 sign_file_role_signer：每个待签文件各签字角色对应的笔迹套 stroke_set_id（及冗余 signer_id）。
表 app_batch_history：批量打印任务历史（列表与详情 JSON）；暂存目录仍在服务端 data/batch_history/<id>/stash。
"""
from __future__ import annotations

import hashlib
import os
import re
import threading
import uuid
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple, Union

import pymysql
from pymysql.cursors import DictCursor

from sign_handlers.date_piece_compose import (
    all_piece_kinds,
    compose_png_horizontal,
    kinds_en_dot_dmy,
    kinds_en_dmy_space,
    kinds_for_iso_date,
    kinds_zh_ymd_dot,
    normalize_piece_kind,
    piece_kind_label,
)

_DB_NAME_RE = re.compile(r"^[a-zA-Z0-9_]{1,64}$")

# 笔迹元件拼接日期：与前端 date_mode 一致
_COMPOSITE_DATE_MODES = frozenset({"composite_en", "composite_zh_ymd", "composite_en_space"})


def is_composite_date_mode(dm: Optional[str]) -> bool:
    return (dm or "").strip().lower() in _COMPOSITE_DATE_MODES


def composite_mode_to_layout(dm: Optional[str]) -> str:
    """date_mode → compose_date_piece_png 的 layout 参数。"""
    k = (dm or "").strip().lower()
    if k == "composite_zh_ymd":
        return "zh_ymd"
    if k == "composite_en_space":
        return "en_space"
    # 兼容旧值 composite_en（原 15.April.2026）：统一为英文空格版
    if k == "composite_en":
        return "en_space"
    return "en_space"


def _sign_ftp_required() -> bool:
    try:
        from runtime_settings.resolve import get_setting

        return bool(get_setting("SIGN_FTP_REQUIRED"))
    except Exception:
        return False


def _ftp_upload_bytes_or_mysql(data: bytes, remote_rel: str) -> Tuple[Optional[str], Optional[str]]:
    """
    优先 FTP；失败返回 (None, err_msg) 供写入 ftp_last_error，并由调用方将内容落 MySQL BLOB。
    未配置 FTP 时返回 (None, None)。SIGN_FTP_REQUIRED 为真且上传失败时抛出异常。
    """
    from ftp_store import try_upload_bytes

    path, err = try_upload_bytes(data, remote_rel)
    if path:
        return path, None
    if err is None:
        return None, None
    if _sign_ftp_required():
        raise RuntimeError(err)
    return None, err


_mysql_init_lock = threading.Lock()
_mysql_inited = False


def mysql_sign_enabled() -> bool:
    try:
        from runtime_settings.resolve import get_setting

        return bool(str(get_setting("MYSQL_HOST") or "").strip())
    except Exception:
        return bool(os.environ.get("MYSQL_HOST", "").strip())


def _config() -> Dict[str, Any]:
    try:
        from runtime_settings.resolve import get_setting

        port = str(get_setting("MYSQL_PORT") or "3306").strip() or "3306"
        try:
            port_i = int(port)
        except ValueError:
            port_i = 3306
        db = str(get_setting("MYSQL_DATABASE") or "aiprintword_sign").strip()
        if not _DB_NAME_RE.match(db):
            db = "aiprintword_sign"
        return {
            "host": str(get_setting("MYSQL_HOST") or "").strip(),
            "port": port_i,
            "user": str(get_setting("MYSQL_USER") or "root").strip(),
            "password": str(get_setting("MYSQL_PASSWORD") or ""),
            "charset": str(get_setting("MYSQL_CHARSET") or "utf8mb4").strip(),
            "database": db,
        }
    except Exception:
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
                    ftp_path VARCHAR(768) NULL,
                    file_size BIGINT NULL,
                    sha256 CHAR(64) NULL,
                    file_data LONGBLOB NULL,
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
                    ftp_path VARCHAR(768) NULL,
                    file_size BIGINT NULL,
                    sha256 CHAR(64) NULL,
                    file_data LONGBLOB NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    KEY idx_signed_created (created_at),
                    KEY idx_signed_source (source_file_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_batch_history (
                    id CHAR(32) NOT NULL PRIMARY KEY,
                    created_at VARCHAR(32) NOT NULL,
                    record_json LONGTEXT NOT NULL,
                    KEY idx_batch_hist_created (created_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_signer (
                    id VARCHAR(32) NOT NULL PRIMARY KEY,
                    display_name VARCHAR(128) NOT NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    KEY idx_signer_created (created_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_signer_stroke (
                    signer_id VARCHAR(32) NOT NULL PRIMARY KEY,
                    sig_png LONGBLOB NULL,
                    date_png LONGBLOB NULL,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    CONSTRAINT fk_stroke_signer FOREIGN KEY (signer_id)
                        REFERENCES sign_signer (id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_stroke_set (
                    id VARCHAR(32) NOT NULL PRIMARY KEY,
                    signer_id VARCHAR(32) NOT NULL,
                    locale VARCHAR(8) NOT NULL DEFAULT 'zh',
                    sig_sha256 CHAR(64) NOT NULL,
                    date_sha256 CHAR(64) NOT NULL,
                    sig_ftp_path VARCHAR(768) NULL,
                    date_ftp_path VARCHAR(768) NULL,
                    sig_size BIGINT NULL,
                    date_size BIGINT NULL,
                    sig_png LONGBLOB NULL,
                    date_png LONGBLOB NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    UNIQUE KEY uk_stroke_sig_date (signer_id, locale, sig_sha256, date_sha256),
                    KEY idx_ss_signer (signer_id)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_stroke_item (
                    id VARCHAR(32) NOT NULL PRIMARY KEY,
                    signer_id VARCHAR(32) NOT NULL,
                    locale VARCHAR(8) NOT NULL DEFAULT 'zh',
                    kind VARCHAR(8) NOT NULL,
                    sha256 CHAR(64) NOT NULL,
                    ftp_path VARCHAR(768) NULL,
                    file_size BIGINT NULL,
                    png LONGBLOB NULL,
                    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    UNIQUE KEY uk_item (signer_id, locale, kind, sha256),
                    KEY idx_item_signer (signer_id),
                    KEY idx_item_kind (kind),
                    CONSTRAINT fk_item_signer FOREIGN KEY (signer_id)
                        REFERENCES sign_signer (id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS sign_file_role_signer (
                    file_id VARCHAR(32) NOT NULL,
                    role_id VARCHAR(64) NOT NULL,
                    signer_id VARCHAR(32) NOT NULL,
                    stroke_set_id VARCHAR(32) NULL,
                    sig_item_id VARCHAR(32) NULL,
                    date_item_id VARCHAR(32) NULL,
                    PRIMARY KEY (file_id, role_id),
                    KEY idx_frs_file (file_id),
                    KEY idx_frs_stroke_set (stroke_set_id),
                    KEY idx_frs_sig_item (sig_item_id),
                    KEY idx_frs_date_item (date_item_id),
                    CONSTRAINT fk_frs_file FOREIGN KEY (file_id)
                        REFERENCES sign_uploaded_file (id) ON DELETE CASCADE,
                    CONSTRAINT fk_frs_signer FOREIGN KEY (signer_id)
                        REFERENCES sign_signer (id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
        conn.commit()
    finally:
        conn.close()


def _ensure_sign_file_columns(conn) -> None:
    """为已存在的旧表补齐列/放宽 BLOB 约束（向后兼容）。"""
    try:
        with conn.cursor() as cur:
            # sign_uploaded_file
            for sql in (
                "ALTER TABLE sign_uploaded_file ADD COLUMN ftp_path VARCHAR(768) NULL",
                "ALTER TABLE sign_uploaded_file ADD COLUMN file_size BIGINT NULL",
                "ALTER TABLE sign_uploaded_file ADD COLUMN sha256 CHAR(64) NULL",
            ):
                try:
                    cur.execute(sql)
                except Exception:
                    pass
            try:
                cur.execute("ALTER TABLE sign_uploaded_file MODIFY COLUMN file_data LONGBLOB NULL")
            except Exception:
                pass
            # sign_signed_output
            for sql in (
                "ALTER TABLE sign_signed_output ADD COLUMN batch_id VARCHAR(32) NULL",
                "ALTER TABLE sign_signed_output ADD COLUMN ftp_path VARCHAR(768) NULL",
                "ALTER TABLE sign_signed_output ADD COLUMN file_size BIGINT NULL",
                "ALTER TABLE sign_signed_output ADD COLUMN sha256 CHAR(64) NULL",
            ):
                try:
                    cur.execute(sql)
                except Exception:
                    pass
            try:
                cur.execute("CREATE INDEX idx_signed_batch ON sign_signed_output (batch_id)")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_signed_output MODIFY COLUMN file_data LONGBLOB NULL")
            except Exception:
                pass
            for sql in (
                "ALTER TABLE sign_uploaded_file ADD COLUMN ftp_last_error VARCHAR(512) NULL",
                "ALTER TABLE sign_signed_output ADD COLUMN ftp_last_error VARCHAR(512) NULL",
            ):
                try:
                    cur.execute(sql)
                except Exception:
                    pass
    except Exception:
        pass


def ensure_sign_mysql() -> None:
    """首次调用时初始化库表（线程安全）。"""
    global _mysql_inited
    if not mysql_sign_enabled():
        return
    with _mysql_init_lock:
        if _mysql_inited:
            return
        init_schema()
        try:
            conn = _connect_db()
            try:
                _ensure_signed_output_table(conn)
                _ensure_sign_file_columns(conn)
                conn.commit()
            finally:
                conn.close()
        except Exception:
            pass
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
                "SELECT id, original_name AS name, ext, created_at, ftp_path, ftp_last_error, "
                " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob "
                "FROM sign_uploaded_file ORDER BY created_at DESC"
            )
            rows = cur.fetchall()
    out = []
    for r in rows:
        ts = r.get("created_at")
        ftp_path = (r.get("ftp_path") or "").strip()
        fe = (r.get("ftp_last_error") or "").strip()
        out.append(
            {
                "id": r["id"],
                "name": r["name"],
                "ext": r["ext"],
                "created_at": ts.isoformat(sep=" ") if ts is not None else None,
                "ftp_uploaded": bool(ftp_path),
                "blob_stored": bool(r.get("has_blob")),
                "ftp_last_error": fe or None,
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
    size = int(len(file_data or b""))
    sha = hashlib.sha256(file_data or b"").hexdigest()
    # 记录里保留原始文件名（可含中文）；FTP 路径使用安全文件名避免乱码/不支持字符
    def _safe_ftp_filename(name: str) -> str:
        base = os.path.basename(name or "document")
        base = base.strip() or "document"
        base = re.sub(r"[^0-9A-Za-z._()\-\s]+", "_", base)
        base = re.sub(r"\s+", " ", base).strip()
        return (base or "document")[:200]

    safe_name = _safe_ftp_filename(original_name or "document")
    remote_rel = f"sign/inbox/{file_id}/{safe_name}"
    ftp_path, ftp_err = _ftp_upload_bytes_or_mysql(file_data or b"", remote_rel)
    err_s = (ftp_err or "")[:512] if ftp_err else None
    with _conn_commit() as conn:
        _ensure_sign_file_columns(conn)
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO sign_uploaded_file (id, original_name, ext, ftp_path, file_size, sha256, file_data, ftp_last_error) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                (
                    file_id,
                    original_name,
                    ext,
                    ftp_path,
                    size,
                    sha,
                    None if ftp_path else file_data,
                    None if ftp_path else err_s,
                ),
            )


def delete_file(file_id: str) -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        # best-effort delete FTP
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT ftp_path FROM sign_uploaded_file WHERE id=%s",
                    (file_id,),
                )
                row = cur.fetchone()
            p = (row or {}).get("ftp_path")
            if p:
                try:
                    from ftp_store import delete_path

                    delete_path(p)
                except Exception:
                    pass
        except Exception:
            pass
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_uploaded_file WHERE id=%s", (file_id,))
            return cur.rowcount


def get_file_row(file_id: str) -> Optional[dict]:
    """返回 id, name, ext, file_data(bytes)；不存在则 None。优先从 FTP 取，兼容旧 BLOB。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, original_name AS name, ext, ftp_path, ftp_last_error, file_data "
                "FROM sign_uploaded_file WHERE id=%s",
                (file_id,),
            )
            row = cur.fetchone()
    if not row:
        return None
    if row.get("file_data"):
        return row
    p = (row.get("ftp_path") or "").strip()
    if p:
        import time

        from ftp_store import download_bytes

        last_err: Optional[Exception] = None
        for attempt in range(3):
            try:
                row["file_data"] = download_bytes(p)
                last_err = None
                break
            except Exception as e:
                last_err = e
                if attempt < 2:
                    time.sleep(0.35 * (attempt + 1))
        if last_err is not None:
            try:
                print(f"[sign] get_file_row FTP retry exhausted id={file_id}: {last_err}")
            except Exception:
                pass
    return row


def _ensure_signed_output_table(conn) -> None:
    """幂等创建已签名结果表（兼容已部署实例在未重启前缺表的情况）。"""
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_signed_output (
                id VARCHAR(32) NOT NULL PRIMARY KEY,
                batch_id VARCHAR(32) NULL,
                source_file_id VARCHAR(32) NULL,
                source_name VARCHAR(512) NOT NULL,
                output_name VARCHAR(512) NOT NULL,
                ext VARCHAR(16) NOT NULL,
                roles_json TEXT NOT NULL,
                ftp_path VARCHAR(768) NULL,
                file_size BIGINT NULL,
                sha256 CHAR(64) NULL,
                file_data LONGBLOB NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_signed_created (created_at),
                KEY idx_signed_source (source_file_id),
                KEY idx_signed_batch (batch_id)
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


def _signed_output_row_to_item(r: dict) -> dict:
    ts = r.get("created_at")
    ftp_path = (r.get("ftp_path") or "").strip()
    fe = (r.get("ftp_last_error") or "").strip()
    return {
        "id": r["id"],
        "batch_id": r.get("batch_id"),
        "source_file_id": r.get("source_file_id"),
        "source_name": r.get("source_name"),
        "name": r.get("output_name"),
        "ext": r.get("ext"),
        "roles_json": r.get("roles_json"),
        "ftp_uploaded": bool(ftp_path),
        "ftp_path": ftp_path or None,
        "blob_stored": bool(r.get("has_blob")),
        "ftp_last_error": fe or None,
        "created_at": ts.isoformat(sep=" ") if ts is not None else None,
    }


def list_signed_outputs_page(
    *,
    q: str = "",
    page: int = 1,
    page_size: int = 10,
) -> Tuple[List[dict], int]:
    """已签名文档列表：按输出文件名/源文件名模糊搜索，分页。"""
    ensure_sign_mysql()
    q = (q or "").strip()
    page = max(1, int(page))
    page_size = max(1, min(int(page_size), 500))
    offset = (page - 1) * page_size
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            if q:
                like = f"%{q}%"
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_signed_output "
                    "WHERE output_name LIKE %s OR source_name LIKE %s",
                    (like, like),
                )
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output "
                    "WHERE output_name LIKE %s OR source_name LIKE %s "
                    "ORDER BY created_at DESC LIMIT %s OFFSET %s",
                    (like, like, page_size, offset),
                )
            else:
                cur.execute("SELECT COUNT(*) AS c FROM sign_signed_output")
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output ORDER BY created_at DESC LIMIT %s OFFSET %s",
                    (page_size, offset),
                )
            rows = cur.fetchall() or []
    out = [_signed_output_row_to_item(dict(r)) for r in rows]
    return out, total


def list_signed_outputs() -> List[dict]:
    items, _ = list_signed_outputs_page(q="", page=1, page_size=500000)
    return items


def insert_signed_output(
    signed_id: str,
    batch_id: Optional[str],
    source_file_id: Optional[str],
    source_name: str,
    output_name: str,
    ext: str,
    roles_json: str,
    file_data: bytes,
) -> None:
    ensure_sign_mysql()
    size = int(len(file_data or b""))
    sha = hashlib.sha256(file_data or b"").hexdigest()
    # 记录里保留 output_name（可含中文）；FTP 路径使用安全文件名避免乱码/不支持字符
    def _safe_ftp_filename(name: str) -> str:
        base = os.path.basename(name or "signed")
        base = base.strip() or "signed"
        base = re.sub(r"[^0-9A-Za-z._()\-\s]+", "_", base)
        base = re.sub(r"\s+", " ", base).strip()
        return (base or "signed")[:200]

    safe_name = _safe_ftp_filename(output_name or "signed")
    remote_rel = f"sign/output/{signed_id}/{safe_name}"
    ftp_path, ftp_err = _ftp_upload_bytes_or_mysql(file_data or b"", remote_rel)
    err_s = (ftp_err or "")[:512] if ftp_err else None
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        _ensure_sign_file_columns(conn)
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO sign_signed_output "
                "(id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, file_size, sha256, file_data, ftp_last_error) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                (
                    signed_id,
                    batch_id,
                    source_file_id,
                    source_name[:512],
                    output_name[:512],
                    ext[:16],
                    roles_json,
                    ftp_path,
                    size,
                    sha,
                    None if ftp_path else file_data,
                    None if ftp_path else err_s,
                ),
            )


def list_signed_batches_page(*, q: str = "", page: int = 1, page_size: int = 10):
    """
    按 batch_id 聚合的批次列表（分页、搜索）。
    搜索 q 会匹配 output_name/source_name/batch_id。
    """
    ensure_sign_mysql()
    q = (q or "").strip()
    page = max(1, int(page))
    page_size = max(1, min(int(page_size), 200))
    offset = (page - 1) * page_size
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            where = "WHERE batch_id IS NOT NULL AND batch_id <> ''"
            args = []
            if q:
                like = f"%{q}%"
                where += " AND (batch_id LIKE %s OR output_name LIKE %s OR source_name LIKE %s)"
                args.extend([like, like, like])
            cur.execute(
                "SELECT COUNT(DISTINCT batch_id) AS c FROM sign_signed_output " + where,
                tuple(args),
            )
            total = int((cur.fetchone() or {}).get("c") or 0)
            cur.execute(
                "SELECT batch_id, MAX(created_at) AS created_at, COUNT(*) AS n "
                "FROM sign_signed_output "
                + where
                + " GROUP BY batch_id ORDER BY MAX(created_at) DESC LIMIT %s OFFSET %s",
                tuple(args + [page_size, offset]),
            )
            rows = cur.fetchall() or []
            # legacy：无 batch_id 的历史记录条数
            if q:
                like = f"%{q}%"
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_signed_output "
                    "WHERE (batch_id IS NULL OR batch_id='') AND (output_name LIKE %s OR source_name LIKE %s)",
                    (like, like),
                )
            else:
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_signed_output WHERE (batch_id IS NULL OR batch_id='')"
                )
            legacy_total = int((cur.fetchone() or {}).get("c") or 0)
    out = []
    for r in rows:
        ts = r.get("created_at")
        out.append(
            {
                "batch_id": r.get("batch_id"),
                "created_at": ts.isoformat(sep=" ") if ts is not None else None,
                "n": int(r.get("n") or 0),
            }
        )
    return out, total, legacy_total


def list_signed_outputs_by_batch(*, batch_id: str, q: str = "") -> list[dict]:
    ensure_sign_mysql()
    q = (q or "").strip()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            if q:
                like = f"%{q}%"
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output WHERE batch_id=%s AND (output_name LIKE %s OR source_name LIKE %s) "
                    "ORDER BY created_at DESC",
                    (batch_id, like, like),
                )
            else:
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output WHERE batch_id=%s ORDER BY created_at DESC",
                    (batch_id,),
                )
            rows = cur.fetchall() or []
    return [_signed_output_row_to_item(dict(r)) for r in rows]


def list_signed_legacy_page(*, q: str = "", page: int = 1, page_size: int = 50):
    """历史记录（batch_id 为空）：分页、按文件名搜索。"""
    ensure_sign_mysql()
    q = (q or "").strip()
    page = max(1, int(page))
    page_size = max(1, min(int(page_size), 200))
    offset = (page - 1) * page_size
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            if q:
                like = f"%{q}%"
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_signed_output "
                    "WHERE (batch_id IS NULL OR batch_id='') AND (output_name LIKE %s OR source_name LIKE %s)",
                    (like, like),
                )
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output "
                    "WHERE (batch_id IS NULL OR batch_id='') AND (output_name LIKE %s OR source_name LIKE %s) "
                    "ORDER BY created_at DESC LIMIT %s OFFSET %s",
                    (like, like, page_size, offset),
                )
            else:
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_signed_output WHERE (batch_id IS NULL OR batch_id='')"
                )
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT id, batch_id, source_file_id, source_name, output_name, ext, roles_json, ftp_path, ftp_last_error, "
                    " (file_data IS NOT NULL AND LENGTH(file_data) > 0) AS has_blob, created_at "
                    "FROM sign_signed_output WHERE (batch_id IS NULL OR batch_id='') "
                    "ORDER BY created_at DESC LIMIT %s OFFSET %s",
                    (page_size, offset),
                )
            rows = cur.fetchall() or []
    return [_signed_output_row_to_item(dict(r)) for r in rows], total


def get_signed_row(signed_id: str) -> Optional[dict]:
    """返回含 file_data(bytes)；不存在则 None。优先从 FTP 取，兼容旧 BLOB。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, output_name AS name, ext, ftp_path, file_data "
                "FROM sign_signed_output WHERE id=%s",
                (signed_id,),
            )
            row = cur.fetchone()
    if not row:
        return None
    if row.get("file_data"):
        return row
    p = (row.get("ftp_path") or "").strip()
    if p:
        try:
            from ftp_store import download_bytes

            row["file_data"] = download_bytes(p)
        except Exception:
            pass
    return row


def delete_signed_output(signed_id: str) -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signed_output_table(conn)
        try:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT ftp_path FROM sign_signed_output WHERE id=%s",
                    (signed_id,),
                )
                row = cur.fetchone()
            p = (row or {}).get("ftp_path")
            if p:
                try:
                    from ftp_store import delete_path

                    delete_path(p)
                except Exception:
                    pass
        except Exception:
            pass
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_signed_output WHERE id=%s", (signed_id,))
            return cur.rowcount


def _emit_ftp_migration_abort(
    stats: Dict[str, Any],
    errs: List[dict],
    *,
    abort_cap: int,
    phase: str,
) -> None:
    """失败次数达到上限时立即输出样例，便于排查，不等到整段跑完。"""
    stats["aborted"] = True
    print(
        f"\n*** 已中止（{phase}）：失败次数达到 {abort_cap}，请先排查 FTP 连接、账号权限与系统设置中的 FTP_*。错误样例：",
        flush=True,
    )
    for e in errs[:5]:
        print(
            f"    [{e.get('table')}] id={e.get('id')!r} {e.get('error')}",
            flush=True,
        )


def migrate_mysql_blobs_to_ftp(
    *,
    batch_size: int = 2000,
    max_total: int = 200000,
    clear_blob: bool = True,
    error_samples: int = 20,
    progress_log: bool = False,
    abort_after_failures: Optional[int] = None,
) -> Dict[str, int]:
    """
    将历史遗留的 MySQL BLOB（sign_uploaded_file/sign_signed_output.file_data）迁移到 FTP。

    迁移条件：
    - ftp_path 为空
    - file_data 非空

    clear_blob=True 时，迁移成功后会将 file_data 置 NULL 以释放 MySQL 空间。
    返回统计：scanned / uploaded / skipped / failed。
    """
    ensure_sign_mysql()
    bs = max(1, min(int(batch_size or 0), 20000))
    mt = max(1, min(int(max_total or 0), 2000000))
    es = max(0, min(int(error_samples or 0), 200))
    abort_cap: Optional[int] = None
    if abort_after_failures is not None:
        af = int(abort_after_failures)
        if af > 0:
            abort_cap = af
    stats = {"scanned": 0, "uploaded": 0, "skipped": 0, "failed": 0, "done": 0}
    errs = []
    abort_early = False

    def _do_table(table: str, name_col: str, id_col: str, kind: str):
        nonlocal stats, errs, abort_early
        processed = 0
        while processed < mt:
            with _conn_commit() as conn:
                with conn.cursor() as cur:
                    cur.execute(
                        f"SELECT {id_col} AS id, {name_col} AS name, ext, ftp_path, file_data "
                        f"FROM {table} "
                        "WHERE (ftp_path IS NULL OR ftp_path='') AND file_data IS NOT NULL AND LENGTH(file_data) > 0 "
                        "ORDER BY created_at DESC LIMIT %s",
                        (bs,),
                    )
                    rows = cur.fetchall() or []
            if not rows:
                break
            if progress_log:
                print(
                    f"  [{table}] 本批 {len(rows)} 条，累计 uploaded={stats['uploaded']} "
                    f"failed={stats['failed']} scanned={stats['scanned']}",
                    flush=True,
                )
            for r in rows:
                if processed >= mt:
                    break
                processed += 1
                stats["scanned"] += 1
                fid = (r.get("id") or "").strip()
                nm = (r.get("name") or "document").strip()
                data = r.get("file_data") or b""
                if not fid or not data:
                    stats["skipped"] += 1
                    continue
                try:
                    sha = hashlib.sha256(data).hexdigest()
                    size = int(len(data))
                    from ftp_store import upload_bytes

                    safe_name = (nm or "document")[:200]
                    remote_rel = f"sign/{kind}/{fid}/{safe_name}"
                    ftp_path = upload_bytes(data, remote_rel)
                    with _conn_commit() as conn:
                        with conn.cursor() as cur:
                            if clear_blob:
                                cur.execute(
                                    f"UPDATE {table} SET ftp_path=%s, file_size=%s, sha256=%s, file_data=NULL WHERE {id_col}=%s",
                                    (ftp_path, size, sha, fid),
                                )
                            else:
                                cur.execute(
                                    f"UPDATE {table} SET ftp_path=%s, file_size=%s, sha256=%s WHERE {id_col}=%s",
                                    (ftp_path, size, sha, fid),
                                )
                    stats["uploaded"] += 1
                except Exception as e:
                    stats["failed"] += 1
                    if es and len(errs) < es:
                        errs.append({"table": table, "id": fid, "error": str(e)[:300]})
                    if abort_cap is not None and stats["failed"] >= abort_cap:
                        _emit_ftp_migration_abort(stats, errs, abort_cap=abort_cap, phase=table)
                        abort_early = True
                        break
            if abort_early:
                break

    _do_table("sign_uploaded_file", "original_name", "id", "inbox")
    if not abort_early:
        _do_table("sign_signed_output", "output_name", "id", "output")
    stats["done"] = stats["uploaded"] + stats["skipped"] + stats["failed"]
    if errs:
        stats["errors"] = errs
    return stats


def verify_and_backfill_ftp_files(
    *,
    limit: int = 2000,
    error_samples: int = 20,
    progress_log: bool = False,
    abort_after_failures: Optional[int] = None,
) -> Dict[str, int]:
    """
    校验/补传：
    - 若 ftp_path 已存在但 FTP 上文件缺失：尝试用 MySQL BLOB 重新上传并覆盖 ftp_path（保持不变）；
    - 若 ftp_path 为空但 BLOB 非空：上传并回写 ftp_path。

    返回统计：checked / backfilled / missing_blob / failed，以及 errors 样例。
    """
    ensure_sign_mysql()
    lim = max(1, min(int(limit or 0), 200000))
    es = max(0, min(int(error_samples or 0), 200))
    abort_cap: Optional[int] = None
    if abort_after_failures is not None:
        af = int(abort_after_failures)
        if af > 0:
            abort_cap = af
    stats = {"checked": 0, "backfilled": 0, "missing_blob": 0, "failed": 0}
    errs = []
    abort_early = False

    def _ftp_exists(ftp, path: str) -> bool:
        try:
            ftp.size(path)
            return True
        except Exception:
            return False

    def _check_table(table: str, name_col: str, id_col: str, kind: str):
        nonlocal stats, errs, abort_early
        if progress_log:
            print(f"[verify] 检查表 {table}（最多 {lim} 行）…", flush=True)
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    f"SELECT {id_col} AS id, {name_col} AS name, ext, ftp_path, file_data "
                    f"FROM {table} "
                    "ORDER BY created_at DESC LIMIT %s",
                    (lim,),
                )
                rows = cur.fetchall() or []
        from ftp_store import _ftp as _ftp_ctx, upload_bytes

        with _ftp_ctx() as ftp:
            for r in rows:
                stats["checked"] += 1
                fid = (r.get("id") or "").strip()
                nm = (r.get("name") or "document").strip()
                ftp_path = (r.get("ftp_path") or "").strip()
                blob = r.get("file_data") or b""
                if not fid:
                    continue
                try:
                    if ftp_path:
                        if _ftp_exists(ftp, ftp_path):
                            continue
                        # ftp missing: try re-upload from blob
                        if not blob:
                            stats["missing_blob"] += 1
                            continue
                        # upload to same logical place (may differ from old ftp_path conventions)
                        safe_name = (nm or "document")[:200]
                        remote_rel = f"sign/{kind}/{fid}/{safe_name}"
                        new_path = upload_bytes(blob, remote_rel)
                        with _conn_commit() as conn:
                            with conn.cursor() as cur:
                                cur.execute(
                                    f"UPDATE {table} SET ftp_path=%s, file_size=%s, sha256=%s WHERE {id_col}=%s",
                                    (new_path, int(len(blob)), hashlib.sha256(blob).hexdigest(), fid),
                                )
                        stats["backfilled"] += 1
                        continue
                    # no ftp_path: upload if blob exists
                    if blob:
                        safe_name = (nm or "document")[:200]
                        remote_rel = f"sign/{kind}/{fid}/{safe_name}"
                        new_path = upload_bytes(blob, remote_rel)
                        with _conn_commit() as conn:
                            with conn.cursor() as cur:
                                cur.execute(
                                    f"UPDATE {table} SET ftp_path=%s, file_size=%s, sha256=%s WHERE {id_col}=%s",
                                    (new_path, int(len(blob)), hashlib.sha256(blob).hexdigest(), fid),
                                )
                        stats["backfilled"] += 1
                except Exception as e:
                    stats["failed"] += 1
                    if es and len(errs) < es:
                        errs.append({"table": table, "id": fid, "error": str(e)[:300]})
                    if abort_cap is not None and stats["failed"] >= abort_cap:
                        _emit_ftp_migration_abort(
                            stats, errs, abort_cap=abort_cap, phase=f"verify {table}"
                        )
                        abort_early = True
                        break

    def _check_stroke_items():
        nonlocal stats, errs, abort_early
        if progress_log:
            print(f"[verify] 检查表 sign_stroke_item（最多 {lim} 行）…", flush=True)
        with _conn_commit() as conn:
            _ensure_signer_tables(conn)
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, signer_id, locale, kind, ftp_path, png "
                    "FROM sign_stroke_item ORDER BY updated_at DESC LIMIT %s",
                    (lim,),
                )
                rows = cur.fetchall() or []
        from ftp_store import _ftp as _ftp_ctx, upload_bytes

        with _ftp_ctx() as ftp:
            for r in rows:
                stats["checked"] += 1
                item_id = (r.get("id") or "").strip()
                ftp_path = (r.get("ftp_path") or "").strip()
                blob = r.get("png") or b""
                if not item_id:
                    continue
                try:
                    target_rel = f"sign/strokeitem/{item_id}.png"
                    if ftp_path:
                        if _ftp_exists(ftp, ftp_path):
                            continue
                        if not blob:
                            stats["missing_blob"] += 1
                            continue
                        new_path = upload_bytes(blob, target_rel)
                        with _conn_commit() as conn:
                            with conn.cursor() as cur:
                                cur.execute(
                                    "UPDATE sign_stroke_item SET ftp_path=%s, file_size=%s WHERE id=%s",
                                    (new_path, int(len(blob)), item_id),
                                )
                        stats["backfilled"] += 1
                        continue
                    if blob:
                        new_path = upload_bytes(blob, target_rel)
                        with _conn_commit() as conn:
                            with conn.cursor() as cur:
                                cur.execute(
                                    "UPDATE sign_stroke_item SET ftp_path=%s, file_size=%s WHERE id=%s",
                                    (new_path, int(len(blob)), item_id),
                                )
                        stats["backfilled"] += 1
                except Exception as e:
                    stats["failed"] += 1
                    if es and len(errs) < es:
                        errs.append({"table": "sign_stroke_item", "id": item_id, "error": str(e)[:300]})
                    if abort_cap is not None and stats["failed"] >= abort_cap:
                        _emit_ftp_migration_abort(
                            stats, errs, abort_cap=abort_cap, phase="verify sign_stroke_item"
                        )
                        abort_early = True
                        break

    _check_table("sign_uploaded_file", "original_name", "id", "inbox")
    if abort_early:
        if errs:
            stats["errors"] = errs
        return stats
    _check_table("sign_signed_output", "output_name", "id", "output")
    if abort_early:
        if errs:
            stats["errors"] = errs
        return stats
    _check_stroke_items()
    if errs:
        stats["errors"] = errs
    return stats


def migrate_stroke_items_blobs_to_ftp(
    *,
    batch_size: int = 2000,
    max_total: int = 200000,
    clear_blob: bool = True,
    error_samples: int = 20,
    progress_log: bool = False,
    abort_after_failures: Optional[int] = None,
) -> Dict[str, int]:
    """
    将 sign_stroke_item 里的 png BLOB 迁移到 FTP。
    条件：ftp_path 为空且 png 非空。
    """
    ensure_sign_mysql()
    bs = max(1, min(int(batch_size or 0), 20000))
    mt = max(1, min(int(max_total or 0), 2000000))
    es = max(0, min(int(error_samples or 0), 200))
    abort_cap: Optional[int] = None
    if abort_after_failures is not None:
        af = int(abort_after_failures)
        if af > 0:
            abort_cap = af
    stats = {"scanned": 0, "uploaded": 0, "skipped": 0, "failed": 0, "done": 0}
    errs = []
    processed = 0
    while processed < mt:
        with _conn_commit() as conn:
            _ensure_signer_tables(conn)
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, png FROM sign_stroke_item "
                    "WHERE (ftp_path IS NULL OR ftp_path='') AND png IS NOT NULL AND LENGTH(png) > 0 "
                    "ORDER BY updated_at DESC LIMIT %s",
                    (bs,),
                )
                rows = cur.fetchall() or []
        if not rows:
            break
        if progress_log:
            print(
                f"  [sign_stroke_item] 本批 {len(rows)} 条，累计 uploaded={stats['uploaded']} "
                f"failed={stats['failed']} scanned={stats['scanned']}",
                flush=True,
            )
        for r in rows:
            if processed >= mt:
                break
            processed += 1
            stats["scanned"] += 1
            item_id = (r.get("id") or "").strip()
            png_b = r.get("png") or b""
            if not item_id or not png_b:
                stats["skipped"] += 1
                continue
            try:
                from ftp_store import upload_bytes

                ftp_path = upload_bytes(png_b, f"sign/strokeitem/{item_id}.png")
                with _conn_commit() as conn:
                    with conn.cursor() as cur:
                        if clear_blob:
                            cur.execute(
                                "UPDATE sign_stroke_item SET ftp_path=%s, file_size=%s, png=NULL WHERE id=%s",
                                (ftp_path, int(len(png_b)), item_id),
                            )
                        else:
                            cur.execute(
                                "UPDATE sign_stroke_item SET ftp_path=%s, file_size=%s WHERE id=%s",
                                (ftp_path, int(len(png_b)), item_id),
                            )
                stats["uploaded"] += 1
            except Exception as e:
                stats["failed"] += 1
                if es and len(errs) < es:
                    errs.append({"table": "sign_stroke_item", "id": item_id, "error": str(e)[:300]})
                if abort_cap is not None and stats["failed"] >= abort_cap:
                    _emit_ftp_migration_abort(
                        stats, errs, abort_cap=abort_cap, phase="sign_stroke_item"
                    )
                    break
        if stats.get("aborted"):
            break
    stats["done"] = stats["uploaded"] + stats["skipped"] + stats["failed"]
    if errs:
        stats["errors"] = errs
    return stats


def migrate_signer_strokes_blobs_to_ftp(
    *,
    batch_size: int = 2000,
    max_total: int = 200000,
    clear_blob: bool = True,
    error_samples: int = 20,
    progress_log: bool = False,
    abort_after_failures: Optional[int] = None,
) -> Dict[str, int]:
    """
    将 sign_signer_stroke 里的 sig_png/date_png 迁移到 FTP（可复用素材）。
    条件：对应 *_ftp_path 为空且对应 PNG BLOB 非空。
    """
    ensure_sign_mysql()
    bs = max(1, min(int(batch_size or 0), 20000))
    mt = max(1, min(int(max_total or 0), 2000000))
    es = max(0, min(int(error_samples or 0), 200))
    abort_cap: Optional[int] = None
    if abort_after_failures is not None:
        af = int(abort_after_failures)
        if af > 0:
            abort_cap = af
    stats = {"scanned": 0, "uploaded": 0, "skipped": 0, "failed": 0, "done": 0}
    errs = []
    processed = 0
    while processed < mt:
        with _conn_commit() as conn:
            _ensure_signer_tables(conn)
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT signer_id, sig_ftp_path, date_ftp_path, sig_png, date_png "
                    "FROM sign_signer_stroke "
                    "WHERE "
                    " ( (sig_ftp_path IS NULL OR sig_ftp_path='') AND sig_png IS NOT NULL AND LENGTH(sig_png) > 0 ) "
                    " OR "
                    " ( (date_ftp_path IS NULL OR date_ftp_path='') AND date_png IS NOT NULL AND LENGTH(date_png) > 0 ) "
                    "ORDER BY updated_at DESC LIMIT %s",
                    (bs,),
                )
                rows = cur.fetchall() or []
        if not rows:
            break
        if progress_log:
            print(
                f"  [sign_signer_stroke] 本批 {len(rows)} 条，累计 uploaded={stats['uploaded']} "
                f"failed={stats['failed']} scanned={stats['scanned']}",
                flush=True,
            )
        for r in rows:
            if processed >= mt:
                break
            processed += 1
            stats["scanned"] += 1
            sid = (r.get("signer_id") or "").strip()
            if not sid:
                stats["skipped"] += 1
                continue
            try:
                from ftp_store import upload_bytes

                sig_path = (r.get("sig_ftp_path") or "").strip()
                date_path = (r.get("date_ftp_path") or "").strip()
                sig_png = r.get("sig_png") or b""
                date_png = r.get("date_png") or b""
                upd = {}
                if (not sig_path) and sig_png:
                    upd["sig_ftp_path"] = upload_bytes(sig_png, f"sign/strokes/{sid}/sig.png")
                    upd["sig_size"] = int(len(sig_png))
                    upd["sig_sha256"] = hashlib.sha256(sig_png).hexdigest()
                    if clear_blob:
                        upd["sig_png"] = None
                if (not date_path) and date_png:
                    upd["date_ftp_path"] = upload_bytes(date_png, f"sign/strokes/{sid}/date.png")
                    upd["date_size"] = int(len(date_png))
                    upd["date_sha256"] = hashlib.sha256(date_png).hexdigest()
                    if clear_blob:
                        upd["date_png"] = None
                if not upd:
                    stats["skipped"] += 1
                    continue
                cols = []
                vals = []
                for k, v in upd.items():
                    cols.append(f"{k}=%s")
                    vals.append(v)
                vals.append(sid)
                with _conn_commit() as conn:
                    _ensure_signer_tables(conn)
                    with conn.cursor() as cur:
                        cur.execute(
                            "UPDATE sign_signer_stroke SET " + ", ".join(cols) + " WHERE signer_id=%s",
                            tuple(vals),
                        )
                stats["uploaded"] += 1
            except Exception as e:
                stats["failed"] += 1
                if es and len(errs) < es:
                    errs.append({"table": "sign_signer_stroke", "id": sid, "error": str(e)[:300]})
                if abort_cap is not None and stats["failed"] >= abort_cap:
                    _emit_ftp_migration_abort(
                        stats, errs, abort_cap=abort_cap, phase="sign_signer_stroke"
                    )
                    break
        if stats.get("aborted"):
            break
    stats["done"] = stats["uploaded"] + stats["skipped"] + stats["failed"]
    if errs:
        stats["errors"] = errs
    return stats


def _ensure_signer_tables(conn) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_signer (
                id VARCHAR(32) NOT NULL PRIMARY KEY,
                display_name VARCHAR(128) NOT NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                KEY idx_signer_created (created_at)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_signer_stroke (
                signer_id VARCHAR(32) NOT NULL PRIMARY KEY,
                sig_ftp_path VARCHAR(768) NULL,
                date_ftp_path VARCHAR(768) NULL,
                sig_size BIGINT NULL,
                date_size BIGINT NULL,
                sig_sha256 CHAR(64) NULL,
                date_sha256 CHAR(64) NULL,
                sig_png LONGBLOB NULL,
                date_png LONGBLOB NULL,
                updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_file_role_signer (
                file_id VARCHAR(32) NOT NULL,
                role_id VARCHAR(64) NOT NULL,
                signer_id VARCHAR(32) NOT NULL,
                PRIMARY KEY (file_id, role_id),
                KEY idx_frs_file (file_id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_stroke_set (
                id VARCHAR(32) NOT NULL PRIMARY KEY,
                signer_id VARCHAR(32) NOT NULL,
                locale VARCHAR(8) NOT NULL DEFAULT 'zh',
                sig_sha256 CHAR(64) NOT NULL,
                date_sha256 CHAR(64) NOT NULL,
                sig_ftp_path VARCHAR(768) NULL,
                date_ftp_path VARCHAR(768) NULL,
                sig_size BIGINT NULL,
                date_size BIGINT NULL,
                sig_png LONGBLOB NULL,
                date_png LONGBLOB NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uk_stroke_sig_date (signer_id, locale, sig_sha256, date_sha256),
                KEY idx_ss_signer (signer_id)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS sign_stroke_item (
                id VARCHAR(32) NOT NULL PRIMARY KEY,
                signer_id VARCHAR(32) NOT NULL,
                locale VARCHAR(8) NOT NULL DEFAULT 'zh',
                kind VARCHAR(8) NOT NULL,
                sha256 CHAR(64) NOT NULL,
                ftp_path VARCHAR(768) NULL,
                file_size BIGINT NULL,
                png LONGBLOB NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                UNIQUE KEY uk_item (signer_id, locale, kind, sha256),
                KEY idx_item_signer (signer_id),
                KEY idx_item_kind (kind)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            """
        )

    # 兼容旧表：补齐列，并允许 BLOB 为空
    try:
        with conn.cursor() as cur:
            for sql in (
                "ALTER TABLE sign_signer_stroke ADD COLUMN sig_ftp_path VARCHAR(768) NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN date_ftp_path VARCHAR(768) NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN sig_size BIGINT NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN date_size BIGINT NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN sig_sha256 CHAR(64) NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN date_sha256 CHAR(64) NULL",
            ):
                try:
                    cur.execute(sql)
                except Exception:
                    pass
            try:
                cur.execute("ALTER TABLE sign_signer_stroke MODIFY COLUMN sig_png LONGBLOB NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_signer_stroke MODIFY COLUMN date_png LONGBLOB NULL")
            except Exception:
                pass
            try:
                cur.execute(
                    "ALTER TABLE sign_file_role_signer ADD COLUMN stroke_set_id VARCHAR(32) NULL"
                )
            except Exception:
                pass
            try:
                cur.execute(
                    "ALTER TABLE sign_file_role_signer ADD KEY idx_frs_stroke_set (stroke_set_id)"
                )
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_file_role_signer ADD COLUMN sig_item_id VARCHAR(32) NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_file_role_signer ADD COLUMN date_item_id VARCHAR(32) NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_file_role_signer ADD KEY idx_frs_sig_item (sig_item_id)")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_file_role_signer ADD KEY idx_frs_date_item (date_item_id)")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_set ADD COLUMN locale VARCHAR(8) NOT NULL DEFAULT 'zh'")
            except Exception:
                pass
            # 尝试升级唯一键：旧库可能没有 locale 维度
            try:
                cur.execute("ALTER TABLE sign_stroke_set DROP INDEX uk_stroke_sig_date")
            except Exception:
                pass
            try:
                cur.execute(
                    "ALTER TABLE sign_stroke_set ADD UNIQUE KEY uk_stroke_sig_date (signer_id, locale, sig_sha256, date_sha256)"
                )
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN locale VARCHAR(8) NOT NULL DEFAULT 'zh'")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN kind VARCHAR(8) NOT NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN sha256 CHAR(64) NOT NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN ftp_path VARCHAR(768) NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN file_size BIGINT NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD COLUMN png LONGBLOB NULL")
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item ADD UNIQUE KEY uk_item (signer_id, locale, kind, sha256)")
            except Exception:
                pass
            for sql in (
                "ALTER TABLE sign_stroke_item ADD COLUMN ftp_last_error VARCHAR(512) NULL",
                "ALTER TABLE sign_signer_stroke ADD COLUMN ftp_last_error VARCHAR(512) NULL",
                "ALTER TABLE sign_stroke_set ADD COLUMN ftp_last_error VARCHAR(512) NULL",
            ):
                try:
                    cur.execute(sql)
                except Exception:
                    pass
            try:
                cur.execute("ALTER TABLE sign_stroke_item MODIFY COLUMN kind VARCHAR(32) NOT NULL")
            except Exception:
                pass
            try:
                cur.execute(
                    "ALTER TABLE sign_file_role_signer ADD COLUMN date_mode VARCHAR(24) NULL"
                )
            except Exception:
                pass
            try:
                cur.execute("ALTER TABLE sign_file_role_signer ADD COLUMN date_iso VARCHAR(32) NULL")
            except Exception:
                pass
    except Exception:
        pass
    try:
        with conn.cursor() as cur:
            cur.execute(
                """
                ALTER TABLE sign_stroke_set
                ADD CONSTRAINT fk_ss_signer FOREIGN KEY (signer_id)
                REFERENCES sign_signer (id) ON DELETE CASCADE
                """
            )
    except Exception:
        pass
    try:
        _stroke_set_migrate_legacy_and_roles(conn)
    except Exception:
        pass
    try:
        _stroke_item_migrate_from_sets(conn)
    except Exception:
        pass


def _stroke_item_migrate_from_sets(conn) -> None:
    """将 sign_stroke_set 中的签名/日期拆成 sign_stroke_item（幂等，默认 locale=zh/en 按行 locale）。"""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id, signer_id, locale, sig_ftp_path, date_ftp_path, sig_png, date_png FROM sign_stroke_set"
        )
        rows = cur.fetchall() or []
    for r in rows:
        row = _hydrate_stroke_storage_row(dict(r))
        signer_id = row.get("signer_id")
        loc = (row.get("locale") or "zh").strip().lower()
        if loc not in ("zh", "en"):
            loc = "zh"
        sig_b = row.get("sig_png") or b""
        date_b = row.get("date_png") or b""
        if sig_b:
            _upsert_stroke_item_core(conn, signer_id, loc, "sig", sig_b, prefer_ftp_path=row.get("sig_ftp_path"))
        if date_b:
            _upsert_stroke_item_core(conn, signer_id, loc, "date", date_b, prefer_ftp_path=row.get("date_ftp_path"))


def _upsert_stroke_item_core(
    conn,
    signer_id: str,
    locale: str,
    kind: str,
    png_b: bytes,
    prefer_ftp_path: Optional[str] = None,
) -> Dict[str, Any]:
    raw = (kind or "").strip().lower()
    pk = normalize_piece_kind(raw)
    piece_mode = bool(pk)
    if piece_mode:
        k = pk or ""
        loc = "en"
        overwrote = False
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM sign_stroke_item WHERE signer_id=%s AND locale=%s AND kind=%s",
                (signer_id, loc, k),
            )
            overwrote = bool(cur.fetchone())
            cur.execute(
                "DELETE FROM sign_stroke_item WHERE signer_id=%s AND locale=%s AND kind=%s",
                (signer_id, loc, k),
            )
        item_id = uuid.uuid4().hex
    else:
        k = raw
        if k not in ("sig", "date"):
            raise ValueError("kind 须为 sig、date 或日期元件（pd0..pd9 / pm01..pm12 / pdot）")
        loc = (locale or "zh").strip().lower()
        if loc not in ("zh", "en"):
            loc = "zh"
        sha = hashlib.sha256(png_b).hexdigest()
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM sign_stroke_item WHERE signer_id=%s AND locale=%s AND kind=%s AND sha256=%s",
                (signer_id, loc, k, sha),
            )
            ex = cur.fetchone()
            overwrote = bool(ex)
            item_id = ex["id"] if ex else uuid.uuid4().hex

    sha = hashlib.sha256(png_b).hexdigest()
    ftp_path = None
    ftp_err_note: Optional[str] = None
    if prefer_ftp_path:
        ftp_path = prefer_ftp_path
    else:
        ftp_path, ftp_err_note = _ftp_upload_bytes_or_mysql(
            png_b, f"sign/strokeitem/{item_id}.png"
        )
    err_s = ((ftp_err_note or "")[:512]) if ftp_err_note else None

    png_store = None if ftp_path else png_b
    with conn.cursor() as cur:
        if not piece_mode and overwrote:
            cur.execute(
                "UPDATE sign_stroke_item SET ftp_path=%s, file_size=%s, png=%s, ftp_last_error=%s, sha256=%s WHERE id=%s",
                (ftp_path, int(len(png_b)), png_store, None if ftp_path else err_s, sha, item_id),
            )
        else:
            cur.execute(
                "INSERT INTO sign_stroke_item (id, signer_id, locale, kind, sha256, ftp_path, file_size, png, ftp_last_error) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (
                    item_id,
                    signer_id,
                    loc,
                    k,
                    sha,
                    ftp_path,
                    int(len(png_b)),
                    png_store,
                    None if ftp_path else err_s,
                ),
            )
    return {"stroke_item_id": item_id, "overwritten": overwrote, "sha256": sha, "kind": k, "locale": loc}


def upsert_signer_stroke_piece(signer_id: str, piece_kind: str, png_b: bytes) -> Dict[str, Any]:
    """录入英文点分日期笔迹元件（locale 固定为 en，同槽位覆盖）。"""
    ensure_sign_mysql()
    if not normalize_piece_kind(piece_kind):
        raise ValueError("无效的 piece_kind（pd0..pd9 / pm01..pm12 / pma01..pma12 / pdot）")
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM sign_signer WHERE id=%s", (signer_id,))
            if not cur.fetchone():
                raise ValueError("签署人不存在")
        return _upsert_stroke_item_core(conn, signer_id, "en", piece_kind, png_b)


def get_stroke_item_row_by_signer_kind(signer_id: str, locale: str, kind: str) -> Optional[dict]:
    ensure_sign_mysql()
    loc = (locale or "zh").strip().lower()
    k = (kind or "").strip().lower()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, signer_id, locale, kind, sha256, ftp_path, ftp_last_error, png FROM sign_stroke_item "
                "WHERE signer_id=%s AND locale=%s AND kind=%s ORDER BY updated_at DESC LIMIT 1",
                (signer_id, loc, k),
            )
            row = cur.fetchone()
    if not row:
        return None
    row = dict(row)
    try:
        from ftp_store import download_bytes

        if (not row.get("png")) and row.get("ftp_path"):
            row["png"] = download_bytes(row["ftp_path"])
    except Exception:
        pass
    return row


def get_piece_png_for_date_compose(signer_id: str, slot: str) -> Optional[bytes]:
    """拼接用：先取该槽位；若为月份整词 pmXX 且无图，则回退到简称 pmaXX。"""
    sk = (slot or "").strip().lower()
    row = get_stroke_item_row_by_signer_kind(signer_id, "en", sk)
    b = (row or {}).get("png")
    if b:
        return b
    if re.fullmatch(r"pm(0[1-9]|1[0-2])", sk):
        alt = "pma" + sk[2:]
        row2 = get_stroke_item_row_by_signer_kind(signer_id, "en", alt)
        return (row2 or {}).get("png")
    return None


def compose_date_piece_png(signer_id: str, iso: str, layout: str) -> Tuple[bytes, str]:
    """
    按 YYYY-MM-DD 与版式生成横向拼接 PNG（locale=en 的 pd*/pm*/pdot 笔迹元件）。
    layout: zh_ymd → 2026.04.15；en_space → 15 April 2026；en_dot → 15.April.2026（兼容旧布局）。
    """
    lay = (layout or "en_dot").strip().lower()
    gaps: Optional[List[int]] = None
    if lay == "zh_ymd":
        kinds, label = kinds_zh_ymd_dot(iso)
    elif lay == "en_space":
        kinds, label, gaps = kinds_en_dmy_space(iso)
    elif lay in ("en_dot", "dot_dmy"):
        kinds, label = kinds_en_dot_dmy(iso)
    else:
        kinds, label = kinds_for_iso_date(iso)
    pngs: List[bytes] = []
    missing: List[str] = []
    for slot in kinds:
        b = get_piece_png_for_date_compose(signer_id, slot)
        if not b:
            sk = (slot or "").strip().lower()
            try:
                from sign_handlers.date_piece_compose import piece_kind_label

                human = piece_kind_label(sk)
            except Exception:
                human = sk
            # 兼容：若仍有人录入了整词 pmXX，缺失提示里也说明可用简称
            if re.fullmatch(r"pm(0[1-9]|1[0-2])", sk):
                missing.append(f"{human}（{sk}；也可录入简称 pma{sk[2:]}）")
            else:
                missing.append(f"{human}（{sk}）" if human and human != sk else sk)
        else:
            if lay == "zh_ymd" and (slot or "").strip().lower() == "pdot":
                # 中文点分日期：句点单独占一格，但点贴在格子的右下角（视觉上靠近前一位数字）
                try:
                    from sign_handlers.date_piece_compose import render_dot_cell_right_bottom

                    pngs.append(render_dot_cell_right_bottom(b, target_h=360, cell_w=110, dot_scale=0.28, margin=2))
                except Exception:
                    pngs.append(b)
            else:
                pngs.append(b)
    if missing:
        lay_h = "中文 2026.04.15" if lay == "zh_ymd" else ("英文 15 Apr 2026" if lay == "en_space" else lay)
        raise ValueError(
            "无法拼接预览：" + lay_h + " 所需的笔迹元件尚未录入。缺少："
            + "，".join(missing)
            + "。请先到「英文点分日期笔迹元件」为该签署人录入对应元件后再预览。"
        )
    # 默认字间距收紧（英文空格版可通过 gaps 单独放大“空格”）
    out = compose_png_horizontal(pngs, gap=3, gaps=gaps, target_h=360)
    return out, label


def compose_en_dot_date_png(signer_id: str, iso: str) -> Tuple[bytes, str]:
    """兼容：等同 en_dot 版式。"""
    return compose_date_piece_png(signer_id, iso, "en_dot")


def get_stroke_item_row(item_id: str) -> Optional[dict]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, signer_id, locale, kind, sha256, ftp_path, ftp_last_error, png FROM sign_stroke_item WHERE id=%s",
                (item_id,),
            )
            row = cur.fetchone()
    if not row:
        return None
    row = dict(row)
    try:
        from ftp_store import download_bytes

        if (not row.get("png")) and row.get("ftp_path"):
            row["png"] = download_bytes(row["ftp_path"])
    except Exception:
        pass
    return row


def list_stroke_items_page(
    *,
    q: str = "",
    page: int = 1,
    page_size: int = 10,
    cat: str = "",
) -> Tuple[List[dict], int]:
    """已入库的签字 PNG 素材（sign_stroke_item），按签署人显示名或 ID 模糊搜索，分页。"""
    ensure_sign_mysql()
    q = (q or "").strip()
    cat = (cat or "").strip().lower()
    page = max(1, int(page))
    page_size = max(1, min(int(page_size), 500))
    offset = (page - 1) * page_size
    where_cat = ""
    cat_args: List[Any] = []
    if cat in ("sig", "signature"):
        where_cat = " AND i.kind='sig' "
    elif cat in ("digit", "digits"):
        where_cat = " AND i.locale='en' AND i.kind REGEXP '^(pd[0-9])$' "
    elif cat in ("en_date", "en", "month", "months"):
        # 英文日期：月份简称元件（pma01..pma12）；兼容旧整词 pm01..pm12
        where_cat = " AND i.locale='en' AND i.kind REGEXP '^(pma(0[1-9]|1[0-2])|pm(0[1-9]|1[0-2]))$' "
    elif cat in ("connector", "dot", "pdot"):
        where_cat = " AND i.locale='en' AND i.kind='pdot' "
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            if q:
                like = f"%{q}%"
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_stroke_item i "
                    "INNER JOIN sign_signer s ON s.id = i.signer_id "
                    "WHERE (s.display_name LIKE %s OR i.signer_id LIKE %s) "
                    + where_cat,
                    tuple([like, like] + cat_args),
                )
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT i.id, i.signer_id, i.locale, i.kind, i.sha256, i.ftp_path, i.ftp_last_error, i.file_size, "
                    " (i.png IS NOT NULL AND LENGTH(i.png) > 0) AS has_blob, i.updated_at, "
                    " s.display_name AS signer_name "
                    "FROM sign_stroke_item i "
                    "INNER JOIN sign_signer s ON s.id = i.signer_id "
                    "WHERE (s.display_name LIKE %s OR i.signer_id LIKE %s) "
                    + where_cat +
                    "ORDER BY i.updated_at DESC LIMIT %s OFFSET %s",
                    tuple([like, like] + cat_args + [page_size, offset]),
                )
            else:
                cur.execute(
                    "SELECT COUNT(*) AS c FROM sign_stroke_item i "
                    "INNER JOIN sign_signer s ON s.id = i.signer_id WHERE 1=1 "
                    + where_cat,
                    tuple(cat_args),
                )
                total = int((cur.fetchone() or {}).get("c") or 0)
                cur.execute(
                    "SELECT i.id, i.signer_id, i.locale, i.kind, i.sha256, i.ftp_path, i.ftp_last_error, i.file_size, "
                    " (i.png IS NOT NULL AND LENGTH(i.png) > 0) AS has_blob, i.updated_at, "
                    " s.display_name AS signer_name "
                    "FROM sign_stroke_item i "
                    "INNER JOIN sign_signer s ON s.id = i.signer_id "
                    "WHERE 1=1 " + where_cat +
                    "ORDER BY i.updated_at DESC LIMIT %s OFFSET %s",
                    tuple(cat_args + [page_size, offset]),
                )
            rows = cur.fetchall() or []
    out: List[dict] = []
    for r in rows:
        r = dict(r)
        ts = r.get("updated_at")
        k = (r.get("kind") or "").strip().lower()
        fe = (r.get("ftp_last_error") or "").strip()
        if normalize_piece_kind(k):
            kl = piece_kind_label(k)
        else:
            kl = "日期" if k == "date" else "签名"
        out.append(
            {
                "id": r["id"],
                "signer_id": r.get("signer_id"),
                "signer_name": r.get("signer_name") or "",
                "locale": r.get("locale") or "zh",
                "kind": k,
                "kind_label": kl,
                "sha256": r.get("sha256"),
                "ftp_uploaded": bool((r.get("ftp_path") or "").strip()),
                "blob_stored": bool(r.get("has_blob")),
                "ftp_last_error": fe or None,
                "updated_at": ts.isoformat(sep=" ") if ts is not None else None,
            }
        )
    return out, total


def delete_stroke_item(item_id: str) -> int:
    """删除一条 sign_stroke_item，并解除角色映射中的引用。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT ftp_path FROM sign_stroke_item WHERE id=%s", (item_id,))
                row = cur.fetchone()
            p = (row or {}).get("ftp_path")
            if p:
                try:
                    from ftp_store import delete_path

                    delete_path(p)
                except Exception:
                    pass
        except Exception:
            pass
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE sign_file_role_signer SET sig_item_id=NULL WHERE sig_item_id=%s",
                (item_id,),
            )
            cur.execute(
                "UPDATE sign_file_role_signer SET date_item_id=NULL WHERE date_item_id=%s",
                (item_id,),
            )
            cur.execute("DELETE FROM sign_stroke_item WHERE id=%s", (item_id,))
            return int(cur.rowcount or 0)


def upsert_signer_stroke_item(signer_id: str, kind: str, png_b: bytes, locale: str = "zh") -> Dict[str, Any]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM sign_signer WHERE id=%s", (signer_id,))
            if not cur.fetchone():
                raise ValueError("签署人不存在")
        res = _upsert_stroke_item_core(conn, signer_id, locale, kind, png_b)
    return res


def _hydrate_stroke_storage_row(row: Optional[dict]) -> Optional[dict]:
    if not row:
        return None
    try:
        from ftp_store import download_bytes

        if (not row.get("sig_png")) and row.get("sig_ftp_path"):
            row["sig_png"] = download_bytes(row["sig_ftp_path"])
        if (not row.get("date_png")) and row.get("date_ftp_path"):
            row["date_png"] = download_bytes(row["date_ftp_path"])
    except Exception:
        pass
    return row


def _merge_signer_stroke_bytes(
    conn, signer_id: str, sig_png: Optional[bytes], date_png: Optional[bytes]
) -> Tuple[Optional[bytes], Optional[bytes]]:
    old: Optional[dict] = None
    with conn.cursor() as cur:
        cur.execute(
            "SELECT sig_ftp_path, date_ftp_path, sig_png, date_png FROM sign_signer_stroke WHERE signer_id=%s",
            (signer_id,),
        )
        r = cur.fetchone()
        if r:
            old = _hydrate_stroke_storage_row(dict(r))
    o_sig = (old or {}).get("sig_png")
    o_date = (old or {}).get("date_png")
    need_fill = (sig_png is None and not o_sig) or (date_png is None and not o_date)
    if need_fill:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT sig_ftp_path, date_ftp_path, sig_png, date_png
                FROM sign_stroke_set WHERE signer_id=%s ORDER BY updated_at DESC LIMIT 1
                """,
                (signer_id,),
            )
            ss = cur.fetchone()
        ss = _hydrate_stroke_storage_row(dict(ss)) if ss else None
        if ss:
            if not o_sig and ss.get("sig_png"):
                o_sig = ss["sig_png"]
            if not o_date and ss.get("date_png"):
                o_date = ss["date_png"]
    sig_b = sig_png if sig_png is not None else o_sig
    date_b = date_png if date_png is not None else o_date
    return sig_b, date_b


def _stroke_set_migrate_legacy_and_roles(conn) -> None:
    """将 sign_signer_stroke 迁入 sign_stroke_set，并回填 sign_file_role_signer.stroke_set_id（幂等）。"""
    with conn.cursor() as cur:
        cur.execute(
            "SELECT signer_id, sig_ftp_path, date_ftp_path, sig_png, date_png, sig_sha256, date_sha256 "
            "FROM sign_signer_stroke"
        )
        legacy_rows = cur.fetchall() or []
    for lr in legacy_rows:
        row = _hydrate_stroke_storage_row(dict(lr))
        if not row:
            continue
        sig_b = row.get("sig_png") or b""
        date_b = row.get("date_png") or b""
        if not sig_b or not date_b:
            continue
        sig_sha = (row.get("sig_sha256") or "").strip() or hashlib.sha256(sig_b).hexdigest()
        date_sha = (row.get("date_sha256") or "").strip() or hashlib.sha256(date_b).hexdigest()
        signer_id = row["signer_id"]
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM sign_stroke_set WHERE signer_id=%s AND locale=%s AND sig_sha256=%s AND date_sha256=%s",
                (signer_id, "zh", sig_sha, date_sha),
            )
            if cur.fetchone():
                continue
            set_id = uuid.uuid4().hex
            sig_path = row.get("sig_ftp_path")
            date_path = row.get("date_ftp_path")
            cur.execute(
                "INSERT INTO sign_stroke_set "
                "(id, signer_id, locale, sig_sha256, date_sha256, sig_ftp_path, date_ftp_path, "
                "sig_size, date_size, sig_png, date_png) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (
                    set_id,
                    signer_id,
                    "zh",
                    sig_sha,
                    date_sha,
                    sig_path,
                    date_path,
                    int(len(sig_b)),
                    int(len(date_b)),
                    None if sig_path else sig_b,
                    None if date_path else date_b,
                ),
            )
    with conn.cursor() as cur:
        cur.execute(
            "SELECT file_id, role_id, signer_id, stroke_set_id FROM sign_file_role_signer "
            "WHERE stroke_set_id IS NULL OR stroke_set_id=''"
        )
        fr_rows = cur.fetchall() or []
    for fr in fr_rows:
        sid = fr.get("signer_id")
        if not sid:
            continue
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id FROM sign_stroke_set WHERE signer_id=%s ORDER BY updated_at DESC LIMIT 1",
                (sid,),
            )
            ss = cur.fetchone()
            if not ss:
                continue
            cur.execute(
                "UPDATE sign_file_role_signer SET stroke_set_id=%s WHERE file_id=%s AND role_id=%s",
                (ss["id"], fr["file_id"], fr["role_id"]),
            )


def _resolve_map_val_to_stroke_set_id(conn, val: str) -> Optional[str]:
    v = (val or "").strip()
    if not v:
        return None
    with conn.cursor() as cur:
        cur.execute("SELECT id FROM sign_stroke_set WHERE id=%s", (v,))
        if cur.fetchone():
            return v
        cur.execute(
            "SELECT id FROM sign_stroke_set WHERE signer_id=%s ORDER BY updated_at DESC LIMIT 1",
            (v,),
        )
        r = cur.fetchone()
        if r:
            return r["id"]
        cur.execute(
            "SELECT signer_id, sig_ftp_path, date_ftp_path, sig_png, date_png FROM sign_signer_stroke WHERE signer_id=%s",
            (v,),
        )
        leg = cur.fetchone()
    leg = _hydrate_stroke_storage_row(dict(leg)) if leg else None
    if not leg:
        return None
    sig_b = leg.get("sig_png") or b""
    date_b = leg.get("date_png") or b""
    if not sig_b or not date_b:
        return None
    sig_sha = hashlib.sha256(sig_b).hexdigest()
    date_sha = hashlib.sha256(date_b).hexdigest()
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id FROM sign_stroke_set WHERE signer_id=%s AND locale=%s AND sig_sha256=%s AND date_sha256=%s",
            (v, "zh", sig_sha, date_sha),
        )
        ex = cur.fetchone()
        if ex:
            return ex["id"]
        set_id = uuid.uuid4().hex
        cur.execute(
            "INSERT INTO sign_stroke_set "
            "(id, signer_id, locale, sig_sha256, date_sha256, sig_ftp_path, date_ftp_path, "
            "sig_size, date_size, sig_png, date_png) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
            (
                set_id,
                v,
                "zh",
                sig_sha,
                date_sha,
                leg.get("sig_ftp_path"),
                leg.get("date_ftp_path"),
                int(len(sig_b)),
                int(len(date_b)),
                None if leg.get("sig_ftp_path") else sig_b,
                None if leg.get("date_ftp_path") else date_b,
            ),
        )
        return set_id


def _upsert_stroke_set_core(
    conn, signer_id: str, locale: str, sig_b: bytes, date_b: bytes
) -> Dict[str, Any]:
    sig_sha = hashlib.sha256(sig_b).hexdigest()
    date_sha = hashlib.sha256(date_b).hexdigest()
    sig_size = int(len(sig_b))
    date_size = int(len(date_b))
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id FROM sign_stroke_set WHERE signer_id=%s AND locale=%s AND sig_sha256=%s AND date_sha256=%s",
            (signer_id, locale, sig_sha, date_sha),
        )
        ex = cur.fetchone()
        overwrote = bool(ex)
        set_id = ex["id"] if ex else uuid.uuid4().hex
    sig_path, sig_err = _ftp_upload_bytes_or_mysql(sig_b, f"sign/strokeset/{set_id}/sig.png")
    date_path, date_err = _ftp_upload_bytes_or_mysql(date_b, f"sign/strokeset/{set_id}/date.png")
    ftp_row_err_parts: List[str] = []
    if not sig_path and sig_err:
        ftp_row_err_parts.append("签名:" + sig_err)
    if not date_path and date_err:
        ftp_row_err_parts.append("日期:" + date_err)
    ftp_last_error = ("; ".join(ftp_row_err_parts))[:512] if ftp_row_err_parts else None
    if sig_path and date_path:
        ftp_last_error = None
    sig_b_store = None if sig_path else sig_b
    date_b_store = None if date_path else date_b
    with conn.cursor() as cur:
        if ex:
            cur.execute(
                "UPDATE sign_stroke_set SET "
                "sig_ftp_path=%s, date_ftp_path=%s, sig_size=%s, date_size=%s, "
                "sig_sha256=%s, date_sha256=%s, sig_png=%s, date_png=%s, ftp_last_error=%s WHERE id=%s",
                (
                    sig_path,
                    date_path,
                    sig_size,
                    date_size,
                    sig_sha,
                    date_sha,
                    sig_b_store,
                    date_b_store,
                    ftp_last_error,
                    set_id,
                ),
            )
        else:
            cur.execute(
                "INSERT INTO sign_stroke_set "
                "(id, signer_id, locale, sig_sha256, date_sha256, sig_ftp_path, date_ftp_path, "
                "sig_size, date_size, sig_png, date_png, ftp_last_error) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (
                    set_id,
                    signer_id,
                    locale,
                    sig_sha,
                    date_sha,
                    sig_path,
                    date_path,
                    sig_size,
                    date_size,
                    sig_b_store,
                    date_b_store,
                    ftp_last_error,
                ),
            )
    return {"stroke_set_id": set_id, "overwritten": overwrote}


def _sync_legacy_signer_stroke(
    conn, signer_id: str, sig_b: bytes, date_b: bytes
) -> None:
    sig_sha = hashlib.sha256(sig_b).hexdigest()
    date_sha = hashlib.sha256(date_b).hexdigest()
    sig_size = int(len(sig_b))
    date_size = int(len(date_b))
    sig_path, sig_err = _ftp_upload_bytes_or_mysql(sig_b, f"sign/strokes/{signer_id}/sig.png")
    date_path, date_err = _ftp_upload_bytes_or_mysql(date_b, f"sign/strokes/{signer_id}/date.png")
    leg_parts: List[str] = []
    if not sig_path and sig_err:
        leg_parts.append("签名:" + sig_err)
    if not date_path and date_err:
        leg_parts.append("日期:" + date_err)
    ftp_last_error = ("; ".join(leg_parts))[:512] if leg_parts else None
    if sig_path and date_path:
        ftp_last_error = None
    sig_b_store = None if sig_path else sig_b
    date_b_store = None if date_path else date_b
    with conn.cursor() as cur:
        cur.execute(
            "INSERT INTO sign_signer_stroke "
            "(signer_id, sig_ftp_path, date_ftp_path, sig_size, date_size, sig_sha256, date_sha256, sig_png, date_png, ftp_last_error) "
            "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
            "ON DUPLICATE KEY UPDATE "
            "sig_ftp_path=VALUES(sig_ftp_path), date_ftp_path=VALUES(date_ftp_path), "
            "sig_size=VALUES(sig_size), date_size=VALUES(date_size), "
            "sig_sha256=VALUES(sig_sha256), date_sha256=VALUES(date_sha256), "
            "sig_png=VALUES(sig_png), date_png=VALUES(date_png), "
            "ftp_last_error=VALUES(ftp_last_error)",
            (
                signer_id,
                sig_path,
                date_path,
                sig_size,
                date_size,
                sig_sha,
                date_sha,
                sig_b_store,
                date_b_store,
                ftp_last_error,
            ),
        )


def list_signers() -> List[dict]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT s.id, s.display_name AS name, s.created_at,
                    ( (st.sig_png IS NOT NULL AND LENGTH(st.sig_png) > 0) OR (st.sig_ftp_path IS NOT NULL AND st.sig_ftp_path <> '') ) AS leg_sig,
                    ( (st.date_png IS NOT NULL AND LENGTH(st.date_png) > 0) OR (st.date_ftp_path IS NOT NULL AND st.date_ftp_path <> '') ) AS leg_date
                FROM sign_signer s
                LEFT JOIN sign_signer_stroke st ON s.id = st.signer_id
                ORDER BY s.created_at DESC
                """
            )
            rows = cur.fetchall()
            cur.execute(
                """
                SELECT id, signer_id, locale, updated_at, sig_sha256, date_sha256
                FROM sign_stroke_set ORDER BY signer_id ASC, locale ASC, updated_at DESC
                """
            )
            set_rows = cur.fetchall() or []
            cur.execute(
                """
                SELECT id, signer_id, locale, kind, updated_at, sha256
                , ftp_path, ftp_last_error, (png IS NOT NULL AND LENGTH(png) > 0) AS has_blob
                FROM sign_stroke_item ORDER BY signer_id ASC, locale ASC, kind ASC, updated_at DESC
                """
            )
            item_rows = cur.fetchall() or []
    sets_by_signer: Dict[str, List[dict]] = {}
    for sr in set_rows:
        sid = sr["signer_id"]
        sets_by_signer.setdefault(sid, []).append(sr)
    items_by_signer: Dict[str, Dict[str, List[dict]]] = {}
    pieces_by_signer: Dict[str, Dict[str, bool]] = {}
    for ir in item_rows:
        sid = ir["signer_id"]
        k = (ir.get("kind") or "").strip().lower()
        pk = normalize_piece_kind(k)
        if pk:
            if sid not in pieces_by_signer:
                pieces_by_signer[sid] = {x: False for x in all_piece_kinds()}
            pieces_by_signer[sid][pk] = True
            continue
        if k not in ("sig", "date"):
            continue
        items_by_signer.setdefault(sid, {}).setdefault(k, []).append(ir)
    out: List[dict] = []
    for r in rows:
        ts = r.get("created_at")
        signer_id = r["id"]
        stroke_sets_raw = sets_by_signer.get(signer_id, [])
        stroke_sets: List[dict] = []
        for i, sr in enumerate(stroke_sets_raw):
            ut = sr.get("updated_at")
            stroke_sets.append(
                {
                    "id": sr["id"],
                    "signer_id": signer_id,
                    "locale": sr.get("locale") or "zh",
                    "updated_at": ut.isoformat(sep=" ") if ut is not None else None,
                    "sig_sha256": sr.get("sig_sha256"),
                    "date_sha256": sr.get("date_sha256"),
                    "label": ("%s · 第 %d 套" % (("中文" if (sr.get("locale") or "zh") == "zh" else "英文"), i + 1)),
                }
            )
        has_from_sets = len(stroke_sets) > 0
        sig_items_raw = (items_by_signer.get(signer_id) or {}).get("sig", []) or []
        date_items_raw = (items_by_signer.get(signer_id) or {}).get("date", []) or []
        sig_items: List[dict] = []
        date_items: List[dict] = []
        for i, it in enumerate(sig_items_raw):
            ut = it.get("updated_at")
            ftp_path = (it.get("ftp_path") or "").strip()
            fe = (it.get("ftp_last_error") or "").strip()
            sig_items.append(
                {
                    "id": it["id"],
                    "signer_id": signer_id,
                    "locale": it.get("locale") or "zh",
                    "kind": "sig",
                    "updated_at": ut.isoformat(sep=" ") if ut is not None else None,
                    "sha256": it.get("sha256"),
                    "ftp_uploaded": bool(ftp_path),
                    "ftp_path": ftp_path or None,
                    "blob_stored": bool(it.get("has_blob")),
                    "ftp_last_error": fe or None,
                    "label": "第 %d 条" % (i + 1),
                }
            )
        for i, it in enumerate(date_items_raw):
            ut = it.get("updated_at")
            ftp_path = (it.get("ftp_path") or "").strip()
            fe = (it.get("ftp_last_error") or "").strip()
            date_items.append(
                {
                    "id": it["id"],
                    "signer_id": signer_id,
                    "locale": it.get("locale") or "zh",
                    "kind": "date",
                    "updated_at": ut.isoformat(sep=" ") if ut is not None else None,
                    "sha256": it.get("sha256"),
                    "ftp_uploaded": bool(ftp_path),
                    "ftp_path": ftp_path or None,
                    "blob_stored": bool(it.get("has_blob")),
                    "ftp_last_error": fe or None,
                    "label": "第 %d 条" % (i + 1),
                }
            )
        has_sig = bool(sig_items) or has_from_sets or bool(r.get("leg_sig"))
        has_date = bool(date_items) or has_from_sets or bool(r.get("leg_date"))
        pie = pieces_by_signer.get(signer_id) or {x: False for x in all_piece_kinds()}
        out.append(
            {
                "id": signer_id,
                "name": r["name"],
                "has_sig": has_sig,
                "has_date": has_date,
                "created_at": ts.isoformat(sep=" ") if ts is not None else None,
                "stroke_sets": stroke_sets,
                "sig_items": sig_items,
                "date_items": date_items,
                "date_piece_en": pie,
            }
        )
    return out


def insert_signer(display_name: str) -> str:
    ensure_sign_mysql()
    sid = uuid.uuid4().hex
    nm = (display_name or "").strip()[:128] or "未命名"
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO sign_signer (id, display_name) VALUES (%s, %s)",
                (sid, nm),
            )
    return sid


def delete_signer(signer_id: str) -> int:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_signer WHERE id=%s", (signer_id,))
            return cur.rowcount


def get_stroke_set_row(stroke_set_id: str) -> Optional[dict]:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT id, signer_id, locale, sig_ftp_path, date_ftp_path, sig_png, date_png "
                "FROM sign_stroke_set WHERE id=%s",
                (stroke_set_id,),
            )
            row = cur.fetchone()
    if not row:
        return None
    return _hydrate_stroke_storage_row(dict(row))


def get_signer_strokes_row(signer_id: str) -> Optional[dict]:
    """兼容：返回该签署人最近更新的一套笔迹（sign_stroke_set），否则回退 sign_signer_stroke。"""
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, signer_id, locale, sig_ftp_path, date_ftp_path, sig_png, date_png
                FROM sign_stroke_set WHERE signer_id=%s ORDER BY updated_at DESC LIMIT 1
                """,
                (signer_id,),
            )
            row = cur.fetchone()
        if row:
            return _hydrate_stroke_storage_row(dict(row))
        with conn.cursor() as cur:
            cur.execute(
                "SELECT signer_id, sig_ftp_path, date_ftp_path, sig_png, date_png FROM sign_signer_stroke WHERE signer_id=%s",
                (signer_id,),
            )
            row = cur.fetchone()
    if not row:
        return None
    return _hydrate_stroke_storage_row(dict(row))


def upsert_signer_strokes(
    signer_id: str, sig_png: Optional[bytes], date_png: Optional[bytes], locale: str = "zh"
) -> Dict[str, Any]:
    """写入 sign_stroke_set（按内容去重覆盖），并同步 legacy sign_signer_stroke 为当前合并结果。"""
    ensure_sign_mysql()
    loc = (locale or "zh").strip().lower()
    if loc not in ("zh", "en"):
        loc = "zh"
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM sign_signer WHERE id=%s", (signer_id,))
            if not cur.fetchone():
                raise ValueError("签署人不存在")
        sig_b, date_b = _merge_signer_stroke_bytes(conn, signer_id, sig_png, date_png)
        if not sig_b or not date_b:
            raise ValueError("请至少提交签名与日期笔迹（可只传其一，另一项从已有笔迹合并）")
        res = _upsert_stroke_set_core(conn, signer_id, loc, sig_b, date_b)
        _sync_legacy_signer_stroke(conn, signer_id, sig_b, date_b)
    return res


def get_file_role_signer_map(file_id: str) -> Dict[str, Any]:
    """返回 role -> {sig, date, date_mode, date_iso}（笔迹素材 id）。兼容旧 stroke_set_id / signer_id。"""
    ensure_sign_mysql()
    out: Dict[str, dict] = {}
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute(
                "SELECT role_id, signer_id, stroke_set_id, sig_item_id, date_item_id, date_mode, date_iso "
                "FROM sign_file_role_signer WHERE file_id=%s",
                (file_id,),
            )
            rows = cur.fetchall() or []
        for r in rows:
            rid = r["role_id"]
            sig_id = (r.get("sig_item_id") or "").strip()
            date_id = (r.get("date_item_id") or "").strip()
            dm = (r.get("date_mode") or "").strip() or None
            diso = (r.get("date_iso") or "").strip() or None
            if sig_id or date_id or is_composite_date_mode(dm):
                out[rid] = {
                    "sig": sig_id or None,
                    "date": date_id or None,
                    "date_mode": dm,
                    "date_iso": diso or None,
                }
                continue
            ssid = (r.get("stroke_set_id") or "").strip()
            if ssid:
                # 旧 set：拆成 items 并回填
                ss = get_stroke_set_row(ssid)
                if ss and ss.get("sig_png") and ss.get("date_png"):
                    loc = (ss.get("locale") or "zh").strip().lower()
                    sig_res = _upsert_stroke_item_core(conn, ss["signer_id"], loc, "sig", ss["sig_png"])
                    date_res = _upsert_stroke_item_core(conn, ss["signer_id"], loc, "date", ss["date_png"])
                    sig_id = sig_res.get("stroke_item_id")
                    date_id = date_res.get("stroke_item_id")
                    with conn.cursor() as cur:
                        cur.execute(
                            "UPDATE sign_file_role_signer SET sig_item_id=%s, date_item_id=%s "
                            "WHERE file_id=%s AND role_id=%s",
                            (sig_id, date_id, file_id, rid),
                        )
                    out[rid] = {"sig": sig_id, "date": date_id, "date_mode": None, "date_iso": None}
                else:
                    out[rid] = {"sig": None, "date": None, "date_mode": None, "date_iso": None}
                continue
            signer_id = (r.get("signer_id") or "").strip()
            if signer_id:
                # 旧 signer_id：找该人最新 sig/date item
                with conn.cursor() as cur:
                    cur.execute(
                        "SELECT id FROM sign_stroke_item WHERE signer_id=%s AND kind='sig' ORDER BY updated_at DESC LIMIT 1",
                        (signer_id,),
                    )
                    s1 = cur.fetchone()
                    cur.execute(
                        "SELECT id FROM sign_stroke_item WHERE signer_id=%s AND kind='date' ORDER BY updated_at DESC LIMIT 1",
                        (signer_id,),
                    )
                    d1 = cur.fetchone()
                sig_id = s1["id"] if s1 else None
                date_id = d1["id"] if d1 else None
                if sig_id or date_id:
                    with conn.cursor() as cur:
                        cur.execute(
                            "UPDATE sign_file_role_signer SET sig_item_id=%s, date_item_id=%s "
                            "WHERE file_id=%s AND role_id=%s",
                            (sig_id, date_id, file_id, rid),
                        )
                out[rid] = {"sig": sig_id, "date": date_id, "date_mode": None, "date_iso": None}
    return out


def set_file_role_signer_map(file_id: str, mapping: Dict[str, Any]) -> None:
    ensure_sign_mysql()
    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sign_file_role_signer WHERE file_id=%s", (file_id,))
            for role_id, val in (mapping or {}).items():
                if not role_id or not val:
                    continue
                rid = str(role_id)[:64]
                sig_item_id = None
                date_item_id = None
                date_mode_v: Optional[str] = None
                date_iso_v: Optional[str] = None
                if isinstance(val, dict):
                    dm_raw = (val.get("date_mode") or "").strip().lower()
                    if dm_raw in _COMPOSITE_DATE_MODES:
                        sig_item_id = str(val.get("sig") or "").strip() or None
                        date_item_id = None
                        date_iso_v = (val.get("date_iso") or "").strip() or None
                        date_mode_v = dm_raw
                        if not sig_item_id or not date_iso_v:
                            raise ValueError("笔迹拼接日期：请绑定签名素材并选择日历日期（YYYY-MM-DD）")
                        try:
                            if dm_raw == "composite_zh_ymd":
                                kinds_zh_ymd_dot(date_iso_v)
                            elif dm_raw == "composite_en_space":
                                kinds_en_dmy_space(date_iso_v)
                            else:
                                kinds_for_iso_date(date_iso_v)
                        except Exception as e:
                            raise ValueError("日期无效，需为 YYYY-MM-DD") from e
                    else:
                        sig_item_id = str(val.get("sig") or "").strip() or None
                        date_item_id = str(val.get("date") or "").strip() or None
                else:
                    # 兼容：传 stroke_set_id / signer_id
                    ssid = _resolve_map_val_to_stroke_set_id(conn, str(val).strip())
                    if ssid:
                        ss = get_stroke_set_row(ssid)
                        if ss and ss.get("sig_png") and ss.get("date_png"):
                            loc = (ss.get("locale") or "zh").strip().lower()
                            sig_item_id = _upsert_stroke_item_core(conn, ss["signer_id"], loc, "sig", ss["sig_png"]).get(
                                "stroke_item_id"
                            )
                            date_item_id = _upsert_stroke_item_core(conn, ss["signer_id"], loc, "date", ss["date_png"]).get(
                                "stroke_item_id"
                            )
                if not is_composite_date_mode(date_mode_v):
                    if not sig_item_id and not date_item_id:
                        continue
                signer_id = None
                with conn.cursor() as cur2:
                    if sig_item_id:
                        cur2.execute("SELECT signer_id FROM sign_stroke_item WHERE id=%s", (sig_item_id,))
                        rr = cur2.fetchone()
                        signer_id = (rr or {}).get("signer_id")
                    if (not signer_id) and date_item_id:
                        cur2.execute("SELECT signer_id FROM sign_stroke_item WHERE id=%s", (date_item_id,))
                        rr = cur2.fetchone()
                        signer_id = (rr or {}).get("signer_id")
                if not signer_id:
                    continue
                cur.execute(
                    "INSERT INTO sign_file_role_signer (file_id, role_id, signer_id, sig_item_id, date_item_id, date_mode, date_iso) "
                    "VALUES (%s, %s, %s, %s, %s, %s, %s)",
                    (file_id, rid, signer_id, sig_item_id, date_item_id, date_mode_v, date_iso_v),
                )
