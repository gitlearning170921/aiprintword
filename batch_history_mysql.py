# -*- coding: utf-8 -*-
"""
批量任务历史：在配置 MYSQL_HOST 时写入 MySQL（与签名等业务同库），列表与详情均从库读取。
暂存目录 data/batch_history/<id>/stash 仍在本地磁盘，供「重试」使用。
"""
from __future__ import annotations

import json
import logging
import os
import re
from typing import Any, Dict, List, Optional

_HID_RE = re.compile(r"^[0-9a-f]{32}$")

from sign_handlers.mysql_store import _conn_commit, ensure_sign_mysql, mysql_sign_enabled

logger = logging.getLogger("aiprintword.batch_history_db")


def compute_display_title(
    original_names: Any, result: Any, default_label: str = "批处理"
) -> str:
    """
    列表/摘要标题：多文件在同一相对路径下时取首路径的第一级目录名，否则取首文件主文件名；
    份数优先 original_names 长度，否则 result.total。
    """
    r = result if isinstance(result, dict) else {}
    try:
        n_total = int(r.get("total") or 0)
    except (TypeError, ValueError):
        n_total = 0
    names: List[str] = []
    for x in original_names or []:
        s = str(x).replace("\\", "/").strip()
        if s:
            names.append(s)
    n = len(names) if names else n_total
    if n <= 0:
        n = n_total or 1
    if not names:
        return f"{default_label}等{n}份文件处理记录"
    first = names[0]
    segs = [p for p in first.split("/") if p.strip()]
    if len(segs) >= 2:
        label = segs[0].strip()
    else:
        base = os.path.basename(first)
        label, _ = os.path.splitext(base)
        label = (label or "").strip()
    if not label:
        label = default_label
    if len(label) > 100:
        label = label[:97] + "…"
    return f"{label}等{n}份文件处理记录"


def enabled() -> bool:
    return mysql_sign_enabled()


def _summary_from_record(rec: Dict[str, Any]) -> Dict[str, Any]:
    r = rec.get("result") or {}
    title = rec.get("display_title") or compute_display_title(
        rec.get("original_names"), r
    )
    return {
        "id": rec["id"],
        "created_at": rec.get("created_at", ""),
        "run_mode": rec.get("run_mode"),
        "title": title,
        "total": int(r.get("total") or 0),
        "ok": int(r.get("ok") or 0),
        "failed": int(r.get("failed") or 0),
        "has_zip": bool(rec.get("download_token")),
        "has_stash": bool(rec.get("has_stash")),
        "success": bool(rec.get("payload_ok")),
    }


def save_record(record: Dict[str, Any]) -> bool:
    if not enabled():
        return False
    hid = record.get("id")
    if not hid or not str(hid).strip():
        return False
    hid = str(hid).strip()
    if not _HID_RE.match(hid):
        return False
    ensure_sign_mysql()
    try:
        js = json.dumps(record, ensure_ascii=False)
    except (TypeError, ValueError) as e:
        logger.warning("batch history json encode failed: %s", e)
        return False
    ca = str(record.get("created_at") or "")
    try:
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO app_batch_history (id, created_at, record_json) VALUES (%s,%s,%s) "
                    "ON DUPLICATE KEY UPDATE created_at=VALUES(created_at), record_json=VALUES(record_json)",
                    (hid, ca, js),
                )
    except Exception as e:
        logger.exception("batch history save_record failed: %s", e)
        return False
    return True


def list_summaries(limit: int) -> List[Dict[str, Any]]:
    if not enabled():
        return []
    ensure_sign_mysql()
    lim = max(1, min(int(limit), 2000))
    try:
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT record_json FROM app_batch_history ORDER BY created_at DESC LIMIT %s",
                    (lim,),
                )
                rows = cur.fetchall()
    except Exception as e:
        logger.exception("batch history list_summaries failed: %s", e)
        return []
    out: List[Dict[str, Any]] = []
    for r in rows:
        try:
            rec = json.loads(r["record_json"])
            if isinstance(rec, dict) and rec.get("id"):
                out.append(_summary_from_record(rec))
        except Exception:
            continue
    return out


def get_record(hid: str) -> Optional[Dict[str, Any]]:
    if not enabled():
        return None
    hid = (hid or "").strip()
    if not _HID_RE.match(hid):
        return None
    ensure_sign_mysql()
    try:
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT record_json FROM app_batch_history WHERE id=%s",
                    (hid,),
                )
                row = cur.fetchone()
    except Exception as e:
        logger.exception("batch history get_record failed: %s", e)
        return None
    if not row:
        return None
    try:
        rec = json.loads(row["record_json"])
        return rec if isinstance(rec, dict) else None
    except Exception:
        return None


def trim_to_max(max_keep: int) -> List[str]:
    """仅保留最近 max_keep 条；返回仍存在的 id 列表（用于清理本地 stash 目录）。"""
    if not enabled():
        return []
    mk = max(1, min(int(max_keep), 2000))
    ensure_sign_mysql()
    try:
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id FROM app_batch_history ORDER BY created_at DESC LIMIT %s",
                    (mk,),
                )
                rows = cur.fetchall()
                if not rows:
                    cur.execute("DELETE FROM app_batch_history")
                    return []
                keep = tuple(r["id"] for r in rows)
                ph = ",".join(["%s"] * len(keep))
                cur.execute(
                    f"DELETE FROM app_batch_history WHERE id NOT IN ({ph})",
                    keep,
                )
                return list(keep)
    except Exception as e:
        logger.exception("batch history trim_to_max failed: %s", e)
        return []


def migrate_from_disk(root: str, max_keep: int) -> Dict[str, Any]:
    """将本地 data/batch_history/<id>/record.json 中尚未入库的记录导入 MySQL。"""
    stats: Dict[str, Any] = {
        "imported": 0,
        "skipped": 0,
        "errors": 0,
        "trim_keep": 0,
    }
    if not enabled():
        stats["error"] = "mysql_disabled"
        return stats
    root = os.path.abspath(root or "")
    if not root or not os.path.isdir(root):
        return stats
    ensure_sign_mysql()
    try:
        for name in os.listdir(root):
            if name == "index.json" or not _HID_RE.match(name):
                continue
            rp = os.path.join(root, name, "record.json")
            if not os.path.isfile(rp):
                continue
            try:
                with open(rp, encoding="utf-8") as f:
                    rec = json.load(f)
            except Exception:
                stats["errors"] += 1
                continue
            if not isinstance(rec, dict) or not rec.get("id"):
                stats["errors"] += 1
                continue
            hid = str(rec["id"]).strip()
            if not _HID_RE.match(hid):
                stats["errors"] += 1
                continue
            if get_record(hid) is not None:
                stats["skipped"] += 1
                continue
            if not rec.get("display_title"):
                rec["display_title"] = compute_display_title(
                    rec.get("original_names"), rec.get("result")
                )
            if save_record(rec):
                stats["imported"] += 1
            else:
                stats["errors"] += 1
    except Exception as e:
        logger.exception("batch history migrate_from_disk failed: %s", e)
        stats["errors"] += 1
    stats["trim_keep"] = len(trim_to_max(max_keep))
    return stats


def backfill_display_titles() -> int:
    """为库中缺少 display_title 的 record_json 补全标题字段。"""
    if not enabled():
        return 0
    ensure_sign_mysql()
    updated = 0
    try:
        with _conn_commit() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT id, record_json FROM app_batch_history")
                rows = cur.fetchall()
                for row in rows:
                    try:
                        rec = json.loads(row["record_json"])
                    except Exception:
                        continue
                    if not isinstance(rec, dict) or rec.get("display_title"):
                        continue
                    rec["display_title"] = compute_display_title(
                        rec.get("original_names"), rec.get("result")
                    )
                    try:
                        js = json.dumps(rec, ensure_ascii=False)
                    except Exception:
                        continue
                    cur.execute(
                        "UPDATE app_batch_history SET record_json=%s WHERE id=%s",
                        (js, row["id"]),
                    )
                    updated += 1
    except Exception as e:
        logger.exception("batch history backfill_display_titles failed: %s", e)
        return updated
    return updated
