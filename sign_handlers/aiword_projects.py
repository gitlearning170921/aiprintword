# -*- coding: utf-8 -*-
"""从 aiword 集成 API 同步项目列表（与 aiword /api/integration/projects 对齐）。"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("aiprintword.projects")


def _aiword_base_url() -> str:
    try:
        from runtime_settings.resolve import get_setting

        return (str(get_setting("AIWORD_BASE_URL") or "")).strip().rstrip("/")
    except Exception:
        return (str(__import__("os").environ.get("AIWORD_BASE_URL") or "")).strip().rstrip("/")


def _aiword_integration_secret() -> str:
    try:
        from runtime_settings.resolve import get_setting

        sec = (str(get_setting("AIWORD_INTEGRATION_SECRET") or "")).strip()
        if sec:
            return sec
        return (str(get_setting("AIWORD_HANDOFF_SECRET") or "")).strip()
    except Exception:
        import os

        return (
            (os.environ.get("AIWORD_INTEGRATION_SECRET") or "").strip()
            or (os.environ.get("AIWORD_HANDOFF_SECRET") or "").strip()
        )


def fetch_projects_from_aiword(*, timeout_sec: float = 12.0) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    """拉取 aiword 项目列表。返回 (projects, error)。"""
    base = _aiword_base_url()
    if not base:
        return [], "未配置 AIWORD_BASE_URL（请在系统设置填写 aiword 服务地址）"
    secret = _aiword_integration_secret()
    if not secret:
        return [], (
            "未配置 AIWORD_INTEGRATION_SECRET（须与 aiword 系统配置 INTEGRATION_SECRET 一致；"
            "也可暂用 AIWORD_HANDOFF_SECRET）"
        )
    url = base + "/api/integration/projects"
    try:
        import requests

        resp = requests.get(
            url,
            headers={"X-Integration-Secret": secret},
            timeout=timeout_sec,
        )
    except Exception as e:
        logger.warning("fetch aiword projects failed: %s", e)
        return [], f"无法连接 aiword（{e}）"

    if resp.status_code == 403:
        return [], "集成密钥无效：请核对 AIWORD_INTEGRATION_SECRET 与 aiword 的 INTEGRATION_SECRET"
    if resp.status_code != 200:
        try:
            body = resp.json()
            msg = body.get("message") or body.get("error") or resp.text[:200]
        except Exception:
            msg = resp.text[:200] if resp.text else f"HTTP {resp.status_code}"
        return [], f"aiword 返回错误：{msg}"

    try:
        data = resp.json()
    except Exception:
        return [], "aiword 项目列表响应不是合法 JSON"

    rows: List[Any]
    if isinstance(data, dict):
        if data.get("ok") is False:
            return [], str(data.get("error") or data.get("message") or "aiword 同步失败")
        rows = data.get("projects")
        if rows is None:
            rows = data.get("items")
        if rows is None and isinstance(data.get("data"), list):
            rows = data.get("data")
    elif isinstance(data, list):
        rows = data
    else:
        rows = []

    if not isinstance(rows, list):
        return [], "aiword 项目列表格式异常"

    out: List[Dict[str, Any]] = []
    for raw in rows:
        if not isinstance(raw, dict):
            continue
        pid = str(raw.get("id") or "").strip()
        name = str(raw.get("name") or "").strip()
        if not pid and not name:
            continue
        out.append(
            {
                "id": pid,
                "name": name,
                "project_key": str(raw.get("projectKey") or raw.get("project_key") or name).strip(),
                "registered_country": str(
                    raw.get("registeredCountry") or raw.get("registered_country") or ""
                ).strip()
                or None,
                "priority": int(raw.get("priority") or 2),
                "priority_label": str(raw.get("priorityLabel") or raw.get("priority_label") or "").strip(),
                "status": str(raw.get("status") or "active").strip().lower(),
                "status_label": str(raw.get("statusLabel") or raw.get("status_label") or "").strip(),
                "updated_at": raw.get("updatedAt") or raw.get("updated_at"),
            }
        )
    return out, None


def _projects_for_api_after_sync() -> List[Dict[str, Any]]:
    """同步完成后返回给前端的列表（优先跳过文件数统计以加快响应）。"""
    try:
        from sign_handlers import mysql_store

        if mysql_store.mysql_sign_enabled():
            return mysql_store.list_projects_with_counts(include_file_counts=False)
    except Exception:
        pass
    from sign_handlers import project_store_local

    return project_store_local.list_projects_with_counts(file_records=None)


def sync_projects_to_store() -> Dict[str, Any]:
    """拉取并 upsert 到 aiprintword 项目缓存（优先 MySQL sign_project_cache）。"""
    rows, err = fetch_projects_from_aiword()
    if err:
        return {"ok": False, "error": err, "count": 0, "inserted": 0, "updated": 0}
    if not rows:
        return {
            "ok": False,
            "error": "aiword 返回 0 个项目，请确认 aiword 中已有项目数据",
            "count": 0,
            "inserted": 0,
            "updated": 0,
        }
    try:
        from sign_handlers import mysql_store

        if mysql_store.mysql_sign_enabled():
            inserted, updated = mysql_store.upsert_project_cache(rows)
            return {
                "ok": True,
                "count": inserted + updated,
                "inserted": inserted,
                "updated": updated,
                "source": "aiword",
                "storage": "mysql",
                "projects": _projects_for_api_after_sync(),
            }
    except Exception as e:
        logger.exception("sync projects to mysql failed")
        return {
            "ok": False,
            "error": f"MySQL 写入项目缓存失败：{e}",
            "count": 0,
            "inserted": 0,
            "updated": 0,
        }

    from sign_handlers import project_store_local

    inserted, updated = project_store_local.upsert_projects_cache(rows)
    return {
        "ok": True,
        "count": inserted + updated,
        "inserted": inserted,
        "updated": updated,
        "source": "aiword",
        "storage": "local",
        "projects": project_store_local.list_projects_with_counts(file_records=None),
    }
