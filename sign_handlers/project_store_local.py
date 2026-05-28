# -*- coding: utf-8 -*-
"""无 MySQL 时的项目缓存（data/sign_projects_cache.json）。"""
from __future__ import annotations

import json
import os
import time
from typing import Any, Dict, List, Optional, Tuple

_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
_CACHE_PATH = os.path.join(_ROOT, "data", "sign_projects_cache.json")


def _load_raw() -> Dict[str, Any]:
    if not os.path.isfile(_CACHE_PATH):
        return {"projects": [], "synced_at": None}
    try:
        with open(_CACHE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {"projects": [], "synced_at": None}


def _normalize_project_row(raw: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not isinstance(raw, dict):
        return None
    pid = str(raw.get("id") or "").strip()
    name = str(raw.get("name") or "").strip()
    if not pid or not name:
        return None
    country = str(
        raw.get("registered_country") or raw.get("registeredCountry") or ""
    ).strip() or None
    try:
        priority = int(raw.get("priority") or 2)
    except Exception:
        priority = 2
    return {
        "id": pid,
        "name": name,
        "project_key": str(raw.get("project_key") or raw.get("projectKey") or name).strip(),
        "registered_country": country,
        "priority": priority,
        "priority_label": str(raw.get("priority_label") or raw.get("priorityLabel") or "").strip(),
        "status": str(raw.get("status") or "active").strip().lower() or "active",
        "status_label": str(raw.get("status_label") or raw.get("statusLabel") or "").strip(),
        "updated_at": raw.get("updated_at") or raw.get("updatedAt"),
    }


def upsert_projects_cache(projects: List[Dict[str, Any]]) -> Tuple[int, int]:
    """按 id 合并写入：多的插入，重复则更新。返回 (新增数, 更新数)。"""
    data = _load_raw()
    by_id: Dict[str, Dict[str, Any]] = {}
    for p in data.get("projects") or []:
        if isinstance(p, dict):
            pid = str(p.get("id") or "").strip()
            if pid:
                by_id[pid] = dict(p)
    inserted = 0
    updated = 0
    for raw in projects or []:
        row = _normalize_project_row(raw)
        if not row:
            continue
        pid = row["id"]
        if pid in by_id:
            updated += 1
        else:
            inserted += 1
        by_id[pid] = row
    os.makedirs(os.path.dirname(_CACHE_PATH), exist_ok=True)
    payload = {
        "projects": list(by_id.values()),
        "synced_at": time.time(),
    }
    with open(_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return inserted, updated


def save_projects_cache(projects: List[Dict[str, Any]]) -> int:
    """兼容旧调用：返回写入条数。"""
    ins, upd = upsert_projects_cache(projects)
    return ins + upd


def list_projects_with_counts(file_records: Optional[List[dict]] = None) -> List[dict]:
    data = _load_raw()
    projects = list(data.get("projects") or [])
    counts: Dict[str, int] = {}
    for rec in file_records or []:
        if not rec:
            continue
        pid = str(rec.get("project_id") or "").strip()
        if pid:
            counts[pid] = counts.get(pid, 0) + 1
    out = []
    for p in projects:
        if not isinstance(p, dict):
            continue
        pid = str(p.get("id") or "").strip()
        row = dict(p)
        row["file_count"] = counts.get(pid, 0)
        row["label"] = _project_label(row)
        out.append(row)
    out.sort(key=lambda x: (-int(x.get("priority") or 2), str(x.get("name") or "")))
    return out


def get_project_by_id(project_id: str) -> Optional[dict]:
    pid = str(project_id or "").strip()
    if not pid:
        return None
    for p in _load_raw().get("projects") or []:
        if isinstance(p, dict) and str(p.get("id") or "").strip() == pid:
            copy = dict(p)
            copy["label"] = _project_label(copy)
            return copy
    return None


def _project_label(p: dict) -> str:
    name = str(p.get("name") or "").strip()
    country = str(p.get("registered_country") or "").strip()
    if country:
        return f"{name}（{country}）"
    return name or str(p.get("project_key") or "")
