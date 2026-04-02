# -*- coding: utf-8 -*-
"""未启用 MySQL 时：签署人笔迹与「文件-角色-签署人」映射存于会话目录（JSON + PNG）。"""
from __future__ import annotations

import json
import os
import uuid
from typing import Dict, List, Optional, Tuple

_MAX_NAME_LEN = 128


def _session_dir(inbox_root: str, sid: str) -> str:
    d = os.path.join(inbox_root, sid)
    os.makedirs(d, exist_ok=True)
    return d


def _strokes_dir(inbox_root: str, sid: str) -> str:
    d = os.path.join(_session_dir(inbox_root, sid), "strokes")
    os.makedirs(d, exist_ok=True)
    return d


def _signers_json_path(inbox_root: str, sid: str) -> str:
    return os.path.join(_session_dir(inbox_root, sid), "signers.json")


def _role_map_json_path(inbox_root: str, sid: str) -> str:
    return os.path.join(_session_dir(inbox_root, sid), "file_role_signer.json")


def _load_json(path: str, default):
    if not os.path.isfile(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as fp:
            return json.load(fp)
    except Exception:
        return default


def _save_json(path: str, data) -> None:
    with open(path, "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, indent=2)


def list_signers(inbox_root: str, sid: str) -> List[dict]:
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list):
        rows = []
    sd = _strokes_dir(inbox_root, sid)
    out: List[dict] = []
    for r in rows:
        if not isinstance(r, dict) or not r.get("id"):
            continue
        sid_signer = r["id"]
        sig_p = os.path.join(sd, f"{sid_signer}_sig.png")
        date_p = os.path.join(sd, f"{sid_signer}_date.png")
        out.append(
            {
                "id": sid_signer,
                "name": r.get("name") or "未命名",
                "has_sig": os.path.isfile(sig_p) and os.path.getsize(sig_p) > 0,
                "has_date": os.path.isfile(date_p) and os.path.getsize(date_p) > 0,
                "created_at": None,
            }
        )
    return out


def insert_signer(inbox_root: str, sid: str, display_name: str) -> str:
    name = (display_name or "").strip()[:_MAX_NAME_LEN] or "未命名"
    signer_id = uuid.uuid4().hex
    path = _signers_json_path(inbox_root, sid)
    rows = _load_json(path, [])
    if not isinstance(rows, list):
        rows = []
    rows.append({"id": signer_id, "name": name})
    _save_json(path, rows)
    return signer_id


def delete_signer(inbox_root: str, sid: str, signer_id: str) -> int:
    path = _signers_json_path(inbox_root, sid)
    rows = _load_json(path, [])
    if not isinstance(rows, list):
        return 0
    n0 = len(rows)
    rows = [r for r in rows if isinstance(r, dict) and r.get("id") != signer_id]
    if len(rows) == n0:
        return 0
    _save_json(path, rows)
    sd = _strokes_dir(inbox_root, sid)
    for suf in ("_sig.png", "_date.png"):
        p = os.path.join(sd, signer_id + suf)
        try:
            if os.path.isfile(p):
                os.remove(p)
        except OSError:
            pass
    # 从映射中移除对该签署人的引用
    rmap = _load_json(_role_map_json_path(inbox_root, sid), {})
    if isinstance(rmap, dict):
        changed = False
        for fid, m in list(rmap.items()):
            if not isinstance(m, dict):
                continue
            nm = {k: v for k, v in m.items() if v != signer_id}
            if nm != m:
                rmap[fid] = nm
                changed = True
        if changed:
            _save_json(_role_map_json_path(inbox_root, sid), rmap)
    return 1


def _stroke_path(inbox_root: str, sid: str, signer_id: str, kind: str) -> str:
    assert kind in ("sig", "date")
    return os.path.join(_strokes_dir(inbox_root, sid), f"{signer_id}_{kind}.png")


def get_strokes(
    inbox_root: str, sid: str, signer_id: str
) -> Tuple[Optional[bytes], Optional[bytes]]:
    sig_p = _stroke_path(inbox_root, sid, signer_id, "sig")
    date_p = _stroke_path(inbox_root, sid, signer_id, "date")
    sig_b = None
    date_b = None
    if os.path.isfile(sig_p):
        with open(sig_p, "rb") as fp:
            sig_b = fp.read()
    if os.path.isfile(date_p):
        with open(date_p, "rb") as fp:
            date_b = fp.read()
    return sig_b, date_b


def upsert_strokes(
    inbox_root: str,
    sid: str,
    signer_id: str,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
) -> None:
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list) or not any(
        isinstance(r, dict) and r.get("id") == signer_id for r in rows
    ):
        raise ValueError("签署人不存在")
    sig_b, date_b = get_strokes(inbox_root, sid, signer_id)
    if sig_png is not None:
        sig_b = sig_png
    if date_png is not None:
        date_b = date_png
    if sig_b is not None:
        with open(_stroke_path(inbox_root, sid, signer_id, "sig"), "wb") as fp:
            fp.write(sig_b)
    if date_b is not None:
        with open(_stroke_path(inbox_root, sid, signer_id, "date"), "wb") as fp:
            fp.write(date_b)


def get_file_role_map(inbox_root: str, sid: str, file_id: str) -> Dict[str, str]:
    rmap = _load_json(_role_map_json_path(inbox_root, sid), {})
    if not isinstance(rmap, dict):
        return {}
    m = rmap.get(file_id)
    if not isinstance(m, dict):
        return {}
    return {str(k): str(v) for k, v in m.items() if k and v}


def set_file_role_map(inbox_root: str, sid: str, file_id: str, mapping: Dict[str, str]) -> None:
    path = _role_map_json_path(inbox_root, sid)
    rmap = _load_json(path, {})
    if not isinstance(rmap, dict):
        rmap = {}
    clean = {str(k)[:64]: str(v) for k, v in mapping.items() if k and v}
    rmap[file_id] = clean
    _save_json(path, rmap)
