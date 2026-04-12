# -*- coding: utf-8 -*-
"""未启用 MySQL 时：签署人笔迹与「文件-角色-笔迹素材」映射存于会话目录（JSON + PNG）。"""
from __future__ import annotations

import hashlib
import json
import os
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

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


def _stroke_sets_json_path(inbox_root: str, sid: str) -> str:
    # 兼容旧版（成对存储）
    return os.path.join(_session_dir(inbox_root, sid), "stroke_sets.json")


def _stroke_items_json_path(inbox_root: str, sid: str) -> str:
    return os.path.join(_session_dir(inbox_root, sid), "stroke_items.json")


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


def _load_stroke_sets(inbox_root: str, sid: str) -> List[dict]:
    rows = _load_json(_stroke_sets_json_path(inbox_root, sid), [])
    if not isinstance(rows, list):
        return []
    return [r for r in rows if isinstance(r, dict) and r.get("id") and r.get("signer_id")]


def _save_stroke_sets(inbox_root: str, sid: str, rows: List[dict]) -> None:
    _save_json(_stroke_sets_json_path(inbox_root, sid), rows)


def _set_png_path(inbox_root: str, sid: str, set_id: str, kind: str) -> str:
    assert kind in ("sig", "date")
    return os.path.join(_strokes_dir(inbox_root, sid), f"{set_id}_{kind}.png")


def _item_png_path(inbox_root: str, sid: str, item_id: str) -> str:
    return os.path.join(_strokes_dir(inbox_root, sid), f"item_{item_id}.png")


def _load_stroke_items(inbox_root: str, sid: str) -> List[dict]:
    rows = _load_json(_stroke_items_json_path(inbox_root, sid), [])
    if not isinstance(rows, list):
        return []
    out = []
    for r in rows:
        if not isinstance(r, dict) or not r.get("id") or not r.get("signer_id"):
            continue
        if (r.get("kind") or "").strip().lower() not in ("sig", "date"):
            continue
        out.append(r)
    return out


def _save_stroke_items(inbox_root: str, sid: str, rows: List[dict]) -> None:
    _save_json(_stroke_items_json_path(inbox_root, sid), rows)


def get_stroke_item_bytes(inbox_root: str, sid: str, item_id: str) -> Optional[bytes]:
    p = _item_png_path(inbox_root, sid, item_id)
    if not os.path.isfile(p) or os.path.getsize(p) <= 0:
        return None
    try:
        with open(p, "rb") as fp:
            return fp.read()
    except Exception:
        return None


def _read_set_bytes(
    inbox_root: str, sid: str, set_id: str
) -> Tuple[Optional[bytes], Optional[bytes]]:
    sig_p = _set_png_path(inbox_root, sid, set_id, "sig")
    date_p = _set_png_path(inbox_root, sid, set_id, "date")
    sig_b = None
    date_b = None
    if os.path.isfile(sig_p) and os.path.getsize(sig_p) > 0:
        with open(sig_p, "rb") as fp:
            sig_b = fp.read()
    if os.path.isfile(date_p) and os.path.getsize(date_p) > 0:
        with open(date_p, "rb") as fp:
            date_b = fp.read()
    return sig_b, date_b


def _legacy_paths(inbox_root: str, sid: str, signer_id: str) -> Tuple[str, str]:
    sd = _strokes_dir(inbox_root, sid)
    return os.path.join(sd, f"{signer_id}_sig.png"), os.path.join(sd, f"{signer_id}_date.png")


def _migrate_legacy_files(inbox_root: str, sid: str) -> None:
    """旧版每人单文件 → stroke_sets.json + 按套文件（幂等）。"""
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list):
        return
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    changed = False
    for r in rows:
        if not isinstance(r, dict) or not r.get("id"):
            continue
        signer_id = r["id"]
        if any(s.get("signer_id") == signer_id for s in stroke_sets):
            continue
        sig_p, date_p = _legacy_paths(inbox_root, sid, signer_id)
        if not (
            os.path.isfile(sig_p)
            and os.path.isfile(date_p)
            and os.path.getsize(sig_p) > 0
            and os.path.getsize(date_p) > 0
        ):
            continue
        with open(sig_p, "rb") as fp:
            sig_b = fp.read()
        with open(date_p, "rb") as fp:
            date_b = fp.read()
        set_id = uuid.uuid4().hex
        now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
        stroke_sets.append(
            {
                "id": set_id,
                "signer_id": signer_id,
                "sig_sha256": hashlib.sha256(sig_b).hexdigest(),
                "date_sha256": hashlib.sha256(date_b).hexdigest(),
                "updated_at": now,
            }
        )
        try:
            os.replace(sig_p, _set_png_path(inbox_root, sid, set_id, "sig"))
            os.replace(date_p, _set_png_path(inbox_root, sid, set_id, "date"))
        except OSError:
            with open(_set_png_path(inbox_root, sid, set_id, "sig"), "wb") as fp:
                fp.write(sig_b)
            with open(_set_png_path(inbox_root, sid, set_id, "date"), "wb") as fp:
                fp.write(date_b)
            try:
                os.remove(sig_p)
                os.remove(date_p)
            except OSError:
                pass
        changed = True
    if changed:
        stroke_sets.sort(key=lambda x: (x.get("signer_id") or "", x.get("updated_at") or ""))
        _save_stroke_sets(inbox_root, sid, stroke_sets)


def _migrate_sets_to_items(inbox_root: str, sid: str) -> None:
    """将 stroke_sets.json 拆为 stroke_items.json（幂等）。"""
    _migrate_legacy_files(inbox_root, sid)
    sets = _load_stroke_sets(inbox_root, sid)
    items = _load_stroke_items(inbox_root, sid)
    if not sets:
        return
    existing_key = {(i.get("signer_id"), i.get("locale") or "zh", i.get("kind"), i.get("sha256")) for i in items}
    changed = False
    for ss in sets:
        sid_signer = ss.get("signer_id")
        loc = (ss.get("locale") or "zh").strip().lower()
        if loc not in ("zh", "en"):
            loc = "zh"
        set_id = ss.get("id")
        if not sid_signer or not set_id:
            continue
        sig_b, date_b = _read_set_bytes(inbox_root, sid, set_id)
        if sig_b:
            sha = hashlib.sha256(sig_b).hexdigest()
            key = (sid_signer, loc, "sig", sha)
            if key not in existing_key:
                item_id = uuid.uuid4().hex
                with open(_item_png_path(inbox_root, sid, item_id), "wb") as fp:
                    fp.write(sig_b)
                items.append(
                    {
                        "id": item_id,
                        "signer_id": sid_signer,
                        "locale": loc,
                        "kind": "sig",
                        "sha256": sha,
                        "updated_at": ss.get("updated_at") or datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                existing_key.add(key)
                changed = True
        if date_b:
            sha = hashlib.sha256(date_b).hexdigest()
            key = (sid_signer, loc, "date", sha)
            if key not in existing_key:
                item_id = uuid.uuid4().hex
                with open(_item_png_path(inbox_root, sid, item_id), "wb") as fp:
                    fp.write(date_b)
                items.append(
                    {
                        "id": item_id,
                        "signer_id": sid_signer,
                        "locale": loc,
                        "kind": "date",
                        "sha256": sha,
                        "updated_at": ss.get("updated_at") or datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                )
                existing_key.add(key)
                changed = True
    if changed:
        _save_stroke_items(inbox_root, sid, items)


def upsert_stroke_item(
    inbox_root: str,
    sid: str,
    signer_id: str,
    kind: str,
    png_b: bytes,
    locale: str = "zh",
) -> Dict[str, Any]:
    """保存签名/日期素材（分开存储，按内容去重覆盖）。"""
    _migrate_sets_to_items(inbox_root, sid)
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list) or not any(
        isinstance(r, dict) and r.get("id") == signer_id for r in rows
    ):
        raise ValueError("签署人不存在")
    k = (kind or "").strip().lower()
    if k not in ("sig", "date"):
        raise ValueError("kind 须为 sig 或 date")
    loc = (locale or "zh").strip().lower()
    if loc not in ("zh", "en"):
        loc = "zh"
    sha = hashlib.sha256(png_b).hexdigest()
    items = _load_stroke_items(inbox_root, sid)
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    overwrote = False
    target_id = None
    for it in items:
        if (
            it.get("signer_id") == signer_id
            and (it.get("locale") or "zh") == loc
            and it.get("kind") == k
            and it.get("sha256") == sha
        ):
            target_id = it["id"]
            overwrote = True
            it["updated_at"] = now
            break
    if not target_id:
        target_id = uuid.uuid4().hex
        items.append(
            {
                "id": target_id,
                "signer_id": signer_id,
                "locale": loc,
                "kind": k,
                "sha256": sha,
                "updated_at": now,
            }
        )
    with open(_item_png_path(inbox_root, sid, target_id), "wb") as fp:
        fp.write(png_b)
    _save_stroke_items(inbox_root, sid, items)
    return {"stroke_item_id": target_id, "overwritten": overwrote, "kind": k, "locale": loc}


def _sets_for_signer(stroke_sets: List[dict], signer_id: str) -> List[dict]:
    return [s for s in stroke_sets if s.get("signer_id") == signer_id]


def _latest_set_for_signer(stroke_sets: List[dict], signer_id: str) -> Optional[dict]:
    cand = _sets_for_signer(stroke_sets, signer_id)
    if not cand:
        return None
    return max(cand, key=lambda x: str(x.get("updated_at") or ""))


def _resolve_local_map_val(inbox_root: str, sid: str, val: str) -> Optional[str]:
    v = (val or "").strip()
    if not v:
        return None
    _migrate_legacy_files(inbox_root, sid)
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    if any(s.get("id") == v for s in stroke_sets):
        return v
    latest = _latest_set_for_signer(stroke_sets, v)
    if latest:
        return str(latest["id"])
    return None


def list_signers(inbox_root: str, sid: str) -> List[dict]:
    _migrate_sets_to_items(inbox_root, sid)
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list):
        rows = []
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    stroke_items = _load_stroke_items(inbox_root, sid)
    by_signer: Dict[str, List[dict]] = {}
    for s in stroke_sets:
        by_signer.setdefault(str(s.get("signer_id")), []).append(s)
    for k in by_signer:
        by_signer[k].sort(key=lambda x: str(x.get("updated_at") or ""), reverse=True)
    items_by_signer: Dict[str, Dict[str, List[dict]]] = {}
    for it in stroke_items:
        sid_signer = str(it.get("signer_id"))
        kind = (it.get("kind") or "").strip().lower()
        if kind not in ("sig", "date"):
            continue
        items_by_signer.setdefault(sid_signer, {}).setdefault(kind, []).append(it)
    for sid_signer in items_by_signer:
        for kind in items_by_signer[sid_signer]:
            items_by_signer[sid_signer][kind].sort(
                key=lambda x: str(x.get("updated_at") or ""), reverse=True
            )
    out: List[dict] = []
    for r in rows:
        if not isinstance(r, dict) or not r.get("id"):
            continue
        sid_signer = r["id"]
        signer_sets = by_signer.get(sid_signer, [])
        stroke_out: List[dict] = []
        for i, ss in enumerate(signer_sets):
            stroke_out.append(
                {
                    "id": ss["id"],
                    "signer_id": sid_signer,
                    "updated_at": ss.get("updated_at"),
                    "sig_sha256": ss.get("sig_sha256"),
                    "date_sha256": ss.get("date_sha256"),
                    "label": "第 %d 套" % (i + 1),
                }
            )
        has_pair = False
        for ss in signer_sets:
            sb, db = _read_set_bytes(inbox_root, sid, ss["id"])
            if sb and db:
                has_pair = True
                break
        sig_items = []
        date_items = []
        for i, it in enumerate((items_by_signer.get(sid_signer) or {}).get("sig", []) or []):
            sig_items.append(
                {
                    "id": it["id"],
                    "signer_id": sid_signer,
                    "locale": it.get("locale") or "zh",
                    "kind": "sig",
                    "updated_at": it.get("updated_at"),
                    "sha256": it.get("sha256"),
                    "label": "第 %d 条" % (i + 1),
                }
            )
        for i, it in enumerate((items_by_signer.get(sid_signer) or {}).get("date", []) or []):
            date_items.append(
                {
                    "id": it["id"],
                    "signer_id": sid_signer,
                    "locale": it.get("locale") or "zh",
                    "kind": "date",
                    "updated_at": it.get("updated_at"),
                    "sha256": it.get("sha256"),
                    "label": "第 %d 条" % (i + 1),
                }
            )
        out.append(
            {
                "id": sid_signer,
                "name": r.get("name") or "未命名",
                "has_sig": bool(sig_items) or has_pair,
                "has_date": bool(date_items) or has_pair,
                "created_at": None,
                "stroke_sets": stroke_out,
                "sig_items": sig_items,
                "date_items": date_items,
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
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    removed_ids = {s["id"] for s in stroke_sets if s.get("signer_id") == signer_id}
    stroke_sets = [s for s in stroke_sets if s.get("signer_id") != signer_id]
    _save_stroke_sets(inbox_root, sid, stroke_sets)
    sd = _strokes_dir(inbox_root, sid)
    for set_id in removed_ids:
        for kind in ("sig", "date"):
            p = _set_png_path(inbox_root, sid, set_id, kind)
            try:
                if os.path.isfile(p):
                    os.remove(p)
            except OSError:
                pass
    for suf in ("_sig.png", "_date.png"):
        p = os.path.join(sd, signer_id + suf)
        try:
            if os.path.isfile(p):
                os.remove(p)
        except OSError:
            pass
    rmap = _load_json(_role_map_json_path(inbox_root, sid), {})
    if isinstance(rmap, dict):
        changed = False
        for fid, m in list(rmap.items()):
            if not isinstance(m, dict):
                continue
            nm = {
                k: v
                for k, v in m.items()
                if v != signer_id and str(v) not in removed_ids
            }
            if nm != m:
                rmap[fid] = nm
                changed = True
        if changed:
            _save_json(_role_map_json_path(inbox_root, sid), rmap)
    return 1


def get_strokes(
    inbox_root: str, sid: str, signer_id: str
) -> Tuple[Optional[bytes], Optional[bytes]]:
    """兼容：该签署人最近更新的一套笔迹。"""
    _migrate_legacy_files(inbox_root, sid)
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    latest = _latest_set_for_signer(stroke_sets, signer_id)
    if not latest:
        return None, None
    return _read_set_bytes(inbox_root, sid, latest["id"])


def get_strokes_for_set(
    inbox_root: str, sid: str, stroke_set_id: str
) -> Tuple[Optional[bytes], Optional[bytes]]:
    _migrate_legacy_files(inbox_root, sid)
    return _read_set_bytes(inbox_root, sid, stroke_set_id)


def upsert_strokes(
    inbox_root: str,
    sid: str,
    signer_id: str,
    sig_png: Optional[bytes],
    date_png: Optional[bytes],
    locale: str = "zh",
) -> Dict[str, Any]:
    _migrate_legacy_files(inbox_root, sid)
    rows = _load_json(_signers_json_path(inbox_root, sid), [])
    if not isinstance(rows, list) or not any(
        isinstance(r, dict) and r.get("id") == signer_id for r in rows
    ):
        raise ValueError("签署人不存在")
    stroke_sets = _load_stroke_sets(inbox_root, sid)
    latest = _latest_set_for_signer(stroke_sets, signer_id)
    sig_b, date_b = (None, None)
    if latest:
        sig_b, date_b = _read_set_bytes(inbox_root, sid, latest["id"])
    if sig_png is not None:
        sig_b = sig_png
    if date_png is not None:
        date_b = date_png
    if not sig_b or not date_b:
        raise ValueError("请至少提交签名与日期笔迹（可只传其一，另一项从已有笔迹合并）")
    loc = (locale or "zh").strip().lower()
    if loc not in ("zh", "en"):
        loc = "zh"
    sig_sha = hashlib.sha256(sig_b).hexdigest()
    date_sha = hashlib.sha256(date_b).hexdigest()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    overwrote = False
    target_id = None
    for s in stroke_sets:
        if (
            s.get("signer_id") == signer_id
            and (s.get("locale") or "zh") == loc
            and s.get("sig_sha256") == sig_sha
            and s.get("date_sha256") == date_sha
        ):
            target_id = s["id"]
            overwrote = True
            s["updated_at"] = now
            break
    if not target_id:
        target_id = uuid.uuid4().hex
        stroke_sets.append(
            {
                "id": target_id,
                "signer_id": signer_id,
                "locale": loc,
                "sig_sha256": sig_sha,
                "date_sha256": date_sha,
                "updated_at": now,
            }
        )
    with open(_set_png_path(inbox_root, sid, target_id, "sig"), "wb") as fp:
        fp.write(sig_b)
    with open(_set_png_path(inbox_root, sid, target_id, "date"), "wb") as fp:
        fp.write(date_b)
    _save_stroke_sets(inbox_root, sid, stroke_sets)
    return {"stroke_set_id": target_id, "overwritten": overwrote}


def get_file_role_map(inbox_root: str, sid: str, file_id: str) -> Dict[str, str]:
    rmap = _load_json(_role_map_json_path(inbox_root, sid), {})
    if not isinstance(rmap, dict):
        return {}
    m = rmap.get(file_id)
    if not isinstance(m, dict):
        return {}
    out: Dict[str, str] = {}
    for k, v in m.items():
        if not k or not v:
            continue
        resolved = _resolve_local_map_val(inbox_root, sid, str(v))
        if resolved:
            out[str(k)] = resolved
    return out


def set_file_role_map(inbox_root: str, sid: str, file_id: str, mapping: Dict[str, str]) -> None:
    path = _role_map_json_path(inbox_root, sid)
    rmap = _load_json(path, {})
    if not isinstance(rmap, dict):
        rmap = {}
    clean: Dict[str, str] = {}
    for k, v in mapping.items():
        if not k or not v:
            continue
        rid = _resolve_local_map_val(inbox_root, sid, str(v).strip())
        if rid:
            clean[str(k)[:64]] = rid
    rmap[file_id] = clean
    _save_json(path, rmap)
