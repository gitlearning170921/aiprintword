# -*- coding: utf-8 -*-
"""识别纠正参考图：仅存 FTP（不落地本地）。"""
from __future__ import annotations

import os
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Optional, Tuple

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}


def _sign_ftp_required() -> bool:
    try:
        from runtime_settings.resolve import get_setting

        return bool(get_setting("SIGN_FTP_REQUIRED"))
    except Exception:
        return False


def _upload_to_ftp(data: bytes, remote_rel: str) -> str:
    from ftp_store import ftp_upload_configured, try_upload_bytes

    if not ftp_upload_configured():
        if _sign_ftp_required():
            raise RuntimeError("未配置 FTP（FTP_HOST），无法保存参考图")
        raise RuntimeError("未配置 FTP，参考图须上传至 FTP 服务器")
    path, err = try_upload_bytes(data, remote_rel)
    if path:
        return path
    msg = (err or "FTP 上传失败").strip()
    if _sign_ftp_required():
        raise RuntimeError(msg)
    raise RuntimeError(msg)


def _normalize_ext(ext: str) -> str:
    e = (ext or "").lower()
    if e in {".jpg", ".jpeg"}:
        return ".jpg"
    if e in _IMAGE_EXTS:
        return e if e != ".jpeg" else ".jpg"
    return ".png"


def build_reference_image_meta(
    *,
    image_id: str,
    filename: str,
    ftp_path: str,
    uploaded_at: Optional[str] = None,
) -> Dict[str, str]:
    ts = uploaded_at or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return {
        "id": str(image_id or "")[:64],
        "filename": str(filename or "")[:200],
        "uploaded_at": str(ts)[:40],
        "ftp_path": str(ftp_path or "")[:768],
    }


def upload_reference_image_bytes(
    data: bytes,
    ext: str,
    *,
    file_id: str = "",
    shared: bool = False,
    filename: str = "",
) -> Dict[str, str]:
    if not data:
        raise ValueError("空文件")
    safe_ext = _normalize_ext(ext)
    image_id = uuid.uuid4().hex[:16]
    if shared:
        remote_rel = f"sign/detect_corrections/shared/{image_id}{safe_ext}"
    else:
        fid = (file_id or "").strip()
        if not fid:
            raise ValueError("缺少 file_id")
        remote_rel = f"sign/detect_corrections/files/{fid}/{image_id}{safe_ext}"
    ftp_path = _upload_to_ftp(data, remote_rel)
    return build_reference_image_meta(
        image_id=image_id,
        filename=filename or f"ref{safe_ext}",
        ftp_path=ftp_path,
    )


def download_reference_image(meta: Dict[str, Any]) -> Tuple[bytes, str]:
    ftp_path = str((meta or {}).get("ftp_path") or "").strip()
    if not ftp_path:
        raise FileNotFoundError("无 ftp_path")
    from ftp_store import download_bytes

    ext = os.path.splitext(ftp_path)[1].lower()
    mime = "image/png"
    if ext in {".jpg", ".jpeg"}:
        mime = "image/jpeg"
    elif ext == ".webp":
        mime = "image/webp"
    return download_bytes(ftp_path), mime


def delete_reference_image_on_ftp(meta: Dict[str, Any], *, file_id: str = "") -> bool:
    """删除 FTP 上的参考图；共享路径仅当路径含 /files/{file_id}/ 时删除（避免误删共用文件）。"""
    ftp_path = str((meta or {}).get("ftp_path") or "").strip()
    if not ftp_path:
        return False
    fid = (file_id or "").strip()
    if "/detect_corrections/shared/" in ftp_path.replace("\\", "/"):
        return False
    if fid and f"/files/{fid}/" not in ftp_path.replace("\\", "/"):
        return False
    from ftp_store import delete_path

    return delete_path(ftp_path)


def legacy_local_path(root: str, file_id: str, image_id: str) -> Optional[str]:
    """兼容旧版本地参考图（只读）。"""
    storage_dir = os.path.join(root, file_id)
    for ext in (".png", ".jpg", ".jpeg", ".webp"):
        p = os.path.join(storage_dir, image_id + ext)
        if os.path.isfile(p):
            return p
    return None
