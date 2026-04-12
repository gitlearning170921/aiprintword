# -*- coding: utf-8 -*-
"""
FTP 文件存储（主动模式）。

约定：
- 只负责“上传/下载/删除”字节与文件，业务端负责决定远端路径组织规则。
- 默认使用主动模式（PASV=False）。
- 连接信息来自环境变量（也可由 runtime_settings 注入到环境变量层）。
"""

from __future__ import annotations

import io
import os
import posixpath
import socket
from contextlib import contextmanager
from typing import Iterator, Optional, Tuple


def _cfg() -> Tuple[str, int, str, str, str]:
    host = (os.environ.get("FTP_HOST") or "10.26.1.221").strip()
    port_s = (os.environ.get("FTP_PORT") or "2121").strip() or "2121"
    try:
        port = int(port_s)
    except Exception:
        port = 2121
    user = (os.environ.get("FTP_USER") or "aiwordftpuser").strip() or "aiwordftpuser"
    pwd = os.environ.get("FTP_PASSWORD") or os.environ.get("FTP_PASS") or ""
    # FTP_BASE_DIR 作为父目录；FTP_APP_DIR（默认 aiprintword）作为应用目录名
    parent_dir = (os.environ.get("FTP_BASE_DIR") or "/upload").strip() or "/upload"
    if not parent_dir.startswith("/"):
        parent_dir = "/" + parent_dir
    app_dir = (os.environ.get("FTP_APP_DIR") or "aiprintword").strip().strip("/") or "aiprintword"
    base_dir = posixpath.join(parent_dir, app_dir)
    return host, port, user, pwd, base_dir


@contextmanager
def _ftp() -> Iterator["FTP"]:
    # 延迟导入，避免非 FTP 场景引入开销
    from ftplib import FTP

    host, port, user, pwd, _ = _cfg()
    ftp = FTP()
    ftp.connect(host=host, port=port, timeout=20)
    # 主动模式
    try:
        ftp.set_pasv(False)
    except Exception:
        pass
    ftp.login(user=user, passwd=pwd)
    try:
        yield ftp
    finally:
        try:
            ftp.quit()
        except Exception:
            try:
                ftp.close()
            except Exception:
                pass


def _ensure_remote_dirs(ftp, remote_dir: str) -> None:
    # remote_dir 使用 posix 路径
    parts = [p for p in remote_dir.strip("/").split("/") if p]
    cur = ""
    for p in parts:
        cur = cur + "/" + p
        try:
            ftp.mkd(cur)
        except Exception:
            # 已存在/无权限等都忽略，由后续 cwd/storbinary 决定是否失败
            pass


def _join_base(rel_path: str) -> str:
    _, _, _, _, base_dir = _cfg()
    rp = rel_path.replace("\\", "/").lstrip("/")
    return posixpath.join(base_dir, rp)


def upload_bytes(data: bytes, remote_rel_path: str) -> str:
    """
    上传 bytes 到 FTP。
    返回远端绝对路径（含 base_dir）。
    """
    if data is None:
        raise ValueError("data is None")
    remote_abs = _join_base(remote_rel_path)
    remote_dir = posixpath.dirname(remote_abs)
    with _ftp() as ftp:
        _ensure_remote_dirs(ftp, remote_dir)
        bio = io.BytesIO(data)
        ftp.storbinary("STOR " + remote_abs, bio)
    return remote_abs


def upload_file(local_path: str, remote_rel_path: str) -> str:
    """
    上传本地文件到 FTP，返回远端绝对路径（含 base_dir）。
    """
    lp = os.path.abspath(local_path)
    if not os.path.isfile(lp):
        raise FileNotFoundError(lp)
    remote_abs = _join_base(remote_rel_path)
    remote_dir = posixpath.dirname(remote_abs)
    with _ftp() as ftp:
        _ensure_remote_dirs(ftp, remote_dir)
        with open(lp, "rb") as f:
            ftp.storbinary("STOR " + remote_abs, f)
    return remote_abs


def download_bytes(remote_abs_or_rel: str) -> bytes:
    """
    下载远端文件为 bytes。
    remote_abs_or_rel：可以是绝对路径（/aiprintword/..）或相对路径（自动拼 base_dir）。
    """
    p = (remote_abs_or_rel or "").strip()
    if not p:
        raise ValueError("remote path empty")
    if not p.startswith("/"):
        p = _join_base(p)
    buf = io.BytesIO()
    with _ftp() as ftp:
        ftp.retrbinary("RETR " + p, buf.write)
    return buf.getvalue()


def delete_path(remote_abs_or_rel: str) -> bool:
    p = (remote_abs_or_rel or "").strip()
    if not p:
        return False
    if not p.startswith("/"):
        p = _join_base(p)
    with _ftp() as ftp:
        try:
            ftp.delete(p)
            return True
        except Exception:
            return False

