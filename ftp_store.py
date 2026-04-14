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


def ftp_upload_configured() -> bool:
    """是否配置了 FTP 主机（有配置才尝试上传）。"""
    try:
        from runtime_settings.resolve import get_setting

        return bool(str(get_setting("FTP_HOST") or "").strip())
    except Exception:
        return bool((os.environ.get("FTP_HOST") or "").strip())


def _short_ftp_err(e: BaseException) -> str:
    s = str(e or "").strip() or type(e).__name__
    if len(s) > 480:
        return s[:477] + "…"
    return s


def try_upload_bytes(data: bytes, remote_rel_path: str) -> Tuple[Optional[str], Optional[str]]:
    """
    优先上传 bytes 到 FTP。
    返回 (远端绝对路径, 错误说明)：
    - 成功：(path, None)
    - 未配置 FTP：(None, None) — 调用方直接走 MySQL/本地，不记为失败
    - 失败：(None, 简短错误信息)
    """
    if data is None:
        return None, "data 为空"
    if not ftp_upload_configured():
        return None, None
    try:
        return upload_bytes(data, remote_rel_path), None
    except Exception as e:
        return None, _short_ftp_err(e)


def try_upload_file(local_path: str, remote_rel_path: str) -> Tuple[Optional[str], Optional[str]]:
    """
    上传本地文件到 FTP；返回值语义同 try_upload_bytes。
    """
    if not ftp_upload_configured():
        return None, None
    try:
        return upload_file(local_path, remote_rel_path), None
    except Exception as e:
        return None, _short_ftp_err(e)


def remote_file_exists(remote_rel_path: str) -> bool:
    """
    远端是否已存在该相对路径文件（相对 FTP_APP_DIR）。
    使用 SIZE；部分服务器不支持时返回 False。
    """
    if not ftp_upload_configured():
        return False
    remote_abs = _join_base(remote_rel_path)
    remote_dir = posixpath.dirname(remote_abs)
    base = posixpath.basename(remote_abs)
    if not base:
        return False
    try:
        with _ftp() as ftp:
            try:
                ftp.voidcmd("TYPE I")
            except Exception:
                pass
            try:
                ftp.cwd(remote_dir)
            except Exception:
                return False
            try:
                if ftp.size(base) is None:
                    return False
                return True
            except Exception:
                return False
    except Exception:
        return False


def _cfg() -> Tuple[str, int, str, str, str, Optional[bool]]:
    # 优先系统设置（runtime_settings），失败则回退环境变量
    try:
        from runtime_settings.resolve import get_setting

        host = str(get_setting("FTP_HOST") or "").strip() or "10.26.1.221"
        port_s = str(get_setting("FTP_PORT") or "").strip() or "2121"
        user = str(get_setting("FTP_USER") or "").strip() or "aiwordftpuser"
        pwd = str(get_setting("FTP_PASSWORD") or "")
        parent_dir = str(get_setting("FTP_BASE_DIR") or "/upload").strip() or "/upload"
        app_dir = str(get_setting("FTP_APP_DIR") or "aiprintword").strip() or "aiprintword"
        pasv = get_setting("FTP_PASV")
        pasv = bool(pasv) if pasv is not None else None
    except Exception:
        host = (os.environ.get("FTP_HOST") or "10.26.1.221").strip()
        port_s = (os.environ.get("FTP_PORT") or "2121").strip() or "2121"
        user = (os.environ.get("FTP_USER") or "aiwordftpuser").strip() or "aiwordftpuser"
        pwd = os.environ.get("FTP_PASSWORD") or os.environ.get("FTP_PASS") or ""
        parent_dir = (os.environ.get("FTP_BASE_DIR") or "/upload").strip() or "/upload"
        app_dir = (os.environ.get("FTP_APP_DIR") or "aiprintword").strip().strip("/") or "aiprintword"
        pasv_raw = (os.environ.get("FTP_PASV") or os.environ.get("FTP_PASSIVE") or "").strip().lower()
        # None = 未指定，由调用方决定是否做重试；True/False = 强制
        if pasv_raw in ("1", "true", "yes", "y", "on"):
            pasv = True
        elif pasv_raw in ("0", "false", "no", "n", "off"):
            pasv = False
        else:
            pasv = None
        # 兼容旧：FTP_PASS
        if not pwd:
            pwd = ""
    try:
        port = int(str(port_s).strip() or "2121")
    except Exception:
        port = 2121
    # FTP_BASE_DIR 作为父目录；FTP_APP_DIR（默认 aiprintword）作为应用目录名
    if not parent_dir.startswith("/"):
        parent_dir = "/" + parent_dir
    app_dir = (app_dir or "aiprintword").strip().strip("/") or "aiprintword"
    base_dir = posixpath.join(parent_dir, app_dir)
    return host, port, user, pwd, base_dir, pasv


@contextmanager
def _ftp(*, pasv: Optional[bool] = None) -> Iterator["FTP"]:
    # 延迟导入，避免非 FTP 场景引入开销
    from ftplib import FTP

    host, port, user, pwd, _, cfg_pasv = _cfg()
    if pasv is None:
        pasv = cfg_pasv
    ftp = FTP()
    ftp.connect(host=host, port=port, timeout=20)
    # 默认沿用历史：主动模式；但若指定 FTP_PASV 则尊重配置
    if pasv is not None:
        try:
            ftp.set_pasv(bool(pasv))
        except Exception:
            pass
    else:
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
    _, _, _, _, base_dir, _ = _cfg()
    rp = rel_path.replace("\\", "/").lstrip("/")
    return posixpath.join(base_dir, rp)


def _should_retry_with_passive(e: Exception) -> bool:
    msg = str(e or "")
    if not msg:
        return False
    # 常见：主动模式在 NAT/防火墙下会 425/500/timeout；被动模式可规避
    return any(
        x in msg
        for x in (
            "425",
            "Can't open data connection",
            "Data connection",
            "timed out",
            "timeout",
            "Connection refused",
        )
    )


def _should_retry_active_after_passive_server_error(e: Exception) -> bool:
    """
    vsftpd 在被动模式下若无法 bind 数据端口，会返回如 500 OOPS: vsf_sysutil_bind。
    此时可改试主动模式（需客户端对服务端可见或防火墙放行 20->客户端）。
    """
    msg = (str(e or "") or "").lower()
    if not msg:
        return False
    return "vsf_sysutil_bind" in msg or "500 oops" in msg


def upload_bytes(data: bytes, remote_rel_path: str) -> str:
    """
    上传 bytes 到 FTP。
    返回远端绝对路径（含 base_dir）。
    """
    if data is None:
        raise ValueError("data is None")
    remote_abs = _join_base(remote_rel_path)
    remote_dir = posixpath.dirname(remote_abs)

    def _stor(pasv_override: Optional[bool]) -> None:
        with _ftp(pasv=pasv_override) as ftp:
            _ensure_remote_dirs(ftp, remote_dir)
            bio = io.BytesIO(data)
            ftp.storbinary("STOR " + remote_abs, bio)

    try:
        _stor(None)
    except Exception as e:
        _, _, _, _, _, cfg_pasv = _cfg()
        if cfg_pasv is None and _should_retry_with_passive(e):
            try:
                _stor(True)
            except Exception as e2:
                if _should_retry_active_after_passive_server_error(e2):
                    _stor(False)
                else:
                    raise
        elif cfg_pasv is True and _should_retry_active_after_passive_server_error(e):
            _stor(False)
        else:
            raise
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

    def _stor_file(pasv_override: Optional[bool]) -> None:
        with _ftp(pasv=pasv_override) as ftp:
            _ensure_remote_dirs(ftp, remote_dir)
            with open(lp, "rb") as f:
                ftp.storbinary("STOR " + remote_abs, f)

    try:
        _stor_file(None)
    except Exception as e:
        _, _, _, _, _, cfg_pasv = _cfg()
        if cfg_pasv is None and _should_retry_with_passive(e):
            try:
                _stor_file(True)
            except Exception as e2:
                if _should_retry_active_after_passive_server_error(e2):
                    _stor_file(False)
                else:
                    raise
        elif cfg_pasv is True and _should_retry_active_after_passive_server_error(e):
            _stor_file(False)
        else:
            raise
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

    def _retr(pasv_override: Optional[bool]) -> None:
        with _ftp(pasv=pasv_override) as ftp:
            ftp.retrbinary("RETR " + p, buf.write)

    try:
        _retr(None)
    except Exception as e:
        _, _, _, _, _, cfg_pasv = _cfg()
        if cfg_pasv is None and _should_retry_with_passive(e):
            try:
                buf.seek(0)
                buf.truncate(0)
                _retr(True)
            except Exception as e2:
                if _should_retry_active_after_passive_server_error(e2):
                    buf.seek(0)
                    buf.truncate(0)
                    _retr(False)
                else:
                    raise
        elif cfg_pasv is True and _should_retry_active_after_passive_server_error(e):
            buf.seek(0)
            buf.truncate(0)
            _retr(False)
        else:
            raise
    return buf.getvalue()


def delete_path(remote_abs_or_rel: str) -> bool:
    p = (remote_abs_or_rel or "").strip()
    if not p:
        return False
    if not p.startswith("/"):
        p = _join_base(p)
    def _del(pasv_override: Optional[bool]) -> bool:
        with _ftp(pasv=pasv_override) as ftp:
            try:
                ftp.delete(p)
                return True
            except Exception:
                return False

    try:
        return _del(None)
    except Exception as e:
        _, _, _, _, _, cfg_pasv = _cfg()
        if cfg_pasv is None and _should_retry_with_passive(e):
            try:
                return _del(True)
            except Exception as e2:
                if _should_retry_active_after_passive_server_error(e2):
                    return _del(False)
                return False
        if cfg_pasv is True and _should_retry_active_after_passive_server_error(e):
            try:
                return _del(False)
            except Exception:
                return False
        return False

