# -*- coding: utf-8 -*-
"""
Batch print web service for Word/Excel/PDF.
"""
import base64
import io
import json
import logging
import os
import socket
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from queue import Queue, Empty
from dataclasses import dataclass, field
from typing import Optional

from flask import (
    Flask,
    request,
    jsonify,
    send_from_directory,
    send_file,
    Response,
    session,
)
from werkzeug.utils import secure_filename

LOG_LEVEL = (os.environ.get("AIPRINTWORD_LOG_LEVEL") or "INFO").upper()
if not logging.getLogger().handlers:
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL, logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
logger = logging.getLogger("aiprintword.web")


def _is_path_under_parent(parent_abs: str, path_abs: str) -> bool:
    """path_abs 是否为 parent_abs 本身或其子路径（大小写按平台归一）。"""
    parent = os.path.normcase(os.path.abspath(parent_abs))
    child = os.path.normcase(os.path.abspath(path_abs))
    if parent == child:
        return True
    if not parent.endswith(os.sep):
        parent = parent + os.sep
    return child.startswith(parent)


def _parse_incremental_output_dir(form):
    """
    处理完成后立即落盘的根目录（运行 Flask 的机器上的绝对路径）。
    若设置环境变量 AIPRINTWORD_ALLOWED_OUTPUT_PARENT，则输出目录必须位于该路径之下。
    """
    raw = (form.get("incremental_output_dir") or "").strip()
    if not raw:
        return None
    path = os.path.abspath(os.path.normpath(raw))
    try:
        from runtime_settings.resolve import get_setting

        allowed_raw = (str(get_setting("AIPRINTWORD_ALLOWED_OUTPUT_PARENT") or "")).strip()
    except Exception:
        allowed_raw = (os.environ.get("AIPRINTWORD_ALLOWED_OUTPUT_PARENT") or "").strip()
    if allowed_raw:
        allowed_abs = os.path.abspath(os.path.normpath(allowed_raw))
        if not _is_path_under_parent(allowed_abs, path):
            raise ValueError(
                "输出目录必须位于环境变量 AIPRINTWORD_ALLOWED_OUTPUT_PARENT 所允许的目录之下。"
                f" 当前允许根：{allowed_abs}"
            )
    try:
        os.makedirs(path, exist_ok=True)
    except OSError as e:
        raise ValueError(f"无法创建或使用输出目录：{path}（{e}）")
    if not os.path.isdir(path):
        raise ValueError(f"输出路径不是目录：{path}")
    return path


@dataclass
class _BatchJob:
    job_id: str
    cancel_event: threading.Event
    pause_event: threading.Event
    skip_current_event: threading.Event = field(default_factory=threading.Event)


_BATCH_JOB_REGISTRY = {}

# 本轮终版下载映射（local）：key -> {created_at, items:[{ftp_path, filename}]}
_FINAL_DOWNLOAD_CACHE = {}

# Allowed file extensions（批处理实际处理的文档类型）
ALLOWED_EXT = {
    ".doc", ".docx", ".docm",
    ".xls", ".xlsx", ".xlsm",
    ".pdf",
}
# 可上传的压缩包：服务端解压后仅将 ALLOWED_EXT 内的成员加入批处理
ARCHIVE_UPLOAD_EXT = {".zip", ".7z", ".rar"}
# 解压防护（单个压缩包）
_ARCHIVE_EXTRACT_MAX_FILES = 500
_ARCHIVE_EXTRACT_MAX_TOTAL_UNCOMPRESSED = 250 * 1024 * 1024
_ARCHIVE_EXTRACT_SINGLE_MAX = 80 * 1024 * 1024

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024  # 256MB total upload limit
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or "aiprintword-dev-secret-change-with-env"

# Project root
ROOT = os.path.dirname(os.path.abspath(__file__))
# 与 /api/aiprintword-build 中 build 字段一致；用于确认 5050 是否加载了当前这份 app.py
AIPRINTWORD_WEB_BUILD = 6
BATCH_EXPORT_ROOT = os.path.join(ROOT, "data", "batch_exports")
BATCH_HISTORY_ROOT = os.path.join(ROOT, "data", "batch_history")
_BATCH_EXPORT_TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")
# 可选：单行口令文件（与 .env 二选一即可，适合服务账号无 .env 拷贝权限时）
_ADMIN_TOKEN_FILE_DEFAULT = os.path.join(ROOT, "data", "admin_token.txt")
_ENV_FILE_ENCODINGS = ("utf-8-sig", "utf-8", "gbk")


def _resolved_dotenv_path() -> str:
    """项目 .env 绝对路径；可通过环境变量 AIPRINTWORD_DOTENV_PATH 指定（服务的工作目录与代码目录不一致时）。"""
    o = (os.environ.get("AIPRINTWORD_DOTENV_PATH") or "").strip().strip('"').strip("'")
    if o:
        return os.path.abspath(os.path.expanduser(o))
    return os.path.join(ROOT, ".env")


def _read_key_from_env_file(path: str, key: str) -> str:
    """仅从文件读取键值（不写 os.environ）；支持 KEY= 与 KEY =；多编码尝试。"""
    if not os.path.isfile(path):
        return ""
    key_norm = (key or "").strip()
    if not key_norm:
        return ""
    for enc in _ENV_FILE_ENCODINGS:
        try:
            with open(path, "r", encoding=enc) as f:
                for raw in f:
                    line = raw.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" not in line:
                        continue
                    k, _, val = line.partition("=")
                    if k.strip() != key_norm:
                        continue
                    val = val.strip()
                    if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                        val = val[1:-1]
                    return (val or "").strip()
            return ""
        except UnicodeDecodeError:
            continue
        except OSError:
            return ""
    return ""


def _parse_env_file_value(path: str, key: str) -> None:
    """从 .env 逐行解析键值并写入 os.environ（支持 KEY=val 与 KEY = val；兜底 dotenv 未生效）。"""
    val = _read_key_from_env_file(path, key)
    if val:
        os.environ[(key or "").strip()] = val


def _load_project_dotenv() -> None:
    """加载项目根目录 .env；保证管理口令等关键变量进入 os.environ。"""
    path = _resolved_dotenv_path()
    try:
        from dotenv import load_dotenv

        try:
            load_dotenv(path, override=True, encoding="utf-8-sig")
        except TypeError:
            load_dotenv(path, override=True)
    except ImportError:
        pass
    if not (os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip():
        _parse_env_file_value(path, "AIPRINTWORD_ADMIN_TOKEN")


_load_project_dotenv()
logger.info(
    ".env 路径=%s 存在=%s AIPRINTWORD_ADMIN_TOKEN=%s",
    _resolved_dotenv_path(),
    os.path.isfile(_resolved_dotenv_path()),
    "已配置" if (os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip() else "未配置",
)


def _sync_admin_token_from_disk() -> None:
    """
    每次管理 API 调用前从磁盘同步口令（不依赖仅启动时 load 一次）。
    解决：Flask 重载子进程、多实例、以及 .env 写成「KEY = 值」时旧版手动解析失败等。
    """
    path = _resolved_dotenv_path()
    try:
        from dotenv import load_dotenv

        try:
            load_dotenv(path, override=True, encoding="utf-8-sig")
        except TypeError:
            load_dotenv(path, override=True)
    except Exception:
        pass
    _parse_env_file_value(path, "AIPRINTWORD_ADMIN_TOKEN")
    if (os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip():
        return
    token_paths = []
    ext = (os.environ.get("AIPRINTWORD_ADMIN_TOKEN_FILE") or "").strip()
    if ext:
        token_paths.append(ext)
    token_paths.append(_ADMIN_TOKEN_FILE_DEFAULT)
    for fp in token_paths:
        if not fp or not os.path.isfile(fp):
            continue
        try:
            with open(fp, encoding="utf-8-sig") as f:
                for raw in f:
                    line = raw.strip()
                    if not line or line.startswith("#"):
                        continue
                    os.environ["AIPRINTWORD_ADMIN_TOKEN"] = line
                    break
        except OSError:
            continue
        if (os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip():
            break


def _batch_history_max() -> int:
    try:
        from runtime_settings.resolve import get_setting

        v = int(get_setting("AIPRINTWORD_HISTORY_MAX"))
        return max(5, min(v, 500))
    except Exception:
        return 50


def _sign_mysql_max_files() -> int:
    try:
        from runtime_settings.resolve import get_setting

        v = int(get_setting("SIGN_MYSQL_MAX_FILES"))
        return max(1, min(v, 10000))
    except Exception:
        return 500


def _sign_mysql_max_signed() -> int:
    try:
        from runtime_settings.resolve import get_setting

        v = int(get_setting("SIGN_MYSQL_MAX_SIGNED"))
        return max(1, min(v, 50000))
    except Exception:
        return 2000


# Common COM error codes
_COM_ERROR_CODES = (-2147352567, -2147467259, -2146827864)


def _format_com_error(e):
    """Format COM exceptions to user-friendly text."""
    try:
        s = str(e)
        if "?" in s and "?" in s:
            return s
    except Exception:
        pass
    try:
        if getattr(e, "args", None) and len(e.args) > 0 and e.args[0] in _COM_ERROR_CODES:
            return "????????????? WPS/Excel ?????????????????"
    except Exception:
        pass
    return str(e)


def _allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXT


def _is_archive_upload(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ARCHIVE_UPLOAD_EXT


def _zip_strip_leading_windows_drive(norm: str) -> str:
    """去掉 ZIP 内常见的「盘符绝对路径」前缀，如 C:/xxx、D:\\xxx（否则首段 C: 含冒号会被安全校验整包拒绝）。"""
    s = norm.replace("\\", "/").strip()
    while len(s) >= 3 and s[0].isalpha() and s[1] == ":" and s[2] == "/":
        s = s[3:].lstrip("/")
    return s


def _zip_inner_path_safe(name: str) -> Optional[str]:
    """ZIP 内相对路径规范化；拒绝路径穿越与非法段。"""
    if not name or str(name).endswith("/"):
        return None
    norm = _zip_strip_leading_windows_drive(str(name))
    if not norm:
        return None
    parts = [p for p in norm.split("/") if p]
    if not parts or ".." in parts:
        return None
    if parts[0] in ("/", "\\"):
        return None
    for p in parts:
        for ch in '<>:"|?*':
            if ch in p:
                return None
    return "/".join(parts)


def _expand_zipfile_members(zf: zipfile.ZipFile, extract_root: str, zip_rel: str):
    """从已打开的 ZipFile 中解压 ALLOWED_EXT 成员到 extract_root；返回 [(绝对路径, 展示相对名), ...]。"""
    zip_rel = zip_rel.replace("\\", "/").strip()
    out = []
    abs_root = os.path.abspath(extract_root)
    if not abs_root.endswith(os.sep):
        abs_root_prefix = abs_root + os.sep
    else:
        abs_root_prefix = abs_root
    total_uncompressed = 0
    infos = zf.infolist()
    for info in infos:
        if info.is_dir():
            continue
        name_candidates = [info.filename]
        # 3.11+ 可用 metadata_encoding；旧版常见：未设 UTF-8 位时中文名被当成 cp437，需 gbk 再解一次
        if sys.version_info < (3, 11) and not (getattr(info, "flag_bits", 0) & 0x800):
            try:
                alt = info.filename.encode("cp437").decode("gbk")
                if alt and alt not in name_candidates:
                    name_candidates.append(alt)
            except (UnicodeEncodeError, UnicodeDecodeError, LookupError):
                pass
        inner = None
        for cand in name_candidates:
            inner = _zip_inner_path_safe(cand)
            if inner is not None:
                break
        if inner is None:
            continue
        ext = os.path.splitext(inner)[1].lower()
        if ext not in ALLOWED_EXT:
            continue
        try:
            sz = int(getattr(info, "file_size", 0) or 0)
        except (TypeError, ValueError):
            sz = 0
        if sz > _ARCHIVE_EXTRACT_SINGLE_MAX:
            continue
        if total_uncompressed + sz > _ARCHIVE_EXTRACT_MAX_TOTAL_UNCOMPRESSED:
            logger.warning("zip uncompressed total cap reached rel=%s", zip_rel)
            break
        dest = os.path.join(extract_root, *inner.split("/"))
        abs_dest = os.path.abspath(dest)
        if not (abs_dest == abs_root or abs_dest.startswith(abs_root_prefix)):
            continue
        try:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
        except OSError:
            continue
        try:
            with zf.open(info, "r") as src, open(dest, "wb") as dst:
                shutil.copyfileobj(src, dst, length=65536)
        except Exception as e:
            logger.warning("zip extract member failed %s in %s: %s", inner, zip_rel, e)
            try:
                if os.path.isfile(dest):
                    os.remove(dest)
            except OSError:
                pass
            continue
        total_uncompressed += sz
        rel_out = zip_rel.rstrip("/") + "/" + inner
        out.append((dest, rel_out))
        if len(out) >= _ARCHIVE_EXTRACT_MAX_FILES:
            logger.warning("zip extract file count cap rel=%s", zip_rel)
            break
    if not out:
        n_nd = sum(1 for i in infos if not i.is_dir())
        if n_nd:
            exts = []
            seen = set()
            for i in infos:
                if i.is_dir():
                    continue
                raw = _zip_strip_leading_windows_drive(str(i.filename))
                base = raw.split("/")[-1] if raw else ""
                ext = os.path.splitext(base)[1].lower() or "(无扩展名)"
                if ext not in seen:
                    seen.add(ext)
                    exts.append(ext)
                if len(exts) >= 12:
                    break
            logger.warning(
                "zip rel=%s: 包内约 %d 个文件项，但未加入批处理（扩展名须为 doc/docx/xls/xlsx/pdf 等；或为路径/解压失败）。样例扩展名: %s",
                zip_rel,
                n_nd,
                ", ".join(exts) if exts else "—",
            )
    return out


def _expand_zip_for_batch(tmp_dir: str, zip_path: str, zip_rel: str):
    """
    解压 zip，将其中符合 ALLOWED_EXT 的文件加入批处理。
    返回 [(绝对路径, 展示用相对名「压缩包名/包内路径」), ...]
    """
    zip_rel = zip_rel.replace("\\", "/").strip()

    def _open_and_expand(metadata_encoding=None):
        """返回 (能否打开 zip 结构, 解压出的文档列表)。"""
        extract_root = os.path.join(tmp_dir, "_unzip", uuid.uuid4().hex[:16])
        try:
            os.makedirs(extract_root, exist_ok=True)
        except OSError:
            return False, []
        try:
            kwargs = {}
            if metadata_encoding is not None and sys.version_info >= (3, 11):
                kwargs["metadata_encoding"] = metadata_encoding
            with zipfile.ZipFile(zip_path, "r", **kwargs) as zf:
                return True, _expand_zipfile_members(zf, extract_root, zip_rel)
        except (zipfile.BadZipFile, OSError) as e:
            if metadata_encoding is None:
                logger.warning("zip open failed rel=%s err=%s", zip_rel, e)
            return False, []

    ok, out = _open_and_expand(None)
    if not ok:
        return []
    if out:
        return out
    if sys.version_info >= (3, 11):
        for enc in ("gbk", "utf-8", "cp437"):
            ok2, out2 = _open_and_expand(enc)
            if ok2 and out2:
                logger.info("zip rel=%s used metadata_encoding=%s", zip_rel, enc)
                return out2
    return []


def _collect_allowed_from_extract_dir(extract_root: str, archive_rel: str):
    """
    遍历已解压目录，收集 ALLOWED_EXT 文件；返回 [(绝对路径, archive_rel/相对路径), ...]
    """
    archive_rel = archive_rel.replace("\\", "/").strip()
    out = []
    try:
        abs_er = os.path.abspath(extract_root)
    except OSError:
        return []
    if not os.path.isdir(abs_er):
        return []
    prefix = abs_er if abs_er.endswith(os.sep) else abs_er + os.sep
    total_bytes = 0
    for root, _, files in os.walk(abs_er):
        for fn in sorted(files):
            if len(out) >= _ARCHIVE_EXTRACT_MAX_FILES:
                return out
            full = os.path.join(root, fn)
            try:
                st = os.stat(full)
            except OSError:
                continue
            if st.st_size > _ARCHIVE_EXTRACT_SINGLE_MAX:
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext not in ALLOWED_EXT:
                continue
            try:
                rel_inner = os.path.relpath(full, abs_er).replace("\\", "/")
            except ValueError:
                continue
            if ".." in rel_inner.split("/"):
                continue
            abs_full = os.path.abspath(full)
            if not (abs_full == abs_er or abs_full.startswith(prefix)):
                continue
            if total_bytes + st.st_size > _ARCHIVE_EXTRACT_MAX_TOTAL_UNCOMPRESSED:
                logger.warning("archive walk total cap archive=%s", archive_rel)
                return out
            total_bytes += st.st_size
            rel_out = archive_rel.rstrip("/") + "/" + rel_inner
            out.append((full, rel_out))
    return out


def _resolve_7zip_cli() -> Optional[str]:
    for key in ("SEVEN_ZIP_EXE", "7ZIP_EXE", "7Z_EXE"):
        p = (os.environ.get(key) or "").strip().strip('"').strip("'")
        if p and os.path.isfile(p):
            return p
    for w in (shutil.which("7z"), shutil.which("7z.exe")):
        if w and os.path.isfile(w):
            return w
    for p in (
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ):
        if os.path.isfile(p):
            return p
    return None


def _expand_via_7zip_cli(tmp_dir: str, archive_path: str, archive_rel: str):
    """使用 7-Zip 命令行解压（支持 .7z / .rar / .zip 等）。"""
    exe = _resolve_7zip_cli()
    if not exe:
        return []
    extract_root = os.path.join(tmp_dir, "_7x", uuid.uuid4().hex[:16])
    try:
        os.makedirs(extract_root, exist_ok=True)
    except OSError:
        return []
    cmd = [exe, "x", os.path.abspath(archive_path), f"-o{extract_root}", "-aoa", "-y", "-bb0"]
    try:
        kwargs = {
            "capture_output": True,
            "text": True,
            "errors": "replace",
            "timeout": 600,
        }
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        r = subprocess.run(cmd, **kwargs)
    except (OSError, subprocess.TimeoutExpired) as e:
        logger.warning("7z cli failed %s: %s", archive_rel, e)
        return []
    if r.returncode != 0:
        logger.warning(
            "7z exit %s archive=%s err=%s",
            r.returncode,
            archive_rel,
            (r.stderr or r.stdout or "")[:500],
        )
        return []
    return _collect_allowed_from_extract_dir(extract_root, archive_rel)


def _expand_7z_py7zr(tmp_dir: str, archive_path: str, archive_rel: str):
    """无 7z.exe 时，用 py7zr 解压 .7z（纯 Python）。"""
    try:
        import py7zr
    except ImportError:
        return []
    extract_root = os.path.join(tmp_dir, "_p7z", uuid.uuid4().hex[:16])
    try:
        os.makedirs(extract_root, exist_ok=True)
        with py7zr.SevenZipFile(archive_path, mode="r") as z:
            z.extractall(path=extract_root)
    except Exception as e:
        logger.warning("py7zr extract failed %s: %s", archive_rel, e)
        return []
    return _collect_allowed_from_extract_dir(extract_root, archive_rel)


def _resolve_unrar_cli() -> Optional[str]:
    for key in ("UNRAR_EXE", "UNRAR"):
        p = (os.environ.get(key) or "").strip().strip('"').strip("'")
        if p and os.path.isfile(p):
            return p
    for w in (shutil.which("unrar"), shutil.which("UnRAR.exe"), shutil.which("unrar.exe")):
        if w and os.path.isfile(w):
            return w
    for p in (
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "WinRAR", "UnRAR.exe"),
        r"C:\Program Files\WinRAR\UnRAR.exe",
        r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
    ):
        if p and os.path.isfile(p):
            return p
    return None


def _expand_rar_unrar(tmp_dir: str, archive_path: str, archive_rel: str):
    """无 7z 时使用官方 UnRAR 解压 .rar。"""
    unrar = _resolve_unrar_cli()
    if not unrar:
        return []
    extract_root = os.path.join(tmp_dir, "_rar", uuid.uuid4().hex[:16])
    try:
        os.makedirs(extract_root, exist_ok=True)
    except OSError:
        return []
    dest = extract_root
    if not dest.endswith(os.sep):
        dest = dest + os.sep
    cmd = [unrar, "x", "-y", "-o+", os.path.abspath(archive_path), dest]
    try:
        kwargs = {
            "capture_output": True,
            "text": True,
            "errors": "replace",
            "timeout": 600,
        }
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        r = subprocess.run(cmd, **kwargs)
    except (OSError, subprocess.TimeoutExpired) as e:
        logger.warning("unrar failed %s: %s", archive_rel, e)
        return []
    if r.returncode != 0:
        logger.warning(
            "unrar exit %s archive=%s err=%s",
            r.returncode,
            archive_rel,
            (r.stderr or r.stdout or "")[:500],
        )
        return []
    return _collect_allowed_from_extract_dir(extract_root, archive_rel)


def _expand_nonzip_archive(tmp_dir: str, archive_path: str, archive_rel: str):
    """解压 .7z / .rar：优先 7-Zip 命令行，其次 py7zr（仅 7z）或 UnRAR（仅 rar）。"""
    ext = os.path.splitext(archive_rel)[1].lower()
    out = _expand_via_7zip_cli(tmp_dir, archive_path, archive_rel)
    if out:
        return out
    if ext == ".7z":
        return _expand_7z_py7zr(tmp_dir, archive_path, archive_rel)
    if ext == ".rar":
        return _expand_rar_unrar(tmp_dir, archive_path, archive_rel)
    return []


def _expand_archive_for_batch(tmp_dir: str, path: str, rel: str):
    """按扩展名解压压缩包并返回待批处理文件列表。"""
    ext = os.path.splitext(rel)[1].lower()
    if ext == ".zip":
        out = _expand_zip_for_batch(tmp_dir, path, rel)
        if out:
            return out
        # 中文 Windows「发送到压缩文件夹」等产生的 ZIP，标准库可能解码不出成员名；用 7-Zip 再试
        out7 = _expand_via_7zip_cli(tmp_dir, path, rel)
        if out7:
            logger.info("zip rel=%s expanded via 7-Zip CLI (stdlib found 0 documents)", rel)
        return out7
    if ext in (".7z", ".rar"):
        return _expand_nonzip_archive(tmp_dir, path, rel)
    return []


def _upload_relpath(client_filename):
    """
    Parse upload filename into safe relative path, e.g. "a/b/c.ext".
    Return None for invalid paths.
    """
    if not client_filename:
        return None
    norm = str(client_filename).replace("\\", "/").strip()
    if ".." in norm:
        return None
    parts = [p for p in norm.split("/") if p and p not in (".", "..")]
    if not parts:
        return None
    safe_parts = []
    for p in parts:
        s = p
        for ch in '<>:"|?*':
            s = s.replace(ch, "_")
        s = s.strip()
        if not s:
            return None
        safe_parts.append(s)
    rel = "/".join(safe_parts)
    if len(rel) > 300:
        rel = rel[-300:]
    return rel


def _write_original_backup(tmp_dir, rel, saved_path):
    """保存上传后立即复制一份原文，供单文件超时后 ZIP 仍包含未处理版本。"""
    backup = os.path.join(tmp_dir, "_aiprint_orig", *rel.split("/"))
    os.makedirs(os.path.dirname(backup), exist_ok=True)
    shutil.copy2(saved_path, backup)
    return backup


def _zip_arcname_timeout_original(original_rel):
    """ZIP 内超时条目：在文件名（不含扩展名）后加 _【超时原文】，保留相对目录结构。"""
    s = (original_rel or "").replace("\\", "/").strip()
    if not s:
        return "【超时原文】.bin"
    parts = [p for p in s.split("/") if p and p not in (".", "..")]
    if not parts:
        return "【超时原文】.bin"
    base, ext = os.path.splitext(parts[-1])
    if ext:
        parts[-1] = base + "_【超时原文】" + ext
    else:
        parts[-1] = parts[-1] + "_【超时原文】"
    return "/".join(parts)


def _detail_is_timeout_skip(d):
    """是否为单文件总超时跳过（ZIP 内改用上传备份并改文件名）。"""
    if not isinstance(d, dict):
        return False
    if d.get("timeout_skip"):
        return True
    return "【超时跳过】" in (d.get("message") or "")


def _get_printers():
    """Return available printers on Windows."""
    try:
        import win32print
        out = []
        for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            out.append(p[2])
        return out
    except Exception:
        return []


@app.route("/")
def index():
    return send_from_directory(os.path.join(ROOT, "static"), "index.html")


@app.route("/favicon.ico")
def favicon():
    """Return empty favicon response to avoid 404."""
    return "", 204


@app.route("/api/printers")
def api_printers():
    """Return available printer list."""
    return jsonify({"printers": _get_printers()})


@app.route("/api/batch-print", methods=["POST"])
def api_batch_print():
    """
    Receive upload and options, execute batch print, return JSON.
    Fields: files, check_signature, accept_revisions, printer_name, copies, dry_run.
    """
    files = request.files.getlist("files") or request.files.getlist("files[]")
    if not files or not any(f.filename for f in files):
        return jsonify({"ok": False, "error": "?????????"}), 400

    try:
        opts = _parse_batch_opts_from_form(request.form)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        original_backup_paths = []
        for f in files:
            rel = _upload_relpath(f.filename)
            if not rel:
                continue
            path = os.path.join(tmp_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            f.save(path)
            if _is_archive_upload(rel):
                for p2, r2 in _expand_archive_for_batch(tmp_dir, path, rel):
                    saved_paths.append(p2)
                    original_names.append(r2)
                    bk = _write_original_backup(tmp_dir, r2, p2)
                    original_backup_paths.append(bk)
            elif _allowed_file(rel):
                saved_paths.append(path)
                original_names.append(rel)
                bk = _write_original_backup(tmp_dir, rel, path)
                original_backup_paths.append(bk)

        if not saved_paths:
            return jsonify(
                {
                    "ok": False,
                    "error": "未找到可处理文件：压缩包内需含 .doc/.docx/.xls/.xlsx/.pdf 等文档。中文 Windows 自带的「压缩文件夹」ZIP 若解压不到文件，请安装 7-Zip（7z.exe 在 PATH 或设置 SEVEN_ZIP_EXE）。.7z 无 7z 时可 pip 安装 py7zr；.rar 需 7-Zip 或 UnRAR。",
                }
            ), 400

        from batch_print import run_batch, build_batch_modification_zip_text
        result = run_batch(
            saved_paths,
            recursive=False,
            check_formal=opts["check_formal"],
            check_signature=opts["check_signature"],
            accept_revisions=opts["accept_revisions"],
            word_content_preserve=opts["word_content_preserve"],
            word_preserve_page_count=opts["word_preserve_page_count"],
            word_image_risk_guard=opts["word_image_risk_guard"],
            word_step_timeout_sec=opts.get("word_step_timeout_sec"),
            word_skip_file_on_timeout=opts.get("word_skip_file_on_timeout"),
            file_timeout_sec=opts.get("file_timeout_sec"),
            word_font_profile=opts.get("doc_font_profile", "mixed"),
            printer_name=opts["printer_name"],
            copies=opts["copies"],
            dry_run=opts["dry_run"],
            skip_print=opts["skip_print"],
            raw_print=opts["raw_print"],
            incremental_output_dir=opts.get("incremental_output_dir"),
            relative_names=original_names,
            incremental_exists_action=opts.get("incremental_exists_action", "overwrite"),
        )
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        if opts.get("skip_print") and result.get("total"):
            try:
                token = uuid.uuid4().hex
                mod_txt = build_batch_modification_zip_text(result["details"])
                _, zip_ftp_ok, zip_ftp_err = _zip_batch_exports(
                    result["details"],
                    # 取消/中断时不打包未处理文件（这些仍是原稿副本）
                    saved_paths[: len(result.get("details") or [])],
                    original_names[: len(result.get("details") or [])],
                    original_backup_paths[: len(result.get("details") or [])],
                    token,
                    modification_report_text=mod_txt,
                )
                result["download_token"] = token
                result["download_filename"] = "processed_documents.zip"
                result["zip_ftp_uploaded"] = bool(zip_ftp_ok)
                if zip_ftp_err:
                    result["zip_ftp_error"] = zip_ftp_err
            except Exception as e:
                return jsonify({"ok": False, "error": _format_com_error(e)}), 500
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass


def _parse_sse_heartbeat_sec(form):
    """
    批处理 SSE：队列 get 超时时间 = 无 progress 时推送 heartbeat 的间隔。
    默认 120 秒；可由表单 sse_heartbeat_sec 或环境变量 AIPRINTWORD_SSE_HEARTBEAT_SEC 覆盖；限制 10–600 秒。
    """
    try:
        raw = None
        if form is not None and hasattr(form, "get"):
            raw = form.get("sse_heartbeat_sec")
        if raw is None or str(raw).strip() == "":
            try:
                from runtime_settings.resolve import get_setting

                raw = get_setting("AIPRINTWORD_SSE_HEARTBEAT_SEC")
            except Exception:
                raw = os.environ.get("AIPRINTWORD_SSE_HEARTBEAT_SEC", "120")
        v = float(str(raw).strip())
    except Exception:
        v = 120.0
    return max(10.0, min(600.0, v))


def _parse_batch_opts_from_form(form):
    """Parse run_mode: formal_save | standard_print | raw_print."""
    incremental_output_dir = _parse_incremental_output_dir(form)
    run_mode = (form.get("run_mode") or "standard_print").strip().lower()
    if run_mode not in ("formal_save", "standard_print", "raw_print"):
        run_mode = "standard_print"
    word_content_preserve = form.get("word_content_preserve", "true").lower() == "true"
    word_preserve_page_count = form.get("word_preserve_page_count", "true").lower() == "true"
    word_image_risk_guard = form.get("word_image_risk_guard", "false").lower() == "true"
    try:
        word_step_timeout_min = float(form.get("word_step_timeout_min", "60") or "60")
        word_step_timeout_min = max(0.0, min(word_step_timeout_min, 24 * 60.0))
    except Exception:
        word_step_timeout_min = 60.0
    word_skip_file_on_timeout = form.get("word_skip_file_on_timeout", "true").lower() == "true"
    try:
        # 默认 45 分钟：为 0 时单文件总超时看门狗不生效（不会自动杀进程跳过）
        file_timeout_min = float(form.get("file_timeout_min", "45") or "45")
        file_timeout_min = max(0.0, min(file_timeout_min, 24 * 60.0))
    except Exception:
        file_timeout_min = 45.0
    incremental_exists_action = (form.get("incremental_exists_action") or "overwrite").strip().lower()
    if incremental_exists_action not in ("overwrite", "skip"):
        incremental_exists_action = "overwrite"
    doc_font_profile = (
        form.get("doc_font_profile")
        or form.get("word_font_profile")
        or "mixed"
    ).strip().lower()
    if doc_font_profile not in ("chinese", "english", "mixed"):
        doc_font_profile = "mixed"
    printer_name = form.get("printer_name") or None
    if printer_name == "":
        printer_name = None
    try:
        copies = int(form.get("copies", "1") or "1")
        copies = max(1, min(copies, 99))
    except ValueError:
        copies = 1

    if run_mode == "formal_save":
        # ?????????????????? ? ???? ? ????/??????????????? ZIP
        return {
            "run_mode": run_mode,
            "check_formal": True,
            "check_signature": True,
            "accept_revisions": True,
            "dry_run": False,
            "skip_print": True,
            "raw_print": False,
            "word_content_preserve": word_content_preserve,
            "word_preserve_page_count": word_preserve_page_count,
            "word_image_risk_guard": word_image_risk_guard,
            "word_step_timeout_sec": word_step_timeout_min * 60.0,
            "word_skip_file_on_timeout": word_skip_file_on_timeout,
            "file_timeout_sec": file_timeout_min * 60.0,
            "doc_font_profile": doc_font_profile,
            "printer_name": printer_name,
            "copies": copies,
            "incremental_output_dir": incremental_output_dir,
            "incremental_exists_action": incremental_exists_action,
        }
    if run_mode == "raw_print":
        return {
            "run_mode": run_mode,
            "check_formal": False,
            "check_signature": False,
            "accept_revisions": False,
            "dry_run": False,
            "skip_print": False,
            "raw_print": True,
            "word_content_preserve": word_content_preserve,
            "word_preserve_page_count": word_preserve_page_count,
            "word_image_risk_guard": word_image_risk_guard,
            "word_step_timeout_sec": word_step_timeout_min * 60.0,
            "word_skip_file_on_timeout": word_skip_file_on_timeout,
            "file_timeout_sec": file_timeout_min * 60.0,
            "doc_font_profile": doc_font_profile,
            "printer_name": printer_name,
            "copies": copies,
            "incremental_output_dir": incremental_output_dir,
            "incremental_exists_action": incremental_exists_action,
        }
    return {
        "run_mode": run_mode,
        "check_formal": True,
        "check_signature": True,
        "accept_revisions": form.get("accept_revisions", "true").lower() == "true",
        "dry_run": form.get("dry_run", "false").lower() == "true",
        "skip_print": False,
        "raw_print": False,
        "word_content_preserve": word_content_preserve,
        "word_preserve_page_count": word_preserve_page_count,
        "word_image_risk_guard": word_image_risk_guard,
        "word_step_timeout_sec": word_step_timeout_min * 60.0,
        "word_skip_file_on_timeout": word_skip_file_on_timeout,
        "file_timeout_sec": file_timeout_min * 60.0,
        "doc_font_profile": doc_font_profile,
        "printer_name": printer_name,
        "copies": copies,
        "incremental_output_dir": incremental_output_dir,
        "incremental_exists_action": incremental_exists_action,
    }


def _zip_arcname_with_processed_ext(original_arcname: str, processed_path: str) -> str:
    """上传名为 .doc/.xls 等时，成品可能已转为 .docx/.xlsx，ZIP 内条目扩展名与磁盘成品一致。"""
    pe = os.path.splitext(processed_path)[1]
    if not pe:
        return original_arcname.replace("\\", "/")
    norm = original_arcname.replace("\\", "/")
    d, base = os.path.split(norm)
    stem = os.path.splitext(base)[0]
    new_base = stem + pe.lower()
    return f"{d}/{new_base}" if d else new_base


def _zip_batch_exports(
    details,
    saved_paths,
    original_names,
    original_backup_paths,
    token,
    modification_report_text=None,
):
    """写入 ZIP：修改明细 + 各文件。成功项必须打包落盘后的成品路径（含 .doc→.docx 等）。"""
    os.makedirs(BATCH_EXPORT_ROOT, exist_ok=True)
    zip_path = os.path.join(BATCH_EXPORT_ROOT, f"{token}.zip")
    backups = original_backup_paths or []
    n = min(len(saved_paths), len(original_names))
    if details is not None:
        n = min(n, len(details))
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        detail_text = (
            str(modification_report_text).strip()
            if modification_report_text is not None
            else ""
        )
        if not detail_text:
            detail_text = "（本批未生成修改说明正文；请核对各文件是否已按预期处理。）"
        zf.writestr(
            "\u4fee\u6539\u660e\u7ec6.txt",
            "\ufeff" + detail_text + "\n",
            compress_type=zipfile.ZIP_DEFLATED,
        )
        for i in range(n):
            name = original_names[i]
            work = saved_paths[i]
            d = details[i] if details and i < len(details) else {}
            backup = backups[i] if i < len(backups) else None
            if _detail_is_timeout_skip(d):
                if backup and os.path.isfile(backup):
                    src = backup
                else:
                    src = work
                arcname = _zip_arcname_timeout_original(name)
            else:
                proc = (d.get("processed_path") or "").strip()
                if d.get("success") and proc and os.path.isfile(proc):
                    src = proc
                    arcname = _zip_arcname_with_processed_ext(name, proc)
                else:
                    src = work
                    arcname = name.replace("\\", "/")
            if os.path.isfile(src):
                zf.write(src, arcname=arcname)
                # 终版成品优先 FTP；失败写入明细供排查，本地 ZIP 仍可用
                if d.get("success") and (d.get("processed_path") or "").strip() and os.path.isfile(src):
                    try:
                        from ftp_store import try_upload_file

                        ftp_p, up_err = try_upload_file(src, f"batch/final/{token}/{arcname}")
                        if ftp_p:
                            d["final_ftp_path"] = ftp_p
                        elif up_err:
                            d["ftp_upload_error"] = up_err
                    except Exception:
                        pass
    # 同步 ZIP 到 FTP（供下载复用；失败不影响主流程）
    zip_ftp_err: Optional[str] = None
    try:
        from ftp_store import try_upload_file

        pz, ez = try_upload_file(zip_path, f"batch/exports/{token}.zip")
        ftp_zip_ok = bool(pz)
        if ez:
            zip_ftp_err = ez
    except Exception:
        ftp_zip_ok = False
    return zip_path, ftp_zip_ok, zip_ftp_err


def _history_index_path():
    os.makedirs(BATCH_HISTORY_ROOT, exist_ok=True)
    return os.path.join(BATCH_HISTORY_ROOT, "index.json")


def _load_history_index():
    try:
        import batch_history_mysql as bhm

        if bhm.enabled():
            return {"entries": bhm.list_summaries(_batch_history_max())}
    except Exception:
        pass
    p = _history_index_path()
    if not os.path.isfile(p):
        return {"entries": []}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("entries"), list):
            return data
    except Exception:
        pass
    return {"entries": []}


def _save_history_index(data):
    p = _history_index_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning("history index save failed: %s", e)


def _history_trim_folders(keep_ids):
    if not os.path.isdir(BATCH_HISTORY_ROOT):
        return
    for name in os.listdir(BATCH_HISTORY_ROOT):
        if name == "index.json":
            continue
        path = os.path.join(BATCH_HISTORY_ROOT, name)
        if not os.path.isdir(path):
            continue
        if name in keep_ids:
            continue
        try:
            shutil.rmtree(path, ignore_errors=True)
        except Exception:
            pass


def _history_summary_from_record(rec):
    from batch_history_mysql import compute_display_title, _safe_ftp_err_text

    r = rec.get("result") or {}
    if not isinstance(r, dict):
        r = {}
    title = rec.get("display_title") or compute_display_title(
        rec.get("original_names"), r
    )
    ze = _safe_ftp_err_text(r.get("zip_ftp_error")) or _safe_ftp_err_text(
        rec.get("zip_ftp_error")
    )
    has_zip = bool(rec.get("download_token"))
    zup = r.get("zip_ftp_uploaded")
    if zup is None:
        zup = rec.get("zip_ftp_uploaded")
    if not has_zip:
        zip_ftp = None
    elif zup is True or zup == 1:
        zip_ftp = True
    elif zup is False or zup == 0:
        zip_ftp = False
    else:
        zip_ftp = None
    return {
        "id": rec["id"],
        "created_at": rec.get("created_at", ""),
        "run_mode": rec.get("run_mode"),
        "title": title,
        "total": int(r.get("total") or 0),
        "ok": int(r.get("ok") or 0),
        "failed": int(r.get("failed") or 0),
        "has_zip": has_zip,
        "has_stash": bool(rec.get("has_stash")),
        "success": bool(rec.get("payload_ok")),
        "zip_ftp": zip_ftp,
        "zip_ftp_error": (ze or None) if has_zip else None,
    }


def _record_to_default_form(record):
    """合并到 CombinedMultiDict 时在后半部分，仅补全表单未传的项。"""
    from werkzeug.datastructures import MultiDict

    o = record.get("options") or {}
    m = MultiDict()
    rm = record.get("run_mode") or o.get("run_mode") or "standard_print"
    m.set("run_mode", rm)

    def tb(x, default=False):
        return "true" if (x if x is not None else default) else "false"

    m.set("accept_revisions", tb(o.get("accept_revisions"), True))
    m.set("dry_run", tb(o.get("dry_run"), False))
    m.set("word_content_preserve", tb(o.get("word_content_preserve"), True))
    m.set("word_preserve_page_count", tb(o.get("word_preserve_page_count"), True))
    m.set("word_image_risk_guard", tb(o.get("word_image_risk_guard"), False))
    m.set("word_skip_file_on_timeout", tb(o.get("word_skip_file_on_timeout"), True))
    fts = o.get("file_timeout_sec")
    if fts is not None:
        try:
            m.set("file_timeout_min", str(max(0, int(float(fts) / 60.0))))
        except Exception:
            m.set("file_timeout_min", "45")
    wst = o.get("word_step_timeout_sec")
    if wst is not None:
        try:
            m.set("word_step_timeout_min", str(max(0, int(float(wst) / 60.0))))
        except Exception:
            m.set("word_step_timeout_min", "60")
    m.set("doc_font_profile", (o.get("doc_font_profile") or "mixed").strip())
    pn = o.get("printer_name")
    m.set("printer_name", pn if pn else "")
    try:
        m.set("copies", str(max(1, min(99, int(o.get("copies", 1))))))
    except Exception:
        m.set("copies", "1")
    iod = o.get("incremental_output_dir")
    m.set("incremental_output_dir", iod if iod else "")
    m.set("incremental_exists_action", (o.get("incremental_exists_action") or "overwrite").strip())
    m.set("sse_heartbeat_sec", str(int(_parse_sse_heartbeat_sec(None))))
    return m


def _batch_history_persist(stream_state, tmp_dir):
    """任务结束后写入历史；有失败或整体报错时整包暂存 stash 供重试。"""
    try:
        os.makedirs(BATCH_HISTORY_ROOT, exist_ok=True)
    except Exception:
        return None
    payload = stream_state.get("final_json")
    opts = stream_state.get("opts") or {}
    hid = uuid.uuid4().hex
    hist_dir = os.path.join(BATCH_HISTORY_ROOT, hid)
    try:
        os.makedirs(hist_dir, exist_ok=True)
    except Exception:
        return None

    ok = bool(payload.get("ok")) if isinstance(payload, dict) else False
    result = (payload or {}).get("result") if ok else None
    err = (payload or {}).get("error") if not ok else None
    failed = int((result or {}).get("failed") or 0)
    needs_stash = (not ok) or (failed > 0)
    stash_rel = "stash"
    stash_path = os.path.join(hist_dir, stash_rel)

    if needs_stash and tmp_dir and os.path.isdir(tmp_dir):
        try:
            shutil.copytree(tmp_dir, stash_path, dirs_exist_ok=True)
        except TypeError:
            if os.path.isdir(stash_path):
                shutil.rmtree(stash_path, ignore_errors=True)
            shutil.copytree(tmp_dir, stash_path)
        except Exception as e:
            logger.warning("history stash copy failed: %s", e)
            needs_stash = False
            try:
                if os.path.isdir(stash_path):
                    shutil.rmtree(stash_path, ignore_errors=True)
            except Exception:
                pass

    has_stash = needs_stash and os.path.isdir(stash_path)
    from batch_history_mysql import compute_display_title

    record = {
        "id": hid,
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "run_mode": opts.get("run_mode"),
        "options": {
            "run_mode": opts.get("run_mode"),
            "printer_name": opts.get("printer_name"),
            "copies": opts.get("copies"),
            "dry_run": opts.get("dry_run"),
            "skip_print": opts.get("skip_print"),
            "raw_print": opts.get("raw_print"),
            "accept_revisions": opts.get("accept_revisions"),
            "word_content_preserve": opts.get("word_content_preserve"),
            "word_preserve_page_count": opts.get("word_preserve_page_count"),
            "word_image_risk_guard": opts.get("word_image_risk_guard"),
            "doc_font_profile": opts.get("doc_font_profile"),
            "incremental_output_dir": opts.get("incremental_output_dir"),
            "incremental_exists_action": opts.get("incremental_exists_action"),
            "file_timeout_sec": opts.get("file_timeout_sec"),
            "word_step_timeout_sec": opts.get("word_step_timeout_sec"),
            "word_skip_file_on_timeout": opts.get("word_skip_file_on_timeout"),
        },
        "original_names": list(stream_state.get("original_names") or []),
        "payload_ok": ok,
        "error": err,
        "result": result,
        "download_token": (result or {}).get("download_token") if result else None,
        "has_stash": has_stash,
        "zip_ftp_uploaded": bool((result or {}).get("zip_ftp_uploaded"))
        if result
        else False,
        "zip_ftp_error": ((result or {}).get("zip_ftp_error") or "").strip() or None
        if result
        else None,
    }
    # 打印模式也上传“终版成品”到 FTP，并在历史记录中写入下载信息
    try:
        if ok and isinstance(result, dict):
            details = result.get("details") or []
            orig = list(stream_state.get("original_names") or [])
            # 仅在有明细时处理
            if isinstance(details, list) and details:
                from ftp_store import try_upload_file

                for i, d in enumerate(details):
                    if not isinstance(d, dict):
                        continue
                    proc = (d.get("processed_path") or "").strip()
                    if not d.get("success") or not proc or (not os.path.isfile(proc)):
                        continue
                    name = orig[i] if i < len(orig) else (d.get("filename") or d.get("path") or f"file_{i+1}")
                    arcname = _zip_arcname_with_processed_ext(str(name), proc)
                    try:
                        ftp_p, up_err = try_upload_file(proc, f"batch/final/{hid}/{arcname}")
                        if ftp_p:
                            d["final_ftp_path"] = ftp_p
                        elif up_err:
                            d["final_ftp_error"] = up_err
                    except Exception:
                        pass
    except Exception:
        pass
    record["display_title"] = compute_display_title(
        record["original_names"], result
    )
    try:
        import batch_history_mysql as bhm

        if bhm.enabled():
            bhm.save_record(record)
            keep_list = bhm.trim_to_max(_batch_history_max())
            _history_trim_folders(set(keep_list))
            return hid
    except Exception as e:
        logger.warning("batch history MySQL 写入失败，回退本地 record.json：%s", e)

    try:
        with open(os.path.join(hist_dir, "record.json"), "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning("history record write failed: %s", e)
        return None

    idx = _load_history_index()
    entries = idx.setdefault("entries", [])
    summ = _history_summary_from_record(record)
    entries.insert(0, summ)
    entries[:] = entries[: _batch_history_max()]
    keep = {e["id"] for e in entries}
    _history_trim_folders(keep)
    _save_history_index(idx)
    return hid


def _history_record_path(hid):
    if not _BATCH_EXPORT_TOKEN_RE.match(hid or ""):
        return None
    p = os.path.join(BATCH_HISTORY_ROOT, hid, "record.json")
    return p if os.path.isfile(p) else None


def _load_history_record(hid):
    try:
        import batch_history_mysql as bhm

        if bhm.enabled():
            rec = bhm.get_record(hid)
            if rec:
                return rec
    except Exception:
        pass
    rp = _history_record_path(hid)
    if not rp:
        return None
    try:
        with open(rp, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _collect_stash_files_for_retry(stash_root, original_names, details, only_failed):
    """从暂存目录收集文件；only_failed 时仅失败条目对应文档 + 全部非文档资源。"""
    failed_set = set()
    for i, d in enumerate(details or []):
        if i < len(original_names) and not d.get("success"):
            failed_set.add(str(original_names[i]).replace("\\", "/"))
    out = []
    for root, dirs, files in os.walk(stash_root):
        rel_root = os.path.relpath(root, stash_root)
        parts = [] if rel_root == "." else rel_root.replace("\\", "/").split("/")
        if "_aiprint_orig" in parts:
            continue
        for fn in files:
            abs_path = os.path.join(root, fn)
            rel = "/".join(parts + [fn]) if parts else fn
            rel = rel.replace("\\", "/")
            ext = os.path.splitext(fn)[1].lower()
            is_doc = ext in ALLOWED_EXT
            if only_failed:
                if is_doc and rel not in failed_set:
                    continue
            out.append((abs_path, rel))
    return out


def _materialize_retry_workspace(stash_root, original_names, details, only_failed):
    pairs = _collect_stash_files_for_retry(stash_root, original_names, details, only_failed)
    if not pairs:
        return None, None, None, None
    tmp_dir = tempfile.mkdtemp(prefix="aiprintword_retry_")
    try:
        for abs_p, rel in pairs:
            dest = os.path.join(tmp_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(abs_p, dest)
        doc_rels = [rel for _abs, rel in pairs if _allowed_file(rel)]
        order_map = {n: i for i, n in enumerate(original_names)}
        doc_rels_sorted = sorted(doc_rels, key=lambda n: order_map.get(n, 10**9))
        saved_paths = []
        original_names_ord = []
        original_backup_paths = []
        for rel in doc_rels_sorted:
            path = os.path.join(tmp_dir, *rel.split("/"))
            if not os.path.isfile(path):
                continue
            saved_paths.append(path)
            original_names_ord.append(rel)
            original_backup_paths.append(_write_original_backup(tmp_dir, rel, path))
        if not saved_paths:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            return None, None, None, None
        return tmp_dir, saved_paths, original_names_ord, original_backup_paths
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return None, None, None, None


def _admin_settings_auth_error():
    _sync_admin_token_from_disk()
    dotenv_path = _resolved_dotenv_path()
    exp = (os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip()
    if not exp:
        exp = _read_key_from_env_file(dotenv_path, "AIPRINTWORD_ADMIN_TOKEN")
        if exp:
            os.environ["AIPRINTWORD_ADMIN_TOKEN"] = exp
    if not exp:
        exists = os.path.isfile(dotenv_path)
        tf_default = os.path.isfile(_ADMIN_TOKEN_FILE_DEFAULT)
        logger.warning(
            "管理 API 仍无口令：ROOT=%s dotenv=%s 存在=%s 默认口令文件存在=%s cwd=%s",
            ROOT,
            dotenv_path,
            exists,
            tf_default,
            os.getcwd(),
        )
        return (
            {
                "error": (
                    "服务端未配置管理口令：请在运行服务的机器上配置 AIPRINTWORD_ADMIN_TOKEN，"
                    "或设置 AIPRINTWORD_DOTENV_PATH 指向 .env 绝对路径，或创建 data/admin_token.txt。"
                    " 部署后本响应应含 diag 字段；若仍只有短 JSON（约 138 字节），说明未加载最新 app.py。"
                ),
                "diag": {
                    "auth_api_build": AIPRINTWORD_WEB_BUILD,
                    "app_root": ROOT,
                    "dotenv_path": dotenv_path,
                    "dotenv_exists": exists,
                    "cwd": os.getcwd(),
                    "admin_token_file_default": _ADMIN_TOKEN_FILE_DEFAULT,
                    "admin_token_file_exists": tf_default,
                },
            },
            503,
        )
    tok = (request.headers.get("X-Admin-Token") or request.args.get("token") or "").strip()
    if tok != exp:
        if not tok:
            return (
                "未授权：未携带管理口令。请在设置页「管理口令」输入与服务器 .env 中 "
                "AIPRINTWORD_ADMIN_TOKEN 完全相同的值后再点「加载」（请求头 X-Admin-Token）。",
                401,
            )
        return (
            "未授权：口令与服务器不一致。请核对页面输入与 .env 中 AIPRINTWORD_ADMIN_TOKEN "
            "是否完全一致（区分大小写、首尾勿多空格）。",
            401,
        )
    return None, None


@app.route("/settings")
def settings_page():
    return send_from_directory(os.path.join(ROOT, "static"), "settings.html")


@app.route("/api/aiprintword-build")
def api_aiprintword_build():
    """
    无鉴权：确认当前进程实际加载的 app.py 路径与构建号。
    若此处 build>=4 而 /api/admin/settings 仍为 138 字节 503，则几乎不可能（请清 CDN/代理缓存或核对是否多台服务）。
    """
    return jsonify(
        {
            "ok": True,
            "build": AIPRINTWORD_WEB_BUILD,
            "app_py": os.path.abspath(__file__),
            "resolved_dotenv": _resolved_dotenv_path(),
        }
    )


@app.route("/api/admin/settings", methods=["GET"])
def api_admin_settings_get():
    err, code = _admin_settings_auth_error()
    if err:
        if isinstance(err, dict):
            return jsonify({"ok": False, **err}), code
        return jsonify({"ok": False, "error": err}), code
    from runtime_settings.resolve import list_all_settings

    return jsonify({"ok": True, "items": list_all_settings(mask_secrets=True)})


@app.route("/api/admin/settings", methods=["PUT"])
def api_admin_settings_put():
    err, code = _admin_settings_auth_error()
    if err:
        if isinstance(err, dict):
            return jsonify({"ok": False, **err}), code
        return jsonify({"ok": False, "error": err}), code
    try:
        body = request.get_json(force=True, silent=True) or {}
    except Exception:
        body = {}
    if not isinstance(body, dict):
        return jsonify({"ok": False, "error": "请求体须为 JSON 对象"}), 400
    updates = {}
    for k, v in body.items():
        ks = str(k).strip()
        if not ks:
            continue
        if v is None:
            updates[ks] = ""
        elif isinstance(v, bool):
            updates[ks] = "1" if v else "0"
        else:
            updates[ks] = str(v).strip()
    try:
        from runtime_settings.resolve import set_settings

        changed = set_settings(updates)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except RuntimeError as e:
        return jsonify({"ok": False, "error": str(e)}), 503
    except Exception as e:
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500
    try:
        from doc_handlers.word_handler import apply_resolved_word_base_settings

        apply_resolved_word_base_settings()
    except Exception:
        pass
    return jsonify({"ok": True, "updated": changed})


@app.route("/api/admin/batch-history/migrate-from-disk", methods=["POST"])
def api_admin_batch_history_migrate_from_disk():
    """将 data/batch_history/<id>/record.json 中尚未入库的记录导入 MySQL，并补全 display_title。"""
    err, code = _admin_settings_auth_error()
    if err:
        if isinstance(err, dict):
            return jsonify({"ok": False, **err}), code
        return jsonify({"ok": False, "error": err}), code
    import batch_history_mysql as bhm

    if not bhm.enabled():
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "未配置 MySQL（MYSQL_HOST），无法写入历史库",
                }
            ),
            400,
        )
    try:
        st = bhm.migrate_from_disk(BATCH_HISTORY_ROOT, _batch_history_max())
        backfilled = bhm.backfill_display_titles()
        return jsonify({"ok": True, **st, "backfilled": backfilled})
    except Exception as e:
        logger.exception("batch history migrate-from-disk failed")
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500


@app.route("/api/admin/sign/migrate-mysql-blobs-to-ftp", methods=["POST"])
def api_admin_sign_migrate_mysql_blobs_to_ftp():
    """将在线签名模块历史遗留的 MySQL BLOB 文件迁移到 FTP。"""
    err, code = _admin_settings_auth_error()
    if err:
        if isinstance(err, dict):
            return jsonify({"ok": False, **err}), code
        return jsonify({"ok": False, "error": err}), code
    try:
        from sign_handlers import mysql_store

        if not mysql_store.mysql_sign_enabled():
            return jsonify({"ok": False, "error": "未配置 MySQL（MYSQL_HOST），无需迁移或无法读取"}), 400
        data = request.get_json(silent=True) or {}
        batch_size = int(data.get("batch_size", 2000) or 2000)
        max_total = int(data.get("max_total", 200000) or 200000)
        clear_blob = bool(data.get("clear_blob", True))
        st = mysql_store.migrate_mysql_blobs_to_ftp(
            batch_size=batch_size, max_total=max_total, clear_blob=clear_blob
        )
        st2 = mysql_store.migrate_signer_strokes_blobs_to_ftp(
            batch_size=batch_size, max_total=max_total, clear_blob=clear_blob
        )
        st3 = mysql_store.migrate_stroke_items_blobs_to_ftp(
            batch_size=batch_size, max_total=max_total, clear_blob=clear_blob
        )
        st4 = mysql_store.verify_and_backfill_ftp_files(limit=batch_size)
        return jsonify({"ok": True, "files": st, "legacy_strokes": st2, "stroke_items": st3, "verify": st4})
    except Exception as e:
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500


@app.route("/api/batch-export/<token>")
def api_batch_export(token):
    """?????????? ZIP?token ???????????"""
    if not _BATCH_EXPORT_TOKEN_RE.match(token or ""):
        return jsonify({"ok": False, "error": "invalid token"}), 404
    path = os.path.join(BATCH_EXPORT_ROOT, token + ".zip")
    if not os.path.isfile(path):
        try:
            from ftp_store import download_bytes

            b = download_bytes(f"batch/exports/{token}.zip")
            os.makedirs(BATCH_EXPORT_ROOT, exist_ok=True)
            with open(path, "wb") as fp:
                fp.write(b)
        except Exception:
            return jsonify({"ok": False, "error": "not found or expired"}), 404
    return send_file(
        path,
        as_attachment=True,
        download_name="processed_documents.zip",
        max_age=0,
    )


def _run_batch_with_progress(
    saved_paths, original_names, original_backup_paths, opts, queue
):
    """Run batch in background and push progress to queue."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    job = opts.get("_job") if isinstance(opts, dict) else None
    cancel_event = getattr(job, "cancel_event", None) if job else None
    pause_event = getattr(job, "pause_event", None) if job else None
    job_id = getattr(job, "job_id", None) if job else None
    skip_current_event = getattr(job, "skip_current_event", None) if job else None

    try:
        from batch_print import run_batch, build_batch_modification_zip_text, _format_duration_cn

        def progress_callback(
            step,
            file_index,
            file_total,
            file_name,
            percent=None,
            eta=None,
            processing_meta=None,
        ):
            name = original_names[file_index - 1] if file_index <= len(original_names) else file_name
            eta_log = ""
            if isinstance(eta, dict) and eta.get("eta_remaining_sec") is not None:
                eta_log = " eta_rem≈%s avg=%ss" % (
                    _format_duration_cn(eta["eta_remaining_sec"]),
                    eta.get("avg_sec_per_file", "?"),
                )
            mode_log = ""
            if isinstance(processing_meta, dict) and processing_meta.get("processingModeLabel"):
                mode_log = " mode=%s" % (processing_meta.get("processingModeLabel"),)
            logger.info(
                "progress idx=%s/%s pct=%s step=%s file=%s%s%s",
                file_index,
                file_total,
                percent,
                step,
                name,
                eta_log,
                mode_log,
            )
            payload = {
                "step": step,
                "fileIndex": file_index,
                "fileTotal": file_total,
                "fileName": name,
                "percent": percent,
            }
            if isinstance(eta, dict):
                payload["eta"] = eta
            if isinstance(processing_meta, dict):
                if processing_meta.get("processingMode") is not None:
                    payload["processingMode"] = processing_meta["processingMode"]
                if processing_meta.get("processingModeLabel"):
                    payload["processingModeLabel"] = processing_meta["processingModeLabel"]
                pts = processing_meta.get("modificationPoints")
                if pts:
                    payload["modificationPoints"] = pts
            queue.put(("progress", payload))

        logger.info(
            "batch worker start total=%s mode=%s dry_run=%s skip_print=%s raw_print=%s check_formal=%s check_signature=%s",
            len(saved_paths),
            opts.get("run_mode"),
            opts.get("dry_run"),
            opts.get("skip_print"),
            opts.get("raw_print"),
            opts.get("check_formal"),
            opts.get("check_signature"),
        )
        result = run_batch(
            saved_paths,
            recursive=False,
            check_formal=opts["check_formal"],
            check_signature=opts["check_signature"],
            accept_revisions=opts["accept_revisions"],
            word_content_preserve=opts["word_content_preserve"],
            word_preserve_page_count=opts["word_preserve_page_count"],
            word_image_risk_guard=opts["word_image_risk_guard"],
            word_step_timeout_sec=opts.get("word_step_timeout_sec"),
            word_skip_file_on_timeout=opts.get("word_skip_file_on_timeout"),
            file_timeout_sec=opts.get("file_timeout_sec"),
            word_font_profile=opts.get("doc_font_profile", "mixed"),
            printer_name=opts["printer_name"],
            copies=opts["copies"],
            dry_run=opts["dry_run"],
            skip_print=opts["skip_print"],
            raw_print=opts["raw_print"],
            progress_callback=progress_callback,
            # cancel/pause 来自 job
            cancel_event=cancel_event,
            pause_event=pause_event,
            incremental_output_dir=opts.get("incremental_output_dir"),
            relative_names=original_names,
            incremental_exists_action=opts.get("incremental_exists_action", "overwrite"),
            skip_current_event=skip_current_event,
        )
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        logger.info(
            "batch worker done total=%s ok=%s failed=%s elapsed=%ss summary=%s",
            result.get("total"),
            result.get("ok"),
            result.get("failed"),
            result.get("batch_elapsed_sec"),
            result.get("eta_summary"),
        )
        if opts.get("skip_print") and result.get("total"):
            try:
                token = uuid.uuid4().hex
                # 取消时仅打包已完成的文件（details 中已有顺序）；未完成的不纳入。
                done_n = len(result.get("details") or [])
                pack_paths = saved_paths[:done_n]
                pack_names = original_names[:done_n]
                pack_backups = original_backup_paths[:done_n]
                mod_txt = build_batch_modification_zip_text(result["details"])
                _, zip_ftp_ok, zip_ftp_err = _zip_batch_exports(
                    result["details"],
                    pack_paths,
                    pack_names,
                    pack_backups,
                    token,
                    modification_report_text=mod_txt,
                )
                result["download_token"] = token
                result["download_filename"] = "processed_documents.zip"
                result["zip_ftp_uploaded"] = bool(zip_ftp_ok)
                if zip_ftp_err:
                    result["zip_ftp_error"] = zip_ftp_err
            except Exception as e:
                logger.exception("batch zip failed: %s", e)
                queue.put(("result", {"ok": False, "error": _format_com_error(e)}))
                return
        # 本轮（local）也提供终版下载：上传成功项终版到 FTP 并下发 final_download_key
        try:
            details = result.get("details") or []
            if isinstance(details, list) and details:
                key = uuid.uuid4().hex
                from ftp_store import try_upload_file

                items = []
                for i, d in enumerate(details):
                    if not isinstance(d, dict):
                        continue
                    proc = (d.get("processed_path") or "").strip()
                    if not d.get("success") or not proc or (not os.path.isfile(proc)):
                        continue
                    name = original_names[i] if i < len(original_names) else (d.get("filename") or d.get("path") or f"file_{i+1}")
                    arcname = _zip_arcname_with_processed_ext(str(name), proc)
                    try:
                        ftp_p, up_err = try_upload_file(proc, f"batch/final/{key}/{arcname}")
                        if ftp_p:
                            d["final_ftp_path"] = ftp_p
                            items.append({"ftp_path": ftp_p, "filename": os.path.basename(arcname) or f"file_{i+1}"})
                        elif up_err:
                            d["final_ftp_error"] = up_err
                    except Exception:
                        pass
                if items:
                    _FINAL_DOWNLOAD_CACHE[key] = {"created_at": time.time(), "items": items}
                    result["final_download_key"] = key
        except Exception:
            pass
        queue.put(("result", {"ok": True, "result": result}))
    except Exception as e:
        logger.exception("batch worker failed: %s", e)
        queue.put(("result", {"ok": False, "error": _format_com_error(e)}))
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        queue.put((None, None))  # ????

        if job and job_id:
            try:
                _BATCH_JOB_REGISTRY.pop(job_id, None)
            except Exception:
                pass


def _sse_message(event, data):
    """Build one SSE message."""
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _sse_batch_stream_body(queue, tmp_dir, stream_state_ref, heartbeat_sec=120.0):
    """从队列读事件并 yield SSE；结束时持久化历史并删除 tmp。heartbeat_sec：无进度时推送间隔（秒）。"""
    hb = max(10.0, min(600.0, float(heartbeat_sec)))
    started_at = time.time()
    last_emit = time.time()
    try:
        while True:
            try:
                event_type, payload = queue.get(timeout=hb)
            except Empty:
                waited = int(time.time() - last_emit)
                yield _sse_message(
                    "heartbeat",
                    {
                        "ts": int(time.time()),
                        "heartbeatSec": int(hb),
                        "message": f"仍在处理中（已 {waited} 秒无进度推送，心跳间隔 {int(hb)} 秒）",
                        "runningForSec": int(time.time() - started_at),
                    },
                )
                continue
            if event_type is None:
                break
            if event_type == "progress":
                last_emit = time.time()
                yield _sse_message("progress", payload)
            elif event_type == "result":
                last_emit = time.time()
                stream_state_ref["final_json"] = payload
                yield _sse_message("result", payload)
                break
    finally:
        try:
            _batch_history_persist(stream_state_ref, tmp_dir)
        except Exception:
            logger.exception("batch history persist error")
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass


@app.route("/api/batch-control/<job_id>", methods=["POST"])
def api_batch_control(job_id):
    """暂停/继续/取消当前批处理任务。"""
    if not _BATCH_EXPORT_TOKEN_RE.match(job_id or ""):
        return jsonify({"ok": False, "error": "invalid job id"}), 404
    job = _BATCH_JOB_REGISTRY.get(job_id)
    if not job:
        return jsonify({"ok": False, "error": "job not found or finished"}), 404
    # 不能写成 form.get or json.get if is_json else ''：multipart 时 is_json 为 False，整式会变成 ''，导致永远 invalid action
    if request.is_json:
        raw = (request.get_json(silent=True) or {}).get("action")
    else:
        raw = request.form.get("action")
    action = str(raw or "").strip().lower()
    if action == "cancel":
        job.cancel_event.set()
        job.pause_event.clear()
        return jsonify({"ok": True, "jobId": job_id, "state": "cancelling"})
    if action == "pause":
        job.pause_event.set()
        return jsonify({"ok": True, "jobId": job_id, "state": "paused"})
    if action in ("resume", "continue"):
        job.pause_event.clear()
        return jsonify({"ok": True, "jobId": job_id, "state": "running"})
    if action in ("skip_current", "skip_file"):
        job.skip_current_event.set()
        return jsonify({"ok": True, "jobId": job_id, "state": "skip_requested"})
    return jsonify({"ok": False, "error": "invalid action"}), 400


@app.route("/api/batch-print-stream", methods=["POST"])
def api_batch_print_stream():
    """
    Streaming batch print via text/event-stream.
    progress: { step, fileIndex, fileTotal, fileName, percent, eta?, processingMode?, processingModeLabel?, modificationPoints? }
    result: { ok, result? or error? }
    """
    files = request.files.getlist("files") or request.files.getlist("files[]")
    if not files or not any(f.filename for f in files):
        return jsonify({"ok": False, "error": "?????????"}), 400

    try:
        opts = _parse_batch_opts_from_form(request.form)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        original_backup_paths = []
        for f in files:
            rel = _upload_relpath(f.filename)
            if not rel:
                continue
            path = os.path.join(tmp_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            f.save(path)
            if _is_archive_upload(rel):
                for p2, r2 in _expand_archive_for_batch(tmp_dir, path, rel):
                    saved_paths.append(p2)
                    original_names.append(r2)
                    bk = _write_original_backup(tmp_dir, r2, p2)
                    original_backup_paths.append(bk)
            elif _allowed_file(rel):
                saved_paths.append(path)
                original_names.append(rel)
                bk = _write_original_backup(tmp_dir, rel, path)
                original_backup_paths.append(bk)

        if not saved_paths:
            return jsonify(
                {
                    "ok": False,
                    "error": "未找到可处理文件：压缩包内需含 .doc/.docx/.xls/.xlsx/.pdf 等文档；ZIP 异常时可安装 7-Zip 后重试；.7z/.rar 需 7-Zip 或 py7zr/UnRAR。",
                }
            ), 400
        logger.info("stream request accepted files=%s", len(saved_paths))

        job_id = uuid.uuid4().hex
        job = _BatchJob(job_id=job_id, cancel_event=threading.Event(), pause_event=threading.Event())
        _BATCH_JOB_REGISTRY[job_id] = job

        queue = Queue()
        opts["_job"] = job
        thread = threading.Thread(
            target=_run_batch_with_progress,
            args=(saved_paths, original_names, original_backup_paths, opts, queue),
            daemon=True,
        )
        thread.start()

        stream_state = {
            "opts": {k: v for k, v in opts.items() if not str(k).startswith("_")},
            "saved_paths": saved_paths,
            "original_names": original_names,
            "original_backup_paths": original_backup_paths,
            "final_json": None,
        }
        heartbeat_sec = _parse_sse_heartbeat_sec(request.form)
        logger.info("batch-print-stream sse heartbeat interval=%ss", int(heartbeat_sec))

        def generate_with_job():
            yield _sse_message("job", {"jobId": job_id})
            yield from _sse_batch_stream_body(
                queue, tmp_dir, stream_state, heartbeat_sec=heartbeat_sec
            )

        return Response(
            generate_with_job(),
            mimetype="text/event-stream",
            headers={
                "Cache-Control": "no-cache",
                "X-Accel-Buffering": "no",
                "Connection": "keep-alive",
            },
        )
    except Exception as e:
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500


def _history_entries_ensure_title(entries):
    """旧版 index 摘要无 title 时，按需加载 record 补全。"""
    from batch_history_mysql import compute_display_title

    out = []
    for e in entries:
        if not isinstance(e, dict):
            continue
        if e.get("title"):
            # 即便已有 title，也可能存在 zip_ftp/zip_ftp_error 字段过期（例如迁移脚本更新了 record.json）
            e2 = dict(e)
            hid = e2.get("id")
            rec = _load_history_record(hid) if hid else None
            if rec:
                try:
                    r = rec.get("result") if isinstance(rec, dict) else None
                    r = r if isinstance(r, dict) else {}
                    zup = r.get("zip_ftp_uploaded")
                    if zup is None:
                        zup = rec.get("zip_ftp_uploaded") if isinstance(rec, dict) else None
                    if zup is True or zup == 1:
                        e2["zip_ftp"] = True
                    elif zup is False or zup == 0:
                        e2["zip_ftp"] = False
                    # 只在 record 有明确值时覆盖 error，避免把前端已有提示清空
                    ze = (r.get("zip_ftp_error") or rec.get("zip_ftp_error") or "").strip() if isinstance(rec, dict) else ""
                    if ze:
                        e2["zip_ftp_error"] = ze
                    else:
                        # 如果 record 里明确为空，则清空
                        if "zip_ftp_error" in e2:
                            e2["zip_ftp_error"] = None
                except Exception:
                    pass
            out.append(e2)
            continue
        e2 = dict(e)
        hid = e2.get("id")
        rec = _load_history_record(hid) if hid else None
        if rec:
            e2["title"] = rec.get("display_title") or compute_display_title(
                rec.get("original_names"), rec.get("result")
            )
            # 补全 zip_ftp/zip_ftp_error（从 record 派生，避免 index.json 过期）
            try:
                r = rec.get("result") if isinstance(rec, dict) else None
                r = r if isinstance(r, dict) else {}
                zup = r.get("zip_ftp_uploaded")
                if zup is None:
                    zup = rec.get("zip_ftp_uploaded") if isinstance(rec, dict) else None
                if zup is True or zup == 1:
                    e2["zip_ftp"] = True
                elif zup is False or zup == 0:
                    e2["zip_ftp"] = False
                ze = (r.get("zip_ftp_error") or rec.get("zip_ftp_error") or "").strip() if isinstance(rec, dict) else ""
                e2["zip_ftp_error"] = ze or None
            except Exception:
                pass
        else:
            rm = e2.get("run_mode") or "批处理"
            tot = int(e2.get("total") or 0)
            n = tot if tot > 0 else 1
            e2["title"] = f"{rm}等{n}份文件处理记录"
        out.append(e2)
    return out


@app.route("/api/batch-history", methods=["GET"])
def api_batch_history_list():
    idx = _load_history_index()
    entries = _history_entries_ensure_title(idx.get("entries", []))
    return jsonify({"ok": True, "entries": entries})


@app.route("/api/batch-history/<hid>", methods=["GET"])
def api_batch_history_get(hid):
    rec = _load_history_record(hid)
    if not rec:
        return jsonify({"ok": False, "error": "记录不存在"}), 404
    return jsonify({"ok": True, "record": rec})


@app.route("/api/batch-history/<hid>/download", methods=["GET"])
def api_batch_history_download(hid):
    rec = _load_history_record(hid)
    if not rec:
        return jsonify({"ok": False, "error": "记录不存在"}), 404
    token = rec.get("download_token")
    if not token or not _BATCH_EXPORT_TOKEN_RE.match(str(token)):
        return jsonify({"ok": False, "error": "该次任务无打包下载"}), 404
    path = os.path.join(BATCH_EXPORT_ROOT, str(token) + ".zip")
    if not os.path.isfile(path):
        try:
            from ftp_store import download_bytes

            b = download_bytes(f"batch/exports/{token}.zip")
            os.makedirs(BATCH_EXPORT_ROOT, exist_ok=True)
            with open(path, "wb") as fp:
                fp.write(b)
        except Exception:
            return jsonify({"ok": False, "error": "文件已删除或已过期"}), 404
    return send_file(
        path,
        as_attachment=True,
        download_name="processed_documents.zip",
        max_age=0,
    )


@app.route("/api/batch-history/<hid>/final/<int:idx>", methods=["GET"])
def api_batch_history_final_download(hid, idx: int):
    """下载某次历史任务的单文件终版成品（来自 FTP）。"""
    rec = _load_history_record(hid)
    if not rec:
        return jsonify({"ok": False, "error": "记录不存在"}), 404
    res = (rec.get("result") or {}) if isinstance(rec, dict) else {}
    details = res.get("details") or []
    if not isinstance(details, list) or idx < 0 or idx >= len(details):
        return jsonify({"ok": False, "error": "无效的文件索引"}), 404
    d = details[idx] if isinstance(details[idx], dict) else {}
    ftp_p = (d.get("final_ftp_path") or "").strip()
    if not ftp_p:
        return jsonify({"ok": False, "error": "该文件无可下载终版"}), 404
    try:
        from ftp_store import download_bytes

        b = download_bytes(ftp_p)
    except Exception as e:
        return jsonify({"ok": False, "error": "FTP 下载失败"}), 502
    # 下载文件名：优先 filename/path，再退回 idx
    nm = (d.get("filename") or d.get("path") or f"file_{idx+1}").replace("\\", "/")
    base = os.path.basename(nm) or f"file_{idx+1}"
    return send_file(
        io.BytesIO(b),
        as_attachment=True,
        download_name=base,
        mimetype="application/octet-stream",
        max_age=0,
    )


@app.route("/api/batch-final/<key>/<int:idx>", methods=["GET"])
def api_batch_final_download_local(key: str, idx: int):
    """下载本轮（local）终版成品：通过内存 key 映射到 FTP 路径。"""
    if not _BATCH_EXPORT_TOKEN_RE.match(key or ""):
        return jsonify({"ok": False, "error": "invalid key"}), 404
    rec = _FINAL_DOWNLOAD_CACHE.get(key)
    if not rec:
        return jsonify({"ok": False, "error": "该下载已过期，请刷新历史列表后从历史记录下载"}), 404
    items = rec.get("items") or []
    if idx < 0 or idx >= len(items):
        return jsonify({"ok": False, "error": "无效的文件索引"}), 404
    it = items[idx] if isinstance(items[idx], dict) else {}
    ftp_p = (it.get("ftp_path") or "").strip()
    fn = (it.get("filename") or f"file_{idx+1}").strip()
    if not ftp_p:
        return jsonify({"ok": False, "error": "无可下载终版"}), 404
    try:
        from ftp_store import download_bytes

        b = download_bytes(ftp_p)
    except Exception:
        return jsonify({"ok": False, "error": "FTP 下载失败"}), 502
    return send_file(
        io.BytesIO(b),
        as_attachment=True,
        download_name=fn,
        mimetype="application/octet-stream",
        max_age=0,
    )


@app.route("/api/batch-history/<hid>/retry", methods=["POST"])
def api_batch_history_retry(hid):
    from werkzeug.datastructures import CombinedMultiDict

    rec = _load_history_record(hid)
    if not rec:
        return jsonify({"ok": False, "error": "记录不存在"}), 404
    if not rec.get("has_stash"):
        return jsonify({"ok": False, "error": "无可用暂存，请重新上传后再处理"}), 400
    stash = os.path.join(BATCH_HISTORY_ROOT, hid, "stash")
    if not os.path.isdir(stash):
        return jsonify({"ok": False, "error": "暂存目录已丢失"}), 400
    only_failed = (request.form.get("only_failed") or "true").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    defaults = _record_to_default_form(rec)
    combined = CombinedMultiDict([request.form, defaults])
    try:
        opts = _parse_batch_opts_from_form(combined)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400

    original_names = rec.get("original_names") or []
    details = (rec.get("result") or {}).get("details") or []
    tmp_dir, saved_paths, original_names_ord, original_backup_paths = _materialize_retry_workspace(
        stash, original_names, details, only_failed
    )
    if not tmp_dir or not saved_paths:
        return jsonify({"ok": False, "error": "未能从暂存中恢复待重试文件"}), 400

    logger.info(
        "history retry hid=%s only_failed=%s docs=%s",
        hid,
        only_failed,
        len(saved_paths),
    )
    job_id = uuid.uuid4().hex
    job = _BatchJob(job_id=job_id, cancel_event=threading.Event(), pause_event=threading.Event())
    _BATCH_JOB_REGISTRY[job_id] = job
    queue = Queue()
    opts["_job"] = job
    thread = threading.Thread(
        target=_run_batch_with_progress,
        args=(saved_paths, original_names_ord, original_backup_paths, opts, queue),
        daemon=True,
    )
    thread.start()

    stream_state = {
        "opts": {k: v for k, v in opts.items() if not str(k).startswith("_")},
        "saved_paths": saved_paths,
        "original_names": original_names_ord,
        "original_backup_paths": original_backup_paths,
        "final_json": None,
    }
    heartbeat_sec = _parse_sse_heartbeat_sec(combined)

    def generate_with_job():
        yield _sse_message("job", {"jobId": job_id})
        yield from _sse_batch_stream_body(
            queue, tmp_dir, stream_state, heartbeat_sec=heartbeat_sec
        )

    return Response(
        generate_with_job(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


# ---------- ????????????????????----------
SIGN_ALLOWED_EXT = {".docx", ".xlsx"}
SIGN_INBOX_ROOT = os.path.join(ROOT, "data", "sign_inbox")
SIGN_MAX_SAVED_FILES = 50
_SIGN_FILE_ID_RE = re.compile(r"^[0-9a-f]{32}$")


def _sign_using_mysql():
    """Use MySQL store when MYSQL_HOST is configured."""
    try:
        from sign_handlers import mysql_store

        return mysql_store.mysql_sign_enabled()
    except Exception:
        return False


def _sign_ensure_session_inbox():
    """Ensure per-session inbox directory exists."""
    if "sign_inbox_sid" not in session:
        session["sign_inbox_sid"] = uuid.uuid4().hex
    if "sign_files" not in session:
        session["sign_files"] = []
    sid = session["sign_inbox_sid"]
    inbox_dir = os.path.join(SIGN_INBOX_ROOT, sid)
    os.makedirs(inbox_dir, exist_ok=True)
    return sid, inbox_dir


def _sign_saved_disk_path(sid: str, file_id: str, ext: str) -> str:
    return os.path.join(SIGN_INBOX_ROOT, sid, file_id + ext.lower())


def _sign_upload_display_name(client_filename):
    """
    ??multipart ????????????????????????????    ??.docx/.xlsx ?? (None, None)??    """
    if not client_filename:
        return None, None
    norm = str(client_filename).replace("\\", "/").strip()
    if ".." in norm:
        norm = os.path.basename(norm)
    parts = [p for p in norm.split("/") if p and p not in (".", "..")]
    if not parts:
        return None, None
    last = parts[-1]
    ext = os.path.splitext(last)[1].lower()
    if ext not in SIGN_ALLOWED_EXT:
        return None, None
    display = "/".join(parts)
    if len(display) > 200:
        display = display[-200:]
    return display, ext


def _safe_display_filename_keep_unicode(name: str) -> str:
    """
    保留中文等非 ASCII 字符的“安全显示文件名”净化。
    仅替换 Windows 非法文件名字符，避免 secure_filename 把中文清空。
    """
    s = str(name or "").strip()
    if not s:
        return "document"
    # 仅取最后一段，避免路径穿越/目录显示
    s = os.path.basename(s.replace("\\", "/"))
    # Windows 非法字符：<>:"/\\|?* 以及控制字符
    s = re.sub(r"[<>:\"/\\\\|?*]+", "_", s)
    s = re.sub(r"[\x00-\x1f]+", "_", s)
    # 尾部点/空格在 Windows 里会被裁剪；这里做个兜底
    s = s.rstrip(". ")
    return s or "document"


def _sign_find_record(file_id: str):
    fid = str(file_id or "")
    for rec in session.get("sign_files") or []:
        if str(rec.get("id") or "") == fid:
            return rec
    return None


def _sign_norm_ext(ext: Optional[str]) -> str:
    e = (ext or ".docx").lower().strip()
    if not e.startswith("."):
        e = "." + e
    return e if e in SIGN_ALLOWED_EXT else ".docx"


def _sign_saved_file_exists(sid: str, file_id: str, ext: Optional[str]) -> bool:
    """会话收件箱中该 id 是否仍有对应磁盘文件（兼容扩展名与记录不一致）。"""
    fid = str(file_id or "").strip()
    if not fid or not _SIGN_FILE_ID_RE.match(fid):
        return False
    primary = _sign_norm_ext(ext)
    path = _sign_saved_disk_path(sid, fid, primary)
    if os.path.isfile(path):
        return True
    for alt in SIGN_ALLOWED_EXT:
        if alt == primary:
            continue
        p2 = _sign_saved_disk_path(sid, fid, alt)
        if os.path.isfile(p2):
            return True
    return False


def _sign_prune_session_files_to_disk(sid: str) -> list:
    """
    以磁盘为准修剪 session['sign_files']：已删文件但 Cookie 未写回时，刷新列表仍会从库里删掉条目。
    """
    records = list(session.get("sign_files") or [])
    pruned = [
        r
        for r in records
        if _sign_saved_file_exists(sid, str(r.get("id") or ""), r.get("ext"))
    ]
    if len(pruned) != len(records):
        session["sign_files"] = pruned
        session.modified = True
    return pruned


def _sign_remove_disk_files_for_id(sid: str, file_id: str, ext: Optional[str]) -> None:
    """删除收件箱内该 file_id 可能存在的 .docx/.xlsx（避免扩展名记录与磁盘不一致导致删不干净）。"""
    fid = str(file_id or "").strip()
    if not fid:
        return
    tried = set()
    for e in (_sign_norm_ext(ext), ".docx", ".xlsx"):
        if e in tried:
            continue
        tried.add(e)
        path = _sign_saved_disk_path(sid, fid, e)
        if os.path.isfile(path):
            try:
                os.remove(path)
            except OSError:
                pass


def _decode_png_data_url_or_b64(s):
    """Decode canvas PNG data URL or raw base64 string."""
    if not s or not str(s).strip():
        return None
    s = str(s).strip()
    if s.startswith("data:image"):
        parts = s.split(",", 1)
        s = parts[1] if len(parts) > 1 else ""
    try:
        return base64.b64decode(s, validate=False)
    except Exception:
        try:
            return base64.b64decode(s)
        except Exception:
            return None


def _sign_process_document_bytes(
    file_bytes: bytes,
    ext: str,
    base_name: str,
    roles: list,
    sig_map: dict,
    date_map: dict,
    source_file_id: Optional[str],
    batch_id: Optional[str] = None,
) -> dict:
    """生成已签名文档字节；MySQL 模式下同时写入 sign_signed_output。返回 dict 含 ok / error / out_bytes 等。"""
    tmp_dir = tempfile.mkdtemp(prefix="aiprintword_sign_")
    try:
        # 临时文件名需兼容 Windows/文件系统；显示名称/记录名称保留原始（可含中文）
        safe_in_name = secure_filename(os.path.basename(base_name or "")) or ("document" + ext)
        if not safe_in_name.lower().endswith(ext):
            safe_in_name = os.path.splitext(safe_in_name)[0] + ext
        in_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex[:8]}_{safe_in_name}")
        with open(in_path, "wb") as fp:
            fp.write(file_bytes)
        from sign_handlers import sign_document

        out_path = sign_document(in_path, sig_map, date_map)
        dl_name = os.path.splitext(os.path.basename(base_name or "document"))[0] + "_signed" + ext
        with open(out_path, "rb") as fp:
            out_bytes = fp.read()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        tmp_dir = None
        signed_row_id = None
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            signed_row_id = uuid.uuid4().hex
            src_id = (
                source_file_id
                if source_file_id and _SIGN_FILE_ID_RE.match(source_file_id or "")
                else None
            )
            mysql_store.insert_signed_output(
                signed_row_id,
                batch_id,
                src_id,
                base_name,
                dl_name,
                ext,
                json.dumps(roles, ensure_ascii=False),
                out_bytes,
            )
        return {
            "ok": True,
            "out_bytes": out_bytes,
            "dl_name": dl_name,
            "signed_id": signed_row_id,
        }
    except Exception as e:
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
        return {"ok": False, "error": str(e)}


@app.route("/sign")
def sign_page():
    """Serve online signature page (default: file signing)."""
    return send_from_directory(os.path.join(ROOT, "static"), "sign_file.html")


@app.route("/sign/materials")
def sign_materials_page():
    """Serve online signature materials page (stroke library input)."""
    return send_from_directory(os.path.join(ROOT, "static"), "sign_materials.html")


@app.route("/api/sign/files", methods=["GET"])
def api_sign_files_list():
    """List pending files for signature."""
    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            files = mysql_store.list_files()
            return jsonify({"ok": True, "files": files})
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL ????: {e}"}), 500
    _sign_ensure_session_inbox()
    sid = session["sign_inbox_sid"]
    pruned = _sign_prune_session_files_to_disk(sid)
    files = []
    for rec in pruned:
        files.append({"id": rec.get("id"), "name": rec.get("name"), "ext": rec.get("ext")})
    return jsonify({"ok": True, "files": files})


@app.route("/api/sign/upload", methods=["POST"])
def api_sign_upload():
    """Save one or multiple .docx/.xlsx files."""
    uploads = request.files.getlist("files")
    if not uploads or not any(f.filename for f in uploads):
        one = request.files.get("file")
        if one and one.filename:
            uploads = [one]
    if not uploads or not any(f.filename for f in uploads):
        return jsonify({"ok": False, "error": "?????"}), 400

    parsed = []
    for upload in uploads:
        if not upload.filename:
            continue
        display_name, ext = _sign_upload_display_name(upload.filename)
        if not display_name:
            continue
        parsed.append((upload, display_name, ext))

    if not parsed:
        return jsonify(
            {
                "ok": False,
                "error": "???????????? .docx / .xlsx?????????????????",
            }
        ), 400

    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            cur_n = mysql_store.count_files()
            _max_f = _sign_mysql_max_files()
            if cur_n + len(parsed) > _max_f:
                return jsonify(
                    {
                        "ok": False,
                        "error": (
                            f"???? {cur_n} ?????? {len(parsed)} ??"
                            f"????????{_max_f} ????????????"
                        ),
                    }
                ), 400
            last_rec = None
            for upload, display_name, ext in parsed:
                file_id = uuid.uuid4().hex
                raw = upload.read()
                if not raw:
                    return jsonify({"ok": False, "error": f"?????{display_name}"}), 400
                mysql_store.insert_file(file_id, display_name, ext, raw)
                last_rec = {"id": file_id, "name": display_name, "ext": ext}
            return jsonify(
                {
                    "ok": True,
                    "file": last_rec,
                    "files": mysql_store.list_files(),
                    "added": len(parsed),
                }
            )
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL ????: {e}"}), 500

    sid, inbox_dir = _sign_ensure_session_inbox()
    records = list(session.get("sign_files") or [])
    if len(records) + len(parsed) > SIGN_MAX_SAVED_FILES:
        return jsonify(
            {
                "ok": False,
                "error": (
                    f"???? {len(records)} ?????? {len(parsed)} ??"
                    f"????????{SIGN_MAX_SAVED_FILES} ????????????"
                ),
            }
        ), 400

    last_rec = None
    for upload, display_name, ext in parsed:
        file_id = uuid.uuid4().hex
        dest = os.path.join(inbox_dir, file_id + ext)
        upload.stream.seek(0)
        upload.save(dest)
        last_rec = {"id": file_id, "name": display_name, "ext": ext}
        records.append(last_rec)
    session["sign_files"] = records
    session.modified = True

    return jsonify({"ok": True, "file": last_rec, "files": records, "added": len(parsed)})


@app.route("/api/sign/files/<file_id>", methods=["DELETE"])
def api_sign_file_delete(file_id):
    """Delete one file record from pending list."""
    if not _SIGN_FILE_ID_RE.match(file_id or ""):
        return jsonify({"ok": False, "error": "??????id"}), 400
    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            n = mysql_store.delete_file(file_id)
            if not n:
                return jsonify({"ok": False, "error": "??????"}), 404
            return jsonify({"ok": True, "files": mysql_store.list_files()})
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL ????: {e}"}), 500
    _sign_ensure_session_inbox()
    sid = session["sign_inbox_sid"]
    rec = _sign_find_record(file_id)
    if not rec:
        return jsonify({"ok": False, "error": "??????"}), 404
    _sign_remove_disk_files_for_id(sid, file_id, rec.get("ext"))
    fid = str(file_id)
    session["sign_files"] = [
        r for r in (session.get("sign_files") or []) if str(r.get("id") or "") != fid
    ]
    session.modified = True
    pruned = _sign_prune_session_files_to_disk(sid)
    return jsonify({"ok": True, "files": pruned})


@app.route("/api/sign/detect", methods=["GET"])
def api_sign_detect():
    """自动识别文档中的签名位/角色（用于前端自动勾选）。"""
    file_id = (request.args.get("file_id") or "").strip()
    if not file_id or not _SIGN_FILE_ID_RE.match(file_id):
        return jsonify({"ok": False, "error": "缺少或无效的 file_id"}), 400

    ext = None
    source_path = None
    mysql_blob = None

    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            row = mysql_store.get_file_row(file_id)
            if not row:
                return jsonify({"ok": False, "error": "未找到该文件"}), 404
            if not row.get("file_data"):
                return jsonify(
                    {
                        "ok": False,
                        "error": "无法读取文件内容（可能 FTP 暂不可用或素材仅存在 FTP 且下载失败）。请稍后重试「重新识别」或检查 FTP 配置。",
                    }
                ), 502
            ext = (row.get("ext") or ".docx").lower()
            if ext not in SIGN_ALLOWED_EXT:
                return jsonify({"ok": False, "error": "不支持的文件类型"}), 400
            mysql_blob = row["file_data"]
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL 读取失败: {e}"}), 500
    else:
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        rec = _sign_find_record(file_id)
        if not rec:
            return jsonify({"ok": False, "error": "未找到该文件"}), 404
        ext = (rec.get("ext") or ".docx").lower()
        if ext not in SIGN_ALLOWED_EXT:
            return jsonify({"ok": False, "error": "不支持的文件类型"}), 400
        source_path = _sign_saved_disk_path(sid, file_id, ext)
        if not os.path.isfile(source_path):
            return jsonify({"ok": False, "error": "文件已不存在"}), 404

    tmp_dir = tempfile.mkdtemp(prefix="aiprintword_detect_")
    try:
        in_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex[:8]}_detect{ext}")
        if mysql_blob is not None:
            with open(in_path, "wb") as fp:
                fp.write(mysql_blob)
        else:
            shutil.copy2(source_path, in_path)

        from sign_handlers.detect_fields import detect_file

        result = detect_file(in_path)
        return jsonify(result)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


@app.route("/api/sign", methods=["POST"])
def api_sign():
    """
    ??????????PNG ??????PNG??    ???????    - file_id?????????????????
    - file??????????????????    ?? roles (JSON)??????sig_{id}?date_{id}??    """
    file_id = (request.form.get("file_id") or "").strip()
    upload = request.files.get("file")

    ext = None
    base_name = None
    source_path = None
    mysql_blob = None

    if file_id:
        if not _SIGN_FILE_ID_RE.match(file_id):
            return jsonify({"ok": False, "error": "??????id"}), 400
        if _sign_using_mysql():
            try:
                from sign_handlers import mysql_store

                mysql_store.ensure_sign_mysql()
                row = mysql_store.get_file_row(file_id)
                if not row or not row.get("file_data"):
                    return jsonify({"ok": False, "error": "??????????????"}), 404
                ext = (row.get("ext") or ".docx").lower()
                if ext not in SIGN_ALLOWED_EXT:
                    return jsonify({"ok": False, "error": "????????"}), 400
                base_name = _safe_display_filename_keep_unicode(row.get("name") or "document") or "document"
                if not base_name.lower().endswith(ext):
                    base_name = os.path.splitext(base_name)[0] + ext
                mysql_blob = row["file_data"]
            except Exception as e:
                return jsonify({"ok": False, "error": f"MySQL ????: {e}"}), 500
        else:
            _sign_ensure_session_inbox()
            sid = session["sign_inbox_sid"]
            rec = _sign_find_record(file_id)
            if not rec:
                return jsonify({"ok": False, "error": "??????????????"}), 404
            ext = (rec.get("ext") or ".docx").lower()
            if ext not in SIGN_ALLOWED_EXT:
                return jsonify({"ok": False, "error": "????????"}), 400
            source_path = _sign_saved_disk_path(sid, file_id, ext)
            if not os.path.isfile(source_path):
                return jsonify({"ok": False, "error": "????????????"}), 404
            base_name = _safe_display_filename_keep_unicode(rec.get("name") or "document") or "document"
            if not base_name.lower().endswith(ext):
                base_name = os.path.splitext(base_name)[0] + ext
            mysql_blob = None
    else:
        if not upload or not upload.filename:
            return jsonify({"ok": False, "error": "????????????????????"}), 400
        ext = os.path.splitext(upload.filename)[1].lower()
        if ext not in SIGN_ALLOWED_EXT:
            return jsonify({"ok": False, "error": "????.docx ??.xlsx"}), 400

    roles_raw = request.form.get("roles") or "[]"
    try:
        roles = json.loads(roles_raw)
    except Exception:
        return jsonify({"ok": False, "error": "?? roles ???? JSON"}), 400
    if not isinstance(roles, list) or not roles:
        return jsonify({"ok": False, "error": "???????????"}), 400

    from sign_handlers import ROLE_ID_TO_KEYWORD
    from sign_handlers.config import role_display_name

    sign_source = (request.form.get("sign_source") or "canvas").strip().lower()
    if sign_source not in ("canvas", "library"):
        sign_source = "canvas"

    sig_map = {}
    date_map = {}
    apply_report = {"applied": [], "skipped": []}
    if sign_source == "library":
        if not file_id:
            return jsonify({"ok": False, "error": "库映射模式仅支持对“已保存到列表”的文件生成"}), 400
        if not _sign_using_mysql():
            return jsonify({"ok": False, "error": "库映射模式需要启用 MySQL（MYSQL_HOST）"}), 400
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            mapping = mysql_store.get_file_role_signer_map(file_id) or {}
            if not isinstance(mapping, dict):
                mapping = {}
            for rid in roles:
                if rid not in ROLE_ID_TO_KEYWORD:
                    return jsonify({"ok": False, "error": f"无效角色 id: {rid}"}), 400
                pair = mapping.get(rid) or {}
                sig_id = (pair.get("sig") if isinstance(pair, dict) else None) or None
                date_id = (pair.get("date") if isinstance(pair, dict) else None) or None
                dm = (pair.get("date_mode") if isinstance(pair, dict) else None) or None
                diso = (pair.get("date_iso") if isinstance(pair, dict) else None) or None
                if mysql_store.is_composite_date_mode(dm):
                    # 拼接日期模式需要：签名素材 + 日历日期；若条件不齐，退回“能签就签”的普通素材模式
                    if not sig_id or not diso:
                        apply_report["skipped"].append(
                            {
                                "role_id": rid,
                                "role": role_display_name(rid),
                                "what": "composite_date",
                                "reason": "拼接日期条件不齐（需绑定签名并选择日历日期），已退回普通素材模式",
                            }
                        )
                        dm = None
                    if dm:
                        srow = mysql_store.get_stroke_item_row(sig_id)
                        if not srow or not srow.get("png"):
                            apply_report["skipped"].append(
                                {
                                    "role_id": rid,
                                    "role": role_display_name(rid),
                                    "what": "sig",
                                    "reason": "签名素材缺失",
                                }
                            )
                            dm = None
                    if dm:
                        sid0 = (srow.get("signer_id") or "").strip()
                        if not sid0:
                            apply_report["skipped"].append(
                                {
                                    "role_id": rid,
                                    "role": role_display_name(rid),
                                    "what": "composite_date",
                                    "reason": "无法解析签署人",
                                }
                            )
                            dm = None
                    if dm:
                        try:
                            lay = mysql_store.composite_mode_to_layout(dm)
                            dbytes, _lbl = mysql_store.compose_date_piece_png(
                                sid0, str(diso).strip(), lay
                            )
                        except Exception as e:
                            apply_report["skipped"].append(
                                {
                                    "role_id": rid,
                                    "role": role_display_name(rid),
                                    "what": "composite_date",
                                    "reason": f"日期拼接失败：{e}",
                                }
                            )
                            dm = None
                    if dm:
                        sig_map[rid] = srow["png"]
                        date_map[rid] = dbytes
                        apply_report["applied"].append(
                            {
                                "role_id": rid,
                                "role": role_display_name(rid),
                                "sig": True,
                                "date": True,
                                "date_mode": "composite",
                            }
                        )
                        continue
                applied_sig = False
                applied_date = False
                if sig_id:
                    srow = mysql_store.get_stroke_item_row(sig_id)
                    if srow and srow.get("png"):
                        sig_map[rid] = srow["png"]
                        applied_sig = True
                if date_id:
                    drow = mysql_store.get_stroke_item_row(date_id)
                    if drow and drow.get("png"):
                        date_map[rid] = drow["png"]
                        applied_date = True
                if applied_sig or applied_date:
                    apply_report["applied"].append(
                        {
                            "role_id": rid,
                            "role": role_display_name(rid),
                            "sig": applied_sig,
                            "date": applied_date,
                            "date_mode": "item",
                        }
                    )
                else:
                    apply_report["skipped"].append(
                        {
                            "role_id": rid,
                            "role": role_display_name(rid),
                            "what": "sig_date",
                            "reason": "未绑定或素材缺失（签名/日期均无可用数据）",
                        }
                    )
        except Exception as e:
            return jsonify({"ok": False, "error": f"库映射读取失败: {e}"}), 500
    else:
        for rid in roles:
            if rid not in ROLE_ID_TO_KEYWORD:
                return jsonify({"ok": False, "error": f"无效角色 id: {rid}"}), 400
            sig_raw = request.form.get(f"sig_{rid}") or ""
            date_raw = request.form.get(f"date_{rid}") or ""
            sig_bytes = _decode_png_data_url_or_b64(sig_raw)
            date_bytes = _decode_png_data_url_or_b64(date_raw)
            applied_sig = False
            applied_date = False
            if sig_bytes:
                sig_map[rid] = sig_bytes
                applied_sig = True
            if date_bytes:
                date_map[rid] = date_bytes
                applied_date = True
            if applied_sig or applied_date:
                apply_report["applied"].append(
                    {
                        "role_id": rid,
                        "role": role_display_name(rid),
                        "sig": applied_sig,
                        "date": applied_date,
                        "date_mode": "canvas",
                    }
                )
            else:
                apply_report["skipped"].append(
                    {
                        "role_id": rid,
                        "role": role_display_name(rid),
                        "what": "sig_date",
                        "reason": "该角色签名与日期画布均为空，已跳过",
                    }
                )

    # 至少要有一个角色提供签名或日期；否则不生成（避免输出与原文档一致却误以为“已签”）
    if (not sig_map) and (not date_map):
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "未选择任何可用素材：请至少为一个角色提供签名或日期（画布手写/库映射二选一）。",
                }
            ),
            400,
        )

    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            _max_s = _sign_mysql_max_signed()
            if mysql_store.count_signed_outputs() >= _max_s:
                return jsonify(
                    {
                        "ok": False,
                        "error": (
                            f"已签名记录已满（最多 {_max_s} 条），"
                            "请在本页「已签名文档」中删除旧记录后再试"
                        ),
                    }
                ), 400
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL 检查失败: {e}"}), 500

    if not base_name and upload and upload.filename:
        base_name = _safe_display_filename_keep_unicode(os.path.basename(upload.filename)) or "document"
        if not base_name.lower().endswith(ext):
            base_name = os.path.splitext(base_name)[0] + ext

    if file_id:
        if mysql_blob is not None:
            file_bytes = mysql_blob
        else:
            with open(source_path, "rb") as fp:
                file_bytes = fp.read()
    else:
        upload.stream.seek(0)
        file_bytes = upload.read()
        if not base_name:
            base_name = _safe_display_filename_keep_unicode(os.path.basename(upload.filename)) or "document"
            if not base_name.lower().endswith(ext):
                base_name = base_name + ext

    res = _sign_process_document_bytes(
        file_bytes,
        ext,
        base_name,
        roles,
        sig_map,
        date_map,
        file_id or None,
        batch_id=uuid.uuid4().hex,
    )
    if not res.get("ok"):
        return jsonify({"ok": False, "error": res.get("error", "失败")}), 500

    dl_name = res["dl_name"]
    out_bytes = res["out_bytes"]
    resp = send_file(
        io.BytesIO(out_bytes),
        as_attachment=True,
        download_name=dl_name,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if ext == ".docx"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
    sid_hdr = res.get("signed_id")
    if sid_hdr:
        resp.headers["X-Signed-Record-Id"] = sid_hdr
    try:
        # 更短更稳：只回传摘要（避免头部过大/截断导致前端解码失败）
        ap = apply_report.get("applied") or []
        sk = apply_report.get("skipped") or []
        ap_txt = "；".join(
            [
                f"{x.get('role') or x.get('role_id')}："
                + ("签名" if x.get("sig") else "")
                + ("+" if x.get("sig") and x.get("date") else "")
                + ("日期" if x.get("date") else "")
                + ("（拼接）" if x.get("date_mode") == "composite" else "")
                for x in ap
                if isinstance(x, dict)
            ]
        )
        sk_txt = "；".join(
            [
                f"{x.get('role') or x.get('role_id')}：{x.get('reason') or '跳过'}"
                for x in sk
                if isinstance(x, dict)
            ]
        )
        summary = {
            "applied_n": len(ap) if isinstance(ap, list) else 0,
            "skipped_n": len(sk) if isinstance(sk, list) else 0,
            "applied": ap_txt[:1200],
            "skipped": sk_txt[:1200],
        }
        rep_s = json.dumps(summary, ensure_ascii=False)
        rep_b64 = base64.b64encode(rep_s.encode("utf-8")).decode("ascii")
        resp.headers["X-Sign-Apply-Summary-B64"] = rep_b64
    except Exception:
        pass
    return resp


@app.route("/api/sign/signers", methods=["GET"])
def api_sign_signers_list():
    """签署人库列表（MySQL 多机共享；否则存于当前会话目录）。"""
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            return jsonify({"ok": True, "db_share": True, "signers": mysql_store.list_signers()})
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import list_signers as local_list_signers

        return jsonify(
            {"ok": True, "db_share": False, "signers": local_list_signers(SIGN_INBOX_ROOT, sid)}
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def _parse_signer_names_from_json(data: dict) -> list[str]:
    """支持 names 数组，或 name 中用中英文逗号/分号/换行分隔的多个姓名。"""
    import re as _re

    out: list[str] = []
    seen: set[str] = set()
    names_field = data.get("names")
    if isinstance(names_field, list):
        for x in names_field:
            n = str(x).strip()[:128]
            if not n:
                continue
            key = n.casefold()
            if key not in seen:
                seen.add(key)
                out.append(n)
    raw = (data.get("name") or "").strip()
    if raw:
        for part in _re.split(r"[,，;；\r\n]+", raw):
            n = part.strip()[:128]
            if not n:
                continue
            key = n.casefold()
            if key not in seen:
                seen.add(key)
                out.append(n)
    return out


@app.route("/api/sign/signers", methods=["POST"])
def api_sign_signers_create():
    data = request.get_json(silent=True) or {}
    names = _parse_signer_names_from_json(data)
    if not names:
        return jsonify({"ok": False, "error": "请填写至少一个签署人名称（可用逗号分隔多个）"}), 400
    if len(names) > 50:
        return jsonify({"ok": False, "error": "一次最多添加 50 个签署人"}), 400
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            created = []
            for name in names:
                nid = mysql_store.insert_signer(name)
                created.append({"id": nid, "name": name})
            return jsonify(
                {
                    "ok": True,
                    "added": len(created),
                    "created": created,
                    "signers": mysql_store.list_signers(),
                }
            )
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import insert_signer as local_insert_signer
        from sign_handlers.sign_library_local import list_signers as local_list_signers

        created = []
        for name in names:
            nid = local_insert_signer(SIGN_INBOX_ROOT, sid, name)
            created.append({"id": nid, "name": name})
        return jsonify(
            {
                "ok": True,
                "added": len(created),
                "created": created,
                "signers": local_list_signers(SIGN_INBOX_ROOT, sid),
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signers/<signer_id>", methods=["DELETE"])
def api_sign_signers_delete(signer_id):
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            n = mysql_store.delete_signer(signer_id)
            if not n:
                return jsonify({"ok": False, "error": "未找到该签署人"}), 404
            return jsonify({"ok": True, "signers": mysql_store.list_signers()})
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import delete_signer as local_delete_signer
        from sign_handlers.sign_library_local import list_signers as local_list_signers

        n = local_delete_signer(SIGN_INBOX_ROOT, sid, signer_id)
        if not n:
            return jsonify({"ok": False, "error": "未找到该签署人"}), 404
        return jsonify({"ok": True, "signers": local_list_signers(SIGN_INBOX_ROOT, sid)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signers/<signer_id>/strokes", methods=["PUT"])
def api_sign_signers_strokes_put(signer_id):
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    sig_raw = request.form.get("sig") or ""
    date_raw = request.form.get("date") or ""
    locale = (request.form.get("locale") or "zh").strip().lower()
    sig_b = _decode_png_data_url_or_b64(sig_raw) if str(sig_raw).strip() else None
    date_b = _decode_png_data_url_or_b64(date_raw) if str(date_raw).strip() else None
    if sig_b is None and date_b is None:
        return jsonify({"ok": False, "error": "请至少提交签名或日期笔迹之一"}), 400
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            sig_res = None
            date_res = None
            if sig_b is not None:
                sig_res = mysql_store.upsert_signer_stroke_item(signer_id, "sig", sig_b, locale=locale)
            if date_b is not None:
                date_res = mysql_store.upsert_signer_stroke_item(signer_id, "date", date_b, locale=locale)
            # 兼容旧返回字段：stroke_set_id 仍返回（若两者都提交则也写入旧 set）
            stroke_set_res = None
            if sig_b is not None and date_b is not None:
                try:
                    stroke_set_res = mysql_store.upsert_signer_strokes(signer_id, sig_b, date_b, locale=locale)
                except Exception:
                    stroke_set_res = None
            return jsonify(
                {
                    "ok": True,
                    "sig_item_id": (sig_res or {}).get("stroke_item_id"),
                    "date_item_id": (date_res or {}).get("stroke_item_id"),
                    "stroke_set_id": (stroke_set_res or {}).get("stroke_set_id"),
                    "overwritten": bool((sig_res or {}).get("overwritten") or (date_res or {}).get("overwritten") or (stroke_set_res or {}).get("overwritten")),
                }
            )
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import upsert_strokes as local_upsert

        # 会话模式：同样拆分保存（仍兼容 stroke_set）
        from sign_handlers.sign_library_local import upsert_stroke_item as local_upsert_item

        sig_res = None
        date_res = None
        if sig_b is not None:
            sig_res = local_upsert_item(SIGN_INBOX_ROOT, sid, signer_id, "sig", sig_b, locale=locale)
        if date_b is not None:
            date_res = local_upsert_item(SIGN_INBOX_ROOT, sid, signer_id, "date", date_b, locale=locale)
        stroke_set_res = None
        if sig_b is not None and date_b is not None:
            try:
                stroke_set_res = local_upsert(SIGN_INBOX_ROOT, sid, signer_id, sig_b, date_b, locale=locale)
            except Exception:
                stroke_set_res = None
        return jsonify(
            {
                "ok": True,
                "sig_item_id": (sig_res or {}).get("stroke_item_id"),
                "date_item_id": (date_res or {}).get("stroke_item_id"),
                "stroke_set_id": (stroke_set_res or {}).get("stroke_set_id"),
                "overwritten": bool((sig_res or {}).get("overwritten") or (date_res or {}).get("overwritten") or (stroke_set_res or {}).get("overwritten")),
            }
        )
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signers/<signer_id>/stroke-piece", methods=["PUT"])
def api_sign_signer_stroke_piece_put(signer_id):
    """录入英文点分日期笔迹元件：数字 0-9、月份 pm01..pm12、连接符 pdot（locale 固定 en）。"""
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "笔迹元件需启用 MySQL（MYSQL_HOST）"}), 400
    piece_kind = (request.form.get("piece_kind") or request.form.get("kind") or "").strip()
    png_raw = request.form.get("png") or ""
    png_b = _decode_png_data_url_or_b64(png_raw) if str(png_raw).strip() else None
    if not png_b:
        return jsonify({"ok": False, "error": "请提交 png 笔迹图"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        res = mysql_store.upsert_signer_stroke_piece(signer_id, piece_kind, png_b)
        return jsonify({"ok": True, **res})
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signers/<signer_id>/stroke-pieces", methods=["PUT"])
def api_sign_signer_stroke_pieces_batch_put(signer_id):
    """批量录入笔迹元件：JSON body { items: [ { piece_kind, png }, ... ] }。"""
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "笔迹元件需启用 MySQL（MYSQL_HOST）"}), 400
    data = request.get_json(silent=True) or {}
    items = data.get("items")
    if not isinstance(items, list) or not items:
        return jsonify({"ok": False, "error": "请求体需包含非空 items 数组"}), 400
    if len(items) > 80:
        return jsonify({"ok": False, "error": "单次最多提交 80 条元件"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        overwrite = True
        try:
            overwrite = bool(data.get("overwrite", True))
        except Exception:
            overwrite = True
        results = []
        for it in items:
            if not isinstance(it, dict):
                results.append({"piece_kind": "", "ok": False, "error": "条目须为 JSON 对象"})
                continue
            pk = (it.get("piece_kind") or it.get("kind") or "").strip()
            png_raw = it.get("png") or ""
            png_b = _decode_png_data_url_or_b64(png_raw) if str(png_raw).strip() else None
            if not png_b:
                results.append({"piece_kind": pk, "ok": False, "error": "缺少 png"})
                continue
            try:
                r = mysql_store.upsert_signer_stroke_piece(signer_id, pk, png_b)
                if (not overwrite) and r and r.get("overwritten"):
                    results.append(
                        {
                            "piece_kind": pk,
                            "ok": False,
                            "error_code": "exists",
                            "error": "该元件已存在，请确认是否覆盖",
                        }
                    )
                else:
                    results.append({"piece_kind": pk, "ok": True, **r})
            except ValueError as e:
                results.append({"piece_kind": pk, "ok": False, "error": str(e)})
            except Exception as e:
                results.append({"piece_kind": pk, "ok": False, "error": str(e) or type(e).__name__})
        return jsonify({"ok": True, "results": results})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signers/<signer_id>/composite-date-preview", methods=["GET"])
def api_sign_composite_date_preview(signer_id):
    """按签署人笔迹元件预览拼接 PNG。iso=YYYY-MM-DD；layout=zh_ymd|en_space|en_dot（默认 en_dot）。"""
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "需启用 MySQL"}), 400
    iso = (request.args.get("iso") or "").strip()
    layout = (request.args.get("layout") or "en_dot").strip().lower()
    if layout not in ("zh_ymd", "en_space", "en_dot"):
        return jsonify({"ok": False, "error": "layout 须为 zh_ymd / en_space / en_dot"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        png_b, label = mysql_store.compose_date_piece_png(signer_id, iso, layout)
        resp = send_file(io.BytesIO(png_b), mimetype="image/png")
        resp.headers["X-Composite-Date-Label"] = label
        return resp
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/sign/signers/<signer_id>/stroke/<kind>", methods=["GET"])
def api_sign_signer_stroke_get(signer_id, kind):
    if not _SIGN_FILE_ID_RE.match(signer_id or ""):
        return jsonify({"ok": False, "error": "无效的签署人 id"}), 400
    if kind not in ("sig", "date"):
        return jsonify({"ok": False, "error": "kind 须为 sig 或 date"}), 400
    try:
        blob = None
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            row = mysql_store.get_signer_strokes_row(signer_id)
            if row:
                blob = row.get("sig_png") if kind == "sig" else row.get("date_png")
        else:
            _sign_ensure_session_inbox()
            sid = session["sign_inbox_sid"]
            from sign_handlers.sign_library_local import get_strokes as local_get

            sig_b, date_b = local_get(SIGN_INBOX_ROOT, sid, signer_id)
            blob = sig_b if kind == "sig" else date_b
        if not blob:
            return Response(status=404)
        return Response(blob, mimetype="image/png")
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/stroke-sets/<set_id>/stroke/<kind>", methods=["GET"])
def api_sign_stroke_set_stroke_get(set_id, kind):
    if not _SIGN_FILE_ID_RE.match(set_id or ""):
        return jsonify({"ok": False, "error": "无效的笔迹套 id"}), 400
    if kind not in ("sig", "date"):
        return jsonify({"ok": False, "error": "kind 须为 sig 或 date"}), 400
    try:
        blob = None
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            row = mysql_store.get_stroke_set_row(set_id)
            if row:
                blob = row.get("sig_png") if kind == "sig" else row.get("date_png")
        else:
            _sign_ensure_session_inbox()
            sid = session["sign_inbox_sid"]
            from sign_handlers.sign_library_local import get_strokes_for_set as local_get_set

            sig_b, date_b = local_get_set(SIGN_INBOX_ROOT, sid, set_id)
            blob = sig_b if kind == "sig" else date_b
        if not blob:
            return Response(status=404)
        return Response(blob, mimetype="image/png")
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/stroke-items/<item_id>/png", methods=["GET"])
def api_sign_stroke_item_get(item_id):
    if not _SIGN_FILE_ID_RE.match(item_id or ""):
        return jsonify({"ok": False, "error": "无效的笔迹素材 id"}), 400
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            row = mysql_store.get_stroke_item_row(item_id)
            blob = (row or {}).get("png")
        else:
            _sign_ensure_session_inbox()
            sid = session["sign_inbox_sid"]
            from sign_handlers.sign_library_local import get_stroke_item_bytes as local_get_item

            blob = local_get_item(SIGN_INBOX_ROOT, sid, item_id)
        if not blob:
            return Response(status=404)
        return Response(blob, mimetype="image/png")
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/files/<file_id>/role-map", methods=["GET"])
def api_sign_file_role_map_get(file_id):
    if not _SIGN_FILE_ID_RE.match(file_id or ""):
        return jsonify({"ok": False, "error": "无效的文件 id"}), 400
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            m = mysql_store.get_file_role_signer_map(file_id)
            return jsonify({"ok": True, "map": m})
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import get_file_role_map as local_get_map

        return jsonify({"ok": True, "map": local_get_map(SIGN_INBOX_ROOT, sid, file_id)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/files/<file_id>/role-map", methods=["PUT"])
def api_sign_file_role_map_put(file_id):
    if not _SIGN_FILE_ID_RE.match(file_id or ""):
        return jsonify({"ok": False, "error": "无效的文件 id"}), 400
    data = request.get_json(silent=True) or {}
    m = data.get("map")
    if not isinstance(m, dict):
        return jsonify({"ok": False, "error": "请求体需包含 map 对象"}), 400
    clean = {}
    for k, v in m.items():
        if not k or not v:
            continue
        if isinstance(v, dict):
            clean[str(k)] = {
                "sig": v.get("sig"),
                "date": v.get("date"),
                "date_mode": v.get("date_mode"),
                "date_iso": v.get("date_iso"),
            }
        else:
            clean[str(k)] = str(v)
    try:
        if _sign_using_mysql():
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            mysql_store.set_file_role_signer_map(file_id, clean)
            return jsonify({"ok": True, "map": mysql_store.get_file_role_signer_map(file_id)})
        _sign_ensure_session_inbox()
        sid = session["sign_inbox_sid"]
        from sign_handlers.sign_library_local import set_file_role_map as local_set_map
        from sign_handlers.sign_library_local import get_file_role_map as local_get_map

        local_set_map(SIGN_INBOX_ROOT, sid, file_id, clean)
        return jsonify({"ok": True, "map": local_get_map(SIGN_INBOX_ROOT, sid, file_id)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/batch", methods=["POST"])
def api_sign_batch():
    """
    按每个文件已保存的 role-map，从签署人库取笔迹批量生成已签名文档。
    仅写入 sign_signed_output（MySQL）；不逐个返回文件流。
    """
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "批量签名需启用 MySQL（MYSQL_HOST）"}), 400
    data = request.get_json(silent=True) or {}
    source = (data.get("source") or "library").strip().lower()
    if source not in ("library", "canvas"):
        source = "library"
    file_ids = data.get("file_ids")
    if not isinstance(file_ids, list) or not file_ids:
        return jsonify({"ok": False, "error": "请提供 file_ids 数组"}), 400
    for fid in file_ids:
        if not _SIGN_FILE_ID_RE.match(str(fid or "")):
            return jsonify({"ok": False, "error": f"无效的文件 id: {fid}"}), 400

    from sign_handlers import ROLE_ID_TO_KEYWORD
    from sign_handlers import mysql_store

    mysql_store.ensure_sign_mysql()
    _max_s = _sign_mysql_max_signed()
    cur_n = mysql_store.count_signed_outputs()
    if cur_n + len(file_ids) > _max_s:
        return jsonify(
            {
                "ok": False,
                "error": (
                    f"已签名记录空间不足（当前 {cur_n}，本次 {len(file_ids)}，"
                    f"上限 {_max_s}），请先删除旧记录"
                ),
            }
        ), 400

    roles_req = data.get("roles")
    if roles_req is not None and (not isinstance(roles_req, list) or not roles_req):
        return jsonify({"ok": False, "error": "roles 必须为非空数组"}), 400
    if isinstance(roles_req, list):
        roles_req = [str(x) for x in roles_req if str(x or "").strip()]
        roles_req = list(dict.fromkeys(roles_req))
        for rr in roles_req:
            if rr not in ROLE_ID_TO_KEYWORD:
                return jsonify({"ok": False, "error": f"无效角色 id: {rr}"}), 400

    canvas_sig_map_raw = data.get("sig_map") if isinstance(data, dict) else None
    canvas_date_map_raw = data.get("date_map") if isinstance(data, dict) else None
    canvas_sig_map: dict = {}
    canvas_date_map: dict = {}
    if source == "canvas":
        if not isinstance(canvas_sig_map_raw, dict) or not isinstance(canvas_date_map_raw, dict):
            return jsonify({"ok": False, "error": "画布模式需提供 sig_map/date_map 对象"}), 400
        for rid, v in canvas_sig_map_raw.items():
            rid_s = str(rid)
            if rid_s in ROLE_ID_TO_KEYWORD:
                b = _decode_png_data_url_or_b64(v or "")
                if b:
                    canvas_sig_map[rid_s] = b
        for rid, v in canvas_date_map_raw.items():
            rid_s = str(rid)
            if rid_s in ROLE_ID_TO_KEYWORD:
                b = _decode_png_data_url_or_b64(v or "")
                if b:
                    canvas_date_map[rid_s] = b
        if roles_req:
            for rr in roles_req:
                if rr not in canvas_sig_map or rr not in canvas_date_map:
                    return jsonify({"ok": False, "error": f"角色 {rr} 缺少画布签名或日期"}), 400

    results = []
    # 允许前端分文件多次提交同一批（用于进度展示）；须为 32 位 hex，否则每次请求自动生成
    _bid = (data.get("batch_id") or "").strip().lower()
    batch_id = _bid if _bid and _SIGN_FILE_ID_RE.match(_bid) else uuid.uuid4().hex
    for fid in file_ids:
        try:
            row = mysql_store.get_file_row(fid)
            if not row or not row.get("file_data"):
                fe = (row.get("ftp_last_error") if isinstance(row, dict) else None) if row else None
                if fe:
                    results.append({"file_id": fid, "ok": False, "error": f"无法读取文件内容：{fe}"})
                else:
                    results.append({"file_id": fid, "ok": False, "error": "无法读取文件内容（可能 FTP 暂不可用）"})
                continue
            ext = (row.get("ext") or ".docx").lower()
            if ext not in SIGN_ALLOWED_EXT:
                results.append({"file_id": fid, "ok": False, "error": "不支持的扩展名"})
                continue
            base_name = _safe_display_filename_keep_unicode(row.get("name") or "document") or "document"
            if not base_name.lower().endswith(ext):
                base_name = os.path.splitext(base_name)[0] + ext
            mapping = mysql_store.get_file_role_signer_map(fid) or {}
            if not isinstance(mapping, dict):
                mapping = {}
            if roles_req:
                roles = [r for r in roles_req if r in ROLE_ID_TO_KEYWORD]
            else:
                roles = [r for r in mapping.keys() if r in ROLE_ID_TO_KEYWORD]
            if not roles:
                results.append({"file_id": fid, "ok": False, "error": "未配置角色-签署人映射"})
                continue
            sig_map = {}
            date_map = {}
            applied = []
            skipped = []
            if source == "canvas":
                for rid in roles:
                    sb = canvas_sig_map.get(rid)
                    db = canvas_date_map.get(rid)
                    if sb:
                        sig_map[rid] = sb
                    if db:
                        date_map[rid] = db
                    if sb or db:
                        applied.append(
                            {
                                "role_id": rid,
                                "sig": bool(sb),
                                "date": bool(db),
                                "date_mode": "canvas",
                            }
                        )
                    else:
                        skipped.append(
                            {
                                "role_id": rid,
                                "reason": "该角色签名与日期画布均为空，已跳过",
                            }
                        )
            else:
                for rid in roles:
                    pair = mapping.get(rid) or {}
                    sig_id = (pair.get("sig") if isinstance(pair, dict) else None) or None
                    date_id = (pair.get("date") if isinstance(pair, dict) else None) or None
                    dm = (pair.get("date_mode") if isinstance(pair, dict) else None) or None
                    diso = (pair.get("date_iso") if isinstance(pair, dict) else None) or None
                    if mysql_store.is_composite_date_mode(dm):
                        # 与单文件 POST /api/sign 一致：拼接失败不整角色放弃，退回「素材签名/整张日期图」能签则签
                        if not sig_id or not diso:
                            skipped.append(
                                {
                                    "role_id": rid,
                                    "reason": "拼接日期条件不齐（需签名素材+日历），已改试整张日期/仅签名",
                                }
                            )
                            dm = None
                        srow = None
                        if dm:
                            srow = mysql_store.get_stroke_item_row(sig_id)
                            if not srow or not srow.get("png"):
                                skipped.append(
                                    {
                                        "role_id": rid,
                                        "reason": "签名素材缺失，无法拼接；已改试整张日期/仅签名",
                                    }
                                )
                                dm = None
                        if dm:
                            sid0 = (srow.get("signer_id") or "").strip()
                            if not sid0:
                                skipped.append(
                                    {
                                        "role_id": rid,
                                        "reason": "无法解析签署人，跳过拼接；已改试整张日期/仅签名",
                                    }
                                )
                                dm = None
                        if dm:
                            try:
                                lay = mysql_store.composite_mode_to_layout(dm)
                                dbytes, _lbl = mysql_store.compose_date_piece_png(
                                    sid0, str(diso).strip(), lay
                                )
                            except Exception as e:
                                skipped.append(
                                    {
                                        "role_id": rid,
                                        "reason": f"日期拼接失败：{e}；已改试整张日期/仅签名",
                                    }
                                )
                                dm = None
                            else:
                                sig_map[rid] = srow["png"]
                                date_map[rid] = dbytes
                                applied.append(
                                    {
                                        "role_id": rid,
                                        "sig": True,
                                        "date": True,
                                        "date_mode": "composite",
                                    }
                                )
                                continue
                    sb = None
                    if sig_id:
                        srow = mysql_store.get_stroke_item_row(sig_id)
                        if srow and srow.get("png"):
                            sb = srow["png"]
                            sig_map[rid] = sb
                    db = None
                    if date_id:
                        drow = mysql_store.get_stroke_item_row(date_id)
                        if drow and drow.get("png"):
                            db = drow["png"]
                            date_map[rid] = db
                    if sb or db:
                        applied.append(
                            {
                                "role_id": rid,
                                "sig": bool(sb),
                                "date": bool(db),
                                "date_mode": "item",
                            }
                        )
                    else:
                        skipped.append(
                            {
                                "role_id": rid,
                                "reason": "未绑定或素材缺失（签名/日期均无可用数据）",
                            }
                        )
            # 至少要有一个角色提供签名或日期；否则该文件不生成
            if (not sig_map) and (not date_map):
                results.append(
                    {
                        "file_id": fid,
                        "ok": False,
                        "error": "未选择任何可用素材：请至少为一个角色提供签名或日期后再生成",
                        "applied_n": 0,
                        "skipped_n": len(skipped),
                        "skipped": "；".join(
                            [
                                f"{x.get('role_id')}：{x.get('reason') or '跳过'}"
                                for x in skipped
                                if isinstance(x, dict)
                            ]
                        )[:1200],
                    }
                )
                continue
            res = _sign_process_document_bytes(
                row["file_data"],
                ext,
                base_name,
                roles,
                sig_map,
                date_map,
                fid,
                batch_id=batch_id,
            )
            if not res.get("ok"):
                results.append(
                    {
                        "file_id": fid,
                        "ok": False,
                        "error": res.get("error", "失败"),
                        "applied_n": len(applied),
                        "skipped_n": len(skipped),
                    }
                )
            else:
                ap_txt = "；".join(
                    [
                        f"{x.get('role_id')}："
                        + ("签名" if x.get("sig") else "")
                        + ("+" if x.get("sig") and x.get("date") else "")
                        + ("日期" if x.get("date") else "")
                        + ("（拼接）" if x.get("date_mode") == "composite" else "")
                        for x in applied
                        if isinstance(x, dict)
                    ]
                )
                sk_txt = "；".join(
                    [
                        f"{x.get('role_id')}：{x.get('reason') or '跳过'}"
                        for x in skipped
                        if isinstance(x, dict)
                    ]
                )
                results.append(
                    {
                        "file_id": fid,
                        "ok": True,
                        "signed_id": res.get("signed_id"),
                        "name": res.get("dl_name"),
                        "applied_n": len(applied),
                        "skipped_n": len(skipped),
                        "applied": ap_txt[:1200],
                        "skipped": sk_txt[:1200],
                    }
                )
        except Exception as e:
            results.append({"file_id": fid, "ok": False, "error": str(e) or type(e).__name__})
    return jsonify({"ok": True, "results": results, "batch_id": batch_id})


@app.route("/api/sign/signed", methods=["GET"])
def api_sign_signed_list():
    """List signed outputs from MySQL shared storage（分页、按文件名搜索）。"""
    if not _sign_using_mysql():
        return jsonify(
            {
                "ok": True,
                "items": [],
                "total": 0,
                "page": 1,
                "page_size": 10,
                "db_share": False,
            }
        )
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        page = max(1, int(request.args.get("page") or 1))
        page_size = max(1, min(100, int(request.args.get("page_size") or 10)))
        items, total = mysql_store.list_signed_outputs_page(
            q=q, page=page, page_size=page_size
        )
        return jsonify(
            {
                "ok": True,
                "items": items,
                "total": total,
                "page": page,
                "page_size": page_size,
                "db_share": True,
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": f"MySQL ????: {e}"}), 500


@app.route("/api/sign/signed/<signed_id>", methods=["GET"])
def api_sign_signed_download(signed_id):
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "????MySQL ????"}), 404
    if not _SIGN_FILE_ID_RE.match(signed_id or ""):
        return jsonify({"ok": False, "error": "????id"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        row = mysql_store.get_signed_row(signed_id)
        if not row or not row.get("file_data"):
            return jsonify({"ok": False, "error": "??????"}), 404
        ext = (row.get("ext") or ".docx").lower()
        name = row.get("name") or ("signed" + ext)
        return send_file(
            io.BytesIO(row["file_data"]),
            as_attachment=True,
            download_name=name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if ext == ".docx"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed/<signed_id>", methods=["DELETE"])
def api_sign_signed_delete(signed_id):
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "????MySQL ????"}), 400
    if not _SIGN_FILE_ID_RE.match(signed_id or ""):
        return jsonify({"ok": False, "error": "????id"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        n = mysql_store.delete_signed_output(signed_id)
        if not n:
            return jsonify({"ok": False, "error": "??????"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed-batches", methods=["GET"])
def api_sign_signed_batches():
    """按批次列出已签名输出（分页、按批次/文件名搜索）。"""
    if not _sign_using_mysql():
        return jsonify(
            {
                "ok": True,
                "batches": [],
                "total": 0,
                "page": 1,
                "page_size": 10,
                "db_share": False,
            }
        )
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        page = max(1, int(request.args.get("page") or 1))
        page_size = max(1, min(100, int(request.args.get("page_size") or 10)))
        batches, total, legacy_total = mysql_store.list_signed_batches_page(
            q=q, page=page, page_size=page_size
        )
        return jsonify(
            {
                "ok": True,
                "batches": batches,
                "legacy_total": legacy_total,
                "total": total,
                "page": page,
                "page_size": page_size,
                "db_share": True,
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed-batch/<batch_id>", methods=["GET"])
def api_sign_signed_batch_items(batch_id):
    """列出某批次内的已签名文件（支持按文件名搜索）。"""
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "未启用 MySQL"}), 404
    if not _SIGN_FILE_ID_RE.match(batch_id or ""):
        return jsonify({"ok": False, "error": "无效批次 id"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        items = mysql_store.list_signed_outputs_by_batch(batch_id=batch_id, q=q)
        return jsonify({"ok": True, "items": items})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed-batch/<batch_id>/zip", methods=["GET"])
def api_sign_signed_batch_zip(batch_id):
    """下载某批次的 zip 包。"""
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "未启用 MySQL"}), 404
    if not _SIGN_FILE_ID_RE.match(batch_id or ""):
        return jsonify({"ok": False, "error": "无效批次 id"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        items = mysql_store.list_signed_outputs_by_batch(batch_id=batch_id, q="")
        if not items:
            return jsonify({"ok": False, "error": "该批次无文件"}), 404
        # 写入临时 zip，避免内存过大
        import tempfile

        fd, zp = tempfile.mkstemp(prefix=f"sign_batch_{batch_id}_", suffix=".zip")
        os.close(fd)
        try:
            with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
                seen = set()
                for it in items:
                    sid = it.get("id")
                    if not sid:
                        continue
                    row = mysql_store.get_signed_row(str(sid))
                    if not row or not row.get("file_data"):
                        continue
                    nm = (it.get("name") or row.get("name") or (sid + (it.get("ext") or ""))) or "signed"
                    nm = os.path.basename(nm)
                    base = nm
                    k = 2
                    while nm in seen:
                        root, ext = os.path.splitext(base)
                        nm = f"{root}({k}){ext}"
                        k += 1
                    seen.add(nm)
                    zf.writestr(nm, row["file_data"])
            dl = f"signed_batch_{batch_id[:8]}.zip"
            return send_file(zp, as_attachment=True, download_name=dl, mimetype="application/zip")
        finally:
            try:
                os.remove(zp)
            except Exception:
                pass
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed-legacy", methods=["GET"])
def api_sign_signed_legacy():
    """历史已签名记录（无 batch_id）：分页、搜索。"""
    if not _sign_using_mysql():
        return jsonify(
            {
                "ok": True,
                "items": [],
                "total": 0,
                "page": 1,
                "page_size": 50,
                "db_share": False,
            }
        )
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        page = max(1, int(request.args.get("page") or 1))
        page_size = max(1, min(200, int(request.args.get("page_size") or 50)))
        items, total = mysql_store.list_signed_legacy_page(q=q, page=page, page_size=page_size)
        return jsonify(
            {
                "ok": True,
                "items": items,
                "total": total,
                "page": page,
                "page_size": page_size,
                "db_share": True,
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed-legacy/zip", methods=["GET"])
def api_sign_signed_legacy_zip():
    """下载历史（无 batch_id）记录 zip 包（按搜索过滤）。"""
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "未启用 MySQL"}), 404
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        # 限制最多打包 300 个，避免 zip 过大
        items, total = mysql_store.list_signed_legacy_page(q=q, page=1, page_size=300)
        if not items:
            return jsonify({"ok": False, "error": "无匹配历史文件"}), 404
        import tempfile

        fd, zp = tempfile.mkstemp(prefix="sign_legacy_", suffix=".zip")
        os.close(fd)
        try:
            with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
                seen = set()
                for it in items:
                    sid = it.get("id")
                    if not sid:
                        continue
                    row = mysql_store.get_signed_row(str(sid))
                    if not row or not row.get("file_data"):
                        continue
                    nm = (it.get("name") or row.get("name") or (sid + (it.get("ext") or ""))) or "signed"
                    nm = os.path.basename(nm)
                    base = nm
                    k = 2
                    while nm in seen:
                        root, ext = os.path.splitext(base)
                        nm = f"{root}({k}){ext}"
                        k += 1
                    seen.add(nm)
                    zf.writestr(nm, row["file_data"])
            dl = "signed_legacy.zip"
            return send_file(zp, as_attachment=True, download_name=dl, mimetype="application/zip")
        finally:
            try:
                os.remove(zp)
            except Exception:
                pass
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/stroke-items", methods=["GET"])
def api_sign_stroke_items_list():
    """已入库签字 PNG 素材列表（分页、按签署人搜索）。"""
    if not _sign_using_mysql():
        return jsonify(
            {
                "ok": True,
                "items": [],
                "total": 0,
                "page": 1,
                "page_size": 10,
                "db_share": False,
            }
        )
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        q = (request.args.get("q") or "").strip()
        cat = (request.args.get("cat") or "").strip()
        page = max(1, int(request.args.get("page") or 1))
        page_size = max(1, min(100, int(request.args.get("page_size") or 10)))
        items, total = mysql_store.list_stroke_items_page(
            q=q, page=page, page_size=page_size, cat=cat
        )
        return jsonify(
            {
                "ok": True,
                "items": items,
                "total": total,
                "page": page,
                "page_size": page_size,
                "db_share": True,
            }
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/stroke-items/<item_id>", methods=["DELETE"])
def api_sign_stroke_item_delete(item_id):
    if not _sign_using_mysql():
        return jsonify({"ok": False, "error": "未启用 MySQL"}), 400
    if not _SIGN_FILE_ID_RE.match(item_id or ""):
        return jsonify({"ok": False, "error": "无效 id"}), 400
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        n = mysql_store.delete_stroke_item(item_id)
        if not n:
            return jsonify({"ok": False, "error": "记录不存在"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    _dp = _resolved_dotenv_path()
    print(
        "[aiprintword] build=%s app.py=%s"
        % (AIPRINTWORD_WEB_BUILD, os.path.abspath(__file__))
    )
    print(
        "[aiprintword] .env=%s exists=%s token_loaded=%s"
        % (
            _dp,
            os.path.isfile(_dp),
            bool((os.environ.get("AIPRINTWORD_ADMIN_TOKEN") or "").strip()),
        )
    )
    print(
        "[aiprintword] 若 token_loaded=False 或浏览器仍见极短 503：关闭其它占用 5050 的窗口后只保留本进程，并确认 .env 在同目录"
    )
    # 启动自检：端口被占用时，直接提示并退出，避免“看似启动成功但页面打不开/超时”
    _port = 5050
    try:
        _sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        _sock.bind(("127.0.0.1", _port))
        _sock.close()
    except OSError:
        print("")
        print("[aiprintword] 启动自检失败：端口 %s 已被占用。" % _port)
        print("[aiprintword] 已有其它 app.py 在跑，先关掉它，再重新启动本服务。")
        if os.name == "nt":
            print("[aiprintword] Windows 排查/关闭示例：")
            print("  netstat -ano | findstr \":%s\"" % _port)
            print("  tasklist /FI \"PID eq <PID>\" /FO LIST")
            print("  taskkill /PID <PID> /F")
        else:
            print("[aiprintword] Linux/macOS 排查/关闭示例：")
            print("  lsof -nP -iTCP:%s -sTCP:LISTEN" % _port)
            print("  kill -9 <PID>")
        raise SystemExit(2)
    app.run(host="0.0.0.0", port=5050, debug=False)
