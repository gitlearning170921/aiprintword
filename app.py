# -*- coding: utf-8 -*-
"""
Batch print web service for Word/Excel/PDF.
"""
import base64
import io
import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from queue import Queue, Empty

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

# Allowed file extensions
ALLOWED_EXT = {
    ".doc", ".docx", ".docm",
    ".xls", ".xlsx", ".xlsm",
    ".pdf",
}

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024  # 256MB total upload limit
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or "aiprintword-dev-secret-change-with-env"

# Project root
ROOT = os.path.dirname(os.path.abspath(__file__))
BATCH_EXPORT_ROOT = os.path.join(ROOT, "data", "batch_exports")
_BATCH_EXPORT_TOKEN_RE = re.compile(r"^[0-9a-f]{32}$")

try:
    from dotenv import load_dotenv

    load_dotenv(os.path.join(ROOT, ".env"), override=False)
except ImportError:
    pass

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

    opts = _parse_batch_opts_from_form(request.form)

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        for f in files:
            rel = _upload_relpath(f.filename)
            if not rel:
                continue
            path = os.path.join(tmp_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            f.save(path)
            if _allowed_file(rel):
                saved_paths.append(path)
                original_names.append(rel)

        if not saved_paths:
            return jsonify({"ok": False, "error": "???????????? .doc/.docx/.xls/.xlsx/.pdf ??"}), 400

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
            word_font_profile=opts.get("doc_font_profile", "mixed"),
            printer_name=opts["printer_name"],
            copies=opts["copies"],
            dry_run=opts["dry_run"],
            skip_print=opts["skip_print"],
            raw_print=opts["raw_print"],
        )
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        if opts.get("skip_print") and result.get("total"):
            try:
                token = uuid.uuid4().hex
                mod_txt = build_batch_modification_zip_text(result["details"])
                _zip_batch_exports(
                    saved_paths, original_names, token, modification_report_text=mod_txt
                )
                result["download_token"] = token
                result["download_filename"] = "processed_documents.zip"
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


def _parse_batch_opts_from_form(form):
    """Parse run_mode: formal_save | standard_print | raw_print."""
    run_mode = (form.get("run_mode") or "standard_print").strip().lower()
    if run_mode not in ("formal_save", "standard_print", "raw_print"):
        run_mode = "standard_print"
    word_content_preserve = form.get("word_content_preserve", "true").lower() == "true"
    word_preserve_page_count = form.get("word_preserve_page_count", "true").lower() == "true"
    word_image_risk_guard = form.get("word_image_risk_guard", "false").lower() == "true"
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
            "doc_font_profile": doc_font_profile,
            "printer_name": printer_name,
            "copies": copies,
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
            "doc_font_profile": doc_font_profile,
            "printer_name": printer_name,
            "copies": copies,
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
        "doc_font_profile": doc_font_profile,
        "printer_name": printer_name,
        "copies": copies,
    }


def _zip_batch_exports(
    saved_paths, original_names, token, modification_report_text=None
):
    """?????????????? UTF-8 BOM ?????.txt?"""
    os.makedirs(BATCH_EXPORT_ROOT, exist_ok=True)
    zip_path = os.path.join(BATCH_EXPORT_ROOT, f"{token}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        detail_text = (
            str(modification_report_text).strip()
            if modification_report_text is not None
            else ""
        )
        if not detail_text:
            detail_text = "?????????????????"
        zf.writestr(
            "\u4fee\u6539\u660e\u7ec6.txt",
            "\ufeff" + detail_text + "\n",
            compress_type=zipfile.ZIP_DEFLATED,
        )
        for p, name in zip(saved_paths, original_names):
            if os.path.isfile(p):
                zf.write(p, arcname=name)
    return zip_path


@app.route("/api/batch-export/<token>")
def api_batch_export(token):
    """?????????? ZIP?token ???????????"""
    if not _BATCH_EXPORT_TOKEN_RE.match(token or ""):
        return jsonify({"ok": False, "error": "invalid token"}), 404
    path = os.path.join(BATCH_EXPORT_ROOT, token + ".zip")
    if not os.path.isfile(path):
        return jsonify({"ok": False, "error": "not found or expired"}), 404
    return send_file(
        path,
        as_attachment=True,
        download_name="processed_documents.zip",
        max_age=0,
    )


def _run_batch_with_progress(saved_paths, original_names, opts, queue):
    """Run batch in background and push progress to queue."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        from batch_print import run_batch, build_batch_modification_zip_text

        def progress_callback(step, file_index, file_total, file_name, percent=None):
            name = original_names[file_index - 1] if file_index <= len(original_names) else file_name
            logger.info(
                "progress idx=%s/%s pct=%s step=%s file=%s",
                file_index,
                file_total,
                percent,
                step,
                name,
            )
            queue.put(
                (
                    "progress",
                    {
                        "step": step,
                        "fileIndex": file_index,
                        "fileTotal": file_total,
                        "fileName": name,
                        "percent": percent,
                    },
                )
            )

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
            word_font_profile=opts.get("doc_font_profile", "mixed"),
            printer_name=opts["printer_name"],
            copies=opts["copies"],
            dry_run=opts["dry_run"],
            skip_print=opts["skip_print"],
            raw_print=opts["raw_print"],
            progress_callback=progress_callback,
        )
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        logger.info(
            "batch worker done total=%s ok=%s failed=%s",
            result.get("total"),
            result.get("ok"),
            result.get("failed"),
        )
        if opts.get("skip_print") and result.get("total"):
            try:
                token = uuid.uuid4().hex
                mod_txt = build_batch_modification_zip_text(result["details"])
                _zip_batch_exports(
                    saved_paths, original_names, token, modification_report_text=mod_txt
                )
                result["download_token"] = token
                result["download_filename"] = "processed_documents.zip"
            except Exception as e:
                logger.exception("batch zip failed: %s", e)
                queue.put(("result", {"ok": False, "error": _format_com_error(e)}))
                return
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


def _sse_message(event, data):
    """Build one SSE message."""
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


@app.route("/api/batch-print-stream", methods=["POST"])
def api_batch_print_stream():
    """
    Streaming batch print via text/event-stream.
    progress: { step, fileIndex, fileTotal, fileName, percent }
    result: { ok, result? or error? }
    """
    files = request.files.getlist("files") or request.files.getlist("files[]")
    if not files or not any(f.filename for f in files):
        return jsonify({"ok": False, "error": "?????????"}), 400

    opts = _parse_batch_opts_from_form(request.form)

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        for f in files:
            rel = _upload_relpath(f.filename)
            if not rel:
                continue
            path = os.path.join(tmp_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            f.save(path)
            if _allowed_file(rel):
                saved_paths.append(path)
                original_names.append(rel)

        if not saved_paths:
            return jsonify({"ok": False, "error": "????????"}), 400
        logger.info("stream request accepted files=%s", len(saved_paths))

        queue = Queue()
        thread = threading.Thread(
            target=_run_batch_with_progress,
            args=(saved_paths, original_names, opts, queue),
            daemon=True,
        )
        thread.start()

        def generate():
            nonlocal tmp_dir
            started_at = time.time()
            last_emit = time.time()
            try:
                while True:
                    try:
                        event_type, payload = queue.get(timeout=5)
                    except Empty:
                        waited = int(time.time() - last_emit)
                        yield _sse_message(
                            "heartbeat",
                            {
                                "ts": int(time.time()),
                                "message": f"???????????????? {waited} ?",
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
                        yield _sse_message("result", payload)
                        break
            finally:
                if tmp_dir and os.path.isdir(tmp_dir):
                    try:
                        shutil.rmtree(tmp_dir, ignore_errors=True)
                    except Exception:
                        pass
                    tmp_dir = None

        return Response(
            generate(),
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


# ---------- ????????????????????----------
SIGN_ALLOWED_EXT = {".docx", ".xlsx"}
SIGN_INBOX_ROOT = os.path.join(ROOT, "data", "sign_inbox")
SIGN_MAX_SAVED_FILES = 50
SIGN_MYSQL_MAX_FILES = int(os.environ.get("SIGN_MYSQL_MAX_FILES", "500") or "500")
SIGN_MYSQL_MAX_FILES = max(1, min(SIGN_MYSQL_MAX_FILES, 10000))
SIGN_MYSQL_MAX_SIGNED = int(os.environ.get("SIGN_MYSQL_MAX_SIGNED", "2000") or "2000")
SIGN_MYSQL_MAX_SIGNED = max(1, min(SIGN_MYSQL_MAX_SIGNED, 50000))
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


def _sign_find_record(file_id: str):
    for rec in session.get("sign_files") or []:
        if rec.get("id") == file_id:
            return rec
    return None


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


@app.route("/sign")
def sign_page():
    """Serve online signature page."""
    return send_from_directory(os.path.join(ROOT, "static"), "sign.html")


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
    files = []
    for rec in session.get("sign_files") or []:
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
            if cur_n + len(parsed) > SIGN_MYSQL_MAX_FILES:
                return jsonify(
                    {
                        "ok": False,
                        "error": (
                            f"???? {cur_n} ?????? {len(parsed)} ??"
                            f"????????{SIGN_MYSQL_MAX_FILES} ????????????"
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
    ext = rec.get("ext") or ".docx"
    path = _sign_saved_disk_path(sid, file_id, ext)
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass
    session["sign_files"] = [r for r in (session.get("sign_files") or []) if r.get("id") != file_id]
    session.modified = True
    return jsonify({"ok": True, "files": session["sign_files"]})


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
                base_name = secure_filename(row.get("name") or "document") or "document"
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
            base_name = secure_filename(rec.get("name") or "document") or "document"
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

    from sign_handlers import sign_document, ROLE_ID_TO_KEYWORD
    from sign_handlers.config import role_display_name

    sig_map = {}
    date_map = {}
    for rid in roles:
        if rid not in ROLE_ID_TO_KEYWORD:
            return jsonify({"ok": False, "error": f"???? id: {rid}"}), 400
        sig_raw = request.form.get(f"sig_{rid}") or ""
        date_raw = request.form.get(f"date_{rid}") or ""
        sig_bytes = _decode_png_data_url_or_b64(sig_raw)
        date_bytes = _decode_png_data_url_or_b64(date_raw)
        if not sig_bytes or not date_bytes:
            return jsonify(
                {"ok": False, "error": f"???{role_display_name(rid)}??????????????"}
            ), 400
        sig_map[rid] = sig_bytes
        date_map[rid] = date_bytes

    if _sign_using_mysql():
        try:
            from sign_handlers import mysql_store

            mysql_store.ensure_sign_mysql()
            if mysql_store.count_signed_outputs() >= SIGN_MYSQL_MAX_SIGNED:
                return jsonify(
                    {
                        "ok": False,
                        "error": (
                            f"?????????? {SIGN_MYSQL_MAX_SIGNED} ???"
                            "???????????????"
                        ),
                    }
                ), 400
        except Exception as e:
            return jsonify({"ok": False, "error": f"MySQL ????? {e}"}), 500

    tmp_dir = tempfile.mkdtemp(prefix="aiprintword_sign_")
    try:
        if not base_name:
            base_name = secure_filename(os.path.basename(upload.filename)) or "document"
            if not base_name.lower().endswith(ext):
                base_name = base_name + ext
        in_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex[:8]}_{base_name}")
        if file_id and _sign_using_mysql():
            with open(in_path, "wb") as fp:
                fp.write(mysql_blob)
        elif file_id and source_path:
            shutil.copy2(source_path, in_path)
        else:
            upload.save(in_path)

        out_path = sign_document(in_path, sig_map, date_map)
        dl_name = os.path.splitext(base_name)[0] + "_signed" + ext
        # ?????????????? send_file ??????????
        with open(out_path, "rb") as fp:
            out_bytes = fp.read()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        tmp_dir = None

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
        if _sign_using_mysql():
            try:
                from sign_handlers import mysql_store

                mysql_store.ensure_sign_mysql()
                signed_row_id = uuid.uuid4().hex
                src_id = file_id if file_id and _SIGN_FILE_ID_RE.match(file_id) else None
                src_nm = base_name or "document"
                mysql_store.insert_signed_output(
                    signed_row_id,
                    src_id,
                    src_nm,
                    dl_name,
                    ext,
                    json.dumps(roles, ensure_ascii=False),
                    out_bytes,
                )
                resp.headers["X-Signed-Record-Id"] = signed_row_id
            except Exception as e:
                app.logger.exception("???????? MySQL ??: %s", e)
        return resp
    except Exception as e:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sign/signed", methods=["GET"])
def api_sign_signed_list():
    """List signed outputs from MySQL shared storage."""
    if not _sign_using_mysql():
        return jsonify({"ok": True, "items": [], "db_share": False})
    try:
        from sign_handlers import mysql_store

        mysql_store.ensure_sign_mysql()
        items = mysql_store.list_signed_outputs()
        return jsonify({"ok": True, "items": items, "db_share": True})
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
        return jsonify({"ok": True, "items": mysql_store.list_signed_outputs()})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
