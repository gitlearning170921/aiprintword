# -*- coding: utf-8 -*-
"""
批量打印 Web 服务：上传 Word/Excel/PDF，检查签字与修订后打印。
支持流式接口返回处理进度（当前任务项与进度）。
"""
import base64
import io
import json
import os
import shutil
import tempfile
import threading
import uuid
from queue import Queue

from flask import (
    Flask,
    request,
    jsonify,
    send_from_directory,
    send_file,
    Response,
)
from werkzeug.utils import secure_filename

# 允许的扩展名
ALLOWED_EXT = {
    ".doc", ".docx", ".docm",
    ".xls", ".xlsx", ".xlsm",
    ".pdf",
}

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024  # 256MB 单次上传总大小

# 项目根目录（用于提供 index）
ROOT = os.path.dirname(os.path.abspath(__file__))

# COM 常见“发生意外”类错误码，转为友好提示（含 Excel -2146827864）
_COM_ERROR_CODES = (-2147352567, -2147467259, -2146827864)


def _format_com_error(e):
    """若为 COM 异常则返回友好提示；已含【步骤名】的说明原样返回。"""
    try:
        s = str(e)
        if "【" in s and "】" in s:
            return s
    except Exception:
        pass
    try:
        if getattr(e, "args", None) and len(e.args) > 0 and e.args[0] in _COM_ERROR_CODES:
            return "文档处理时发生意外，请确认 WPS/Excel 已安装、文档未被占用，或稍后重试。"
    except Exception:
        pass
    return str(e)


def _allowed_file(filename):
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXT


def _get_printers():
    """获取系统打印机列表（仅 Windows）。"""
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
    """避免浏览器请求 favicon 时返回 404。"""
    return "", 204


@app.route("/api/printers")
def api_printers():
    """返回可用打印机列表。"""
    return jsonify({"printers": _get_printers()})


@app.route("/api/batch-print", methods=["POST"])
def api_batch_print():
    """
    接收上传的文件与选项，执行检查与打印，返回结果 JSON。
    表单字段：files (多文件), check_signature, accept_revisions, printer_name, copies, dry_run
    """
    files = request.files.getlist("files") or request.files.getlist("files[]")
    if not files or not any(f.filename for f in files):
        return jsonify({"ok": False, "error": "请至少选择一个文件"}), 400

    check_formal = request.form.get("check_formal", "true").lower() == "true"
    check_signature = request.form.get("check_signature", "true").lower() == "true"
    accept_revisions = request.form.get("accept_revisions", "true").lower() == "true"
    printer_name = request.form.get("printer_name") or None
    if printer_name == "":
        printer_name = None
    try:
        copies = int(request.form.get("copies", "1") or "1")
        copies = max(1, min(copies, 99))
    except ValueError:
        copies = 1
    dry_run = request.form.get("dry_run", "false").lower() == "true"

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        for f in files:
            if not f.filename or not _allowed_file(f.filename):
                continue
            base = os.path.basename(f.filename)
            safe_name = f"{uuid.uuid4().hex[:8]}_{base}"
            path = os.path.join(tmp_dir, safe_name)
            f.save(path)
            saved_paths.append(path)
            original_names.append(base)

        if not saved_paths:
            return jsonify({"ok": False, "error": "没有可处理的文件（仅支持 .doc/.docx/.xls/.xlsx/.pdf 等）"}), 400

        from batch_print import run_batch
        result = run_batch(
            saved_paths,
            recursive=False,
            check_formal=check_formal,
            check_signature=check_signature,
            accept_revisions=accept_revisions,
            printer_name=printer_name,
            copies=copies,
            dry_run=dry_run,
        )
        # 用原始文件名替换路径便于前端展示
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": _format_com_error(e)}), 500
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass


def _run_batch_with_progress(saved_paths, original_names, opts, queue):
    """在后台线程中执行 run_batch，通过 queue 推送进度与最终结果。"""
    # 工作线程必须初始化 COM，避免 RPC -2147023170（远程过程调用失败）
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass
    try:
        from batch_print import run_batch
        def progress_callback(step, file_index, file_total, file_name, percent=None):
            name = original_names[file_index - 1] if file_index <= len(original_names) else file_name
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
        result = run_batch(
            saved_paths,
            recursive=False,
            check_formal=opts["check_formal"],
            check_signature=opts["check_signature"],
            accept_revisions=opts["accept_revisions"],
            printer_name=opts["printer_name"],
            copies=opts["copies"],
            dry_run=opts["dry_run"],
            progress_callback=progress_callback,
        )
        for i, d in enumerate(result["details"]):
            d["filename"] = original_names[i] if i < len(original_names) else os.path.basename(d["path"])
        queue.put(("result", {"ok": True, "result": result}))
    except Exception as e:
        queue.put(("result", {"ok": False, "error": _format_com_error(e)}))
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass
        queue.put((None, None))  # 结束标记


def _sse_message(event, data):
    """生成一条 SSE 消息。"""
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


@app.route("/api/batch-print-stream", methods=["POST"])
def api_batch_print_stream():
    """
    流式批量打印：返回 text/event-stream，持续推送 progress 事件与最终 result 事件。
    progress: { step, fileIndex, fileTotal, fileName, percent }
    result: { ok, result? 或 error? }
    """
    files = request.files.getlist("files") or request.files.getlist("files[]")
    if not files or not any(f.filename for f in files):
        return jsonify({"ok": False, "error": "请至少选择一个文件"}), 400

    opts = {
        "check_formal": request.form.get("check_formal", "true").lower() == "true",
        "check_signature": request.form.get("check_signature", "true").lower() == "true",
        "accept_revisions": request.form.get("accept_revisions", "true").lower() == "true",
        "printer_name": request.form.get("printer_name") or None,
        "copies": int(request.form.get("copies", "1") or "1"),
        "dry_run": request.form.get("dry_run", "false").lower() == "true",
    }
    if opts["printer_name"] == "":
        opts["printer_name"] = None
    opts["copies"] = max(1, min(opts["copies"], 99))

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="aiprintword_")
        saved_paths = []
        original_names = []
        for f in files:
            if not f.filename or not _allowed_file(f.filename):
                continue
            base = os.path.basename(f.filename)
            safe_name = f"{uuid.uuid4().hex[:8]}_{base}"
            path = os.path.join(tmp_dir, safe_name)
            f.save(path)
            saved_paths.append(path)
            original_names.append(base)

        if not saved_paths:
            return jsonify({"ok": False, "error": "没有可处理的文件"}), 400

        queue = Queue()
        thread = threading.Thread(
            target=_run_batch_with_progress,
            args=(saved_paths, original_names, opts, queue),
            daemon=True,
        )
        thread.start()

        def generate():
            nonlocal tmp_dir
            try:
                while True:
                    event_type, payload = queue.get()
                    if event_type is None:
                        break
                    if event_type == "progress":
                        yield _sse_message("progress", payload)
                    elif event_type == "result":
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


# ---------- 在线签名（独立功能，不影响批量打印路由）----------
SIGN_ALLOWED_EXT = {".docx", ".xlsx"}


def _decode_png_data_url_or_b64(s):
    """解码前端 canvas.toDataURL('image/png') 或纯 base64。"""
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
    """在线签名页面（与首页批量打印分离）。"""
    return send_from_directory(os.path.join(ROOT, "static"), "sign.html")


@app.route("/api/sign", methods=["POST"])
def api_sign():
    """
    上传 .docx/.xlsx，按角色插入手写签名 PNG 与手写日期 PNG。
    表单：file, roles (JSON 数组，如 ["author","reviewer"]),
         每个角色 sig_{id}、date_{id} 为 PNG 的 dataURL 或 base64。
    """
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify({"ok": False, "error": "请选择文件"}), 400
    ext = os.path.splitext(upload.filename)[1].lower()
    if ext not in SIGN_ALLOWED_EXT:
        return jsonify({"ok": False, "error": "仅支持 .docx 或 .xlsx"}), 400

    roles_raw = request.form.get("roles") or "[]"
    try:
        roles = json.loads(roles_raw)
    except Exception:
        return jsonify({"ok": False, "error": "参数 roles 不是合法 JSON"}), 400
    if not isinstance(roles, list) or not roles:
        return jsonify({"ok": False, "error": "请至少勾选一个签字角色"}), 400

    from sign_handlers import sign_document, ROLE_ID_TO_KEYWORD

    sig_map = {}
    date_map = {}
    for rid in roles:
        if rid not in ROLE_ID_TO_KEYWORD:
            return jsonify({"ok": False, "error": f"未知角色 id: {rid}"}), 400
        sig_raw = request.form.get(f"sig_{rid}") or ""
        date_raw = request.form.get(f"date_{rid}") or ""
        sig_bytes = _decode_png_data_url_or_b64(sig_raw)
        date_bytes = _decode_png_data_url_or_b64(date_raw)
        if not sig_bytes or not date_bytes:
            return jsonify({"ok": False, "error": f"角色「{ROLE_ID_TO_KEYWORD[rid]}」需同时提供签名与日期手写图"}), 400
        sig_map[rid] = sig_bytes
        date_map[rid] = date_bytes

    tmp_dir = tempfile.mkdtemp(prefix="aiprintword_sign_")
    try:
        base_name = secure_filename(os.path.basename(upload.filename)) or "document"
        if not base_name.lower().endswith(ext):
            base_name = base_name + ext
        in_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex[:8]}_{base_name}")
        upload.save(in_path)

        out_path = sign_document(in_path, sig_map, date_map)
        dl_name = os.path.splitext(base_name)[0] + "_signed" + ext
        # 先读入内存再删临时目录，避免 send_file 流式读取时目录已被删
        with open(out_path, "rb") as fp:
            out_bytes = fp.read()
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        tmp_dir = None

        return send_file(
            io.BytesIO(out_bytes),
            as_attachment=True,
            download_name=dl_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if ext == ".docx"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as e:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)
