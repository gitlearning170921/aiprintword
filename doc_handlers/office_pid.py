# -*- coding: utf-8 -*-
"""
解析 Word/Excel COM 对应的本机进程 PID，供单文件超时 / 取消 / 手动跳过时 taskkill。
WPS 等场景下 Application.Hwnd 常为 0，需用窗口链或 tasklist 前后差分兜底。
"""
import csv
import io
import logging
import subprocess

logger = logging.getLogger("aiprintword.office_pid")


def tasklist_pids_for_images(image_names):
    """当前机器上指定映像名的进程 PID 集合。"""
    pids = set()
    for imagename in image_names:
        try:
            r = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {imagename}", "/FO", "CSV", "/NH"],
                capture_output=True,
                timeout=45,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            raw = r.stdout or b""
            text = raw.decode("gbk", errors="replace") if isinstance(raw, bytes) else str(raw)
            for row in csv.reader(io.StringIO(text)):
                if len(row) >= 2:
                    try:
                        pids.add(int(row[1].strip('"')))
                    except ValueError:
                        pass
        except Exception as e:
            logger.debug("tasklist %s failed: %s", imagename, e)
    return pids


def hwnd_chain_to_pid(app, win32process):
    """依次尝试 Application / ActiveWindow / Windows(1..) 的 HWND 换 PID。"""
    if win32process is None or app is None:
        return None
    hwnds = []
    try:
        h = int(getattr(app, "Hwnd", 0) or 0)
        if h > 0:
            hwnds.append(h)
    except Exception:
        pass
    try:
        aw = getattr(app, "ActiveWindow", None)
        if aw is not None:
            h = int(getattr(aw, "Hwnd", 0) or 0)
            if h > 0:
                hwnds.append(h)
    except Exception:
        pass
    try:
        windows = getattr(app, "Windows", None)
        if windows is not None:
            n = int(getattr(windows, "Count", 0) or 0)
            for i in range(1, min(n, 40) + 1):
                try:
                    w = windows(i)
                    h = int(getattr(w, "Hwnd", 0) or 0)
                    if h > 0:
                        hwnds.append(h)
                except Exception:
                    pass
    except Exception:
        pass
    seen = set()
    for hwnd in hwnds:
        if hwnd in seen:
            continue
        seen.add(hwnd)
        try:
            _, pid = win32process.GetWindowThreadProcessId(int(hwnd))
            if pid:
                return int(pid)
        except Exception:
            pass
    return None


def resolve_office_app_pid(app, win32process, host_image_names, before_pids):
    """
    Dispatch 之后解析 PID：优先 HWND 链；失败则用 tasklist 与 Dispatch 前的 PID 集合做差。
    """
    pid = hwnd_chain_to_pid(app, win32process)
    if pid:
        return pid
    after = tasklist_pids_for_images(host_image_names)
    new = after - before_pids
    if not new:
        logger.warning(
            "无法解析 Office PID（无有效 HWND，且未观察到新进程 %s），超时/跳过可能无法结束进程",
            host_image_names,
        )
        return None
    if len(new) == 1:
        return next(iter(new))
    chosen = max(new)
    logger.warning("观察到多个新进程 PID %s，按启发式选用 %s", new, chosen)
    return chosen


def refresh_pid_after_doc_open(app, win32process):
    """打开文档后常出现窗口句柄，再尝试一次 HWND 链（不用 tasklist 差分）。"""
    return hwnd_chain_to_pid(app, win32process)
