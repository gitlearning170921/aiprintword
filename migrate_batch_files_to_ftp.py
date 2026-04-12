#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
迁移批处理相关本地文件到 FTP（主动模式）。

覆盖范围：
1) data/batch_exports/<token>.zip -> FTP: batch/exports/<token>.zip
2) 将历史 ZIP 包内的“终版文件”逐条上传 -> FTP: batch/final/<hid>/<arcname>
   - arcname 与 ZIP 内路径一致（保留相对路径结构）
   - 会跳过 ZIP 根目录的「修改明细.txt」
3) （可选）data/batch_history/<hid>/stash/** -> FTP: batch/history_stash/<hid>/**

用法：
  python migrate_batch_files_to_ftp.py

可选环境变量：
  MIGRATE_BATCH_EXPORTS=1   # 默认 1
  MIGRATE_BATCH_FINAL_FROM_ZIP=1  # 默认 1
  MIGRATE_BATCH_STASH=0     # 默认 0（stash 属于临时/原始文件，一般不需要上 FTP）
  MIGRATE_LIMIT=50000       # 最多迁移文件数（防止一次跑太久）
"""

from __future__ import annotations

import json
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
os.chdir(ROOT)


def _load_env() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    p = (os.environ.get("AIPRINTWORD_DOTENV_PATH") or "").strip()
    if p and os.path.isfile(p):
        load_dotenv(p, override=True, encoding="utf-8-sig")
        return
    d = os.path.join(ROOT, ".env")
    if os.path.isfile(d):
        load_dotenv(d, override=True, encoding="utf-8-sig")


def main() -> int:
    _load_env()
    from ftp_store import upload_file, upload_bytes

    do_exports = str(os.environ.get("MIGRATE_BATCH_EXPORTS") or "1").strip() not in ("0", "false", "False")
    do_final = str(os.environ.get("MIGRATE_BATCH_FINAL_FROM_ZIP") or "1").strip() not in ("0", "false", "False")
    do_stash = str(os.environ.get("MIGRATE_BATCH_STASH") or "0").strip() not in ("0", "false", "False")
    lim_s = (os.environ.get("MIGRATE_LIMIT") or "50000").strip() or "50000"
    try:
        lim = int(lim_s)
    except Exception:
        lim = 50000
    lim = max(1, min(lim, 5000000))

    stats = {"files_seen": 0, "uploaded": 0, "failed": 0}

    if do_exports:
        exports_root = os.path.join(ROOT, "data", "batch_exports")
        if os.path.isdir(exports_root):
            for name in os.listdir(exports_root):
                if stats["files_seen"] >= lim:
                    break
                if not name.lower().endswith(".zip"):
                    continue
                token = os.path.splitext(name)[0]
                lp = os.path.join(exports_root, name)
                if not os.path.isfile(lp):
                    continue
                stats["files_seen"] += 1
                try:
                    upload_file(lp, f"batch/exports/{token}.zip")
                    stats["uploaded"] += 1
                except Exception:
                    stats["failed"] += 1

    if do_final:
        hist_root = os.path.join(ROOT, "data", "batch_history")
        exports_root = os.path.join(ROOT, "data", "batch_exports")
        if os.path.isdir(hist_root) and os.path.isdir(exports_root):
            for hid in os.listdir(hist_root):
                if stats["files_seen"] >= lim:
                    break
                recp = os.path.join(hist_root, hid, "record.json")
                if not os.path.isfile(recp):
                    continue
                try:
                    with open(recp, encoding="utf-8") as f:
                        rec = json.load(f)
                except Exception:
                    continue
                token = (rec.get("download_token") or "").strip()
                if not token:
                    continue
                zp = os.path.join(exports_root, token + ".zip")
                if not os.path.isfile(zp):
                    continue
                try:
                    with zipfile.ZipFile(zp, "r") as zf:
                        for zi in zf.infolist():
                            if stats["files_seen"] >= lim:
                                break
                            if zi.is_dir():
                                continue
                            arc = (zi.filename or "").replace("\\", "/").lstrip("/")
                            if not arc:
                                continue
                            if arc == "修改明细.txt":
                                continue
                            stats["files_seen"] += 1
                            try:
                                data = zf.read(zi)
                                upload_bytes(data, f"batch/final/{hid}/{arc}")
                                stats["uploaded"] += 1
                            except Exception:
                                stats["failed"] += 1
                except Exception:
                    stats["failed"] += 1

    if do_stash:
        hist_root = os.path.join(ROOT, "data", "batch_history")
        if os.path.isdir(hist_root):
            for hid in os.listdir(hist_root):
                if stats["files_seen"] >= lim:
                    break
                stash = os.path.join(hist_root, hid, "stash")
                if not os.path.isdir(stash):
                    continue
                for root, dirs, files in os.walk(stash):
                    if stats["files_seen"] >= lim:
                        break
                    rel_root = os.path.relpath(root, stash).replace("\\", "/")
                    if rel_root == ".":
                        rel_root = ""
                    for fn in files:
                        if stats["files_seen"] >= lim:
                            break
                        lp = os.path.join(root, fn)
                        if not os.path.isfile(lp):
                            continue
                        stats["files_seen"] += 1
                        try:
                            relp = (rel_root + "/" + fn).lstrip("/")
                            upload_file(lp, f"batch/history_stash/{hid}/{relp}")
                            stats["uploaded"] += 1
                        except Exception:
                            stats["failed"] += 1

    print("migrate_batch_files_to_ftp:", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

