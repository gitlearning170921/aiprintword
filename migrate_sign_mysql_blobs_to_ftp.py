#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将在线签名模块历史遗留的 MySQL BLOB 文件迁移到 FTP（主动模式）。

迁移范围：
- sign_uploaded_file.file_data  -> FTP，回写 ftp_path/file_size/sha256（可选清空 BLOB）
- sign_signed_output.file_data  -> FTP，回写 ftp_path/file_size/sha256（可选清空 BLOB）

用法（项目根目录）：
  python migrate_sign_mysql_blobs_to_ftp.py
可选环境变量：
  MIGRATE_BATCH_SIZE=2000
  MIGRATE_MAX_TOTAL=200000
  MIGRATE_CLEAR_BLOB=1   # 1/0，默认 1
"""

from __future__ import annotations

import os
import sys

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
    from sign_handlers import mysql_store

    if not mysql_store.mysql_sign_enabled():
        print("未启用 MySQL（MYSQL_HOST 为空），退出。")
        return 1
    bs_s = (os.environ.get("MIGRATE_BATCH_SIZE") or "2000").strip() or "2000"
    try:
        bs = int(bs_s)
    except Exception:
        bs = 2000
    mt_s = (os.environ.get("MIGRATE_MAX_TOTAL") or "200000").strip() or "200000"
    try:
        mt = int(mt_s)
    except Exception:
        mt = 200000
    clear_blob = str(os.environ.get("MIGRATE_CLEAR_BLOB") or "1").strip() not in ("0", "false", "False")
    st = mysql_store.migrate_mysql_blobs_to_ftp(batch_size=bs, max_total=mt, clear_blob=clear_blob)
    st2 = mysql_store.migrate_signer_strokes_blobs_to_ftp(batch_size=bs, max_total=mt, clear_blob=clear_blob)
    st3 = mysql_store.verify_and_backfill_ftp_files(limit=bs, error_samples=20)
    print("migrate_mysql_blobs_to_ftp:", st)
    print("migrate_signer_strokes_blobs_to_ftp:", st2)
    print("verify_and_backfill_ftp_files:", st3)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

