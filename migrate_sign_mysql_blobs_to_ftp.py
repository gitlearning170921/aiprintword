#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将在线签名模块历史遗留的 MySQL BLOB 文件迁移到 FTP（主动模式）。

迁移范围：
- sign_uploaded_file.file_data  -> FTP，回写 ftp_path/file_size/sha256（可选清空 BLOB）
- sign_signed_output.file_data  -> FTP，回写 ftp_path/file_size/sha256（可选清空 BLOB）

用法（项目根目录）：
  python migrate_sign_mysql_blobs_to_ftp.py

可选参数：
  --batch-size 2000
  --max-total 200000
  --clear-blob 1   # 1/0，默认 1
  --abort-after-failures 3   # 累计失败达到此次数即中止并打印错误样例；0=不限制

若终端长时间无输出，可改用无缓冲运行：python -u migrate_sign_mysql_blobs_to_ftp.py ...
"""

from __future__ import annotations

import argparse
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


def _stdout_line_buffering() -> None:
    try:
        sys.stdout.reconfigure(line_buffering=True)  # type: ignore[attr-defined]
    except Exception:
        pass


def main() -> int:
    _load_env()
    _stdout_line_buffering()

    ap = argparse.ArgumentParser(description="将签名相关 MySQL BLOB 迁移到 FTP")
    ap.add_argument("--batch-size", type=int, default=2000)
    ap.add_argument("--max-total", type=int, default=200000)
    ap.add_argument("--clear-blob", type=int, default=1, help="1=迁移成功后清空 MySQL BLOB，0=保留")
    ap.add_argument(
        "--abort-after-failures",
        type=int,
        default=3,
        help="FTP 等失败累计达到此次数则立即中止并打印错误样例；0 表示不限制",
    )
    args = ap.parse_args()
    bs = max(1, int(args.batch_size))
    mt = max(1, int(args.max_total))
    clear_blob = int(args.clear_blob) != 0
    aa = int(args.abort_after_failures)
    abort_after = None if aa <= 0 else aa

    print(
        f"开始迁移：batch_size={bs} max_total={mt} clear_blob={clear_blob} "
        f"abort_after_failures={abort_after if abort_after is not None else '无限制'}（进度会按批输出）",
        flush=True,
    )

    from sign_handlers import mysql_store

    if not mysql_store.mysql_sign_enabled():
        print("未启用 MySQL（MYSQL_HOST 为空），退出。", flush=True)
        return 1

    print("MySQL 已配置，正在执行迁移（若数据量大或 FTP 较慢，请耐心等待）…", flush=True)

    pl = True
    print("[1/4] sign_uploaded_file / sign_signed_output …", flush=True)
    st = mysql_store.migrate_mysql_blobs_to_ftp(
        batch_size=bs,
        max_total=mt,
        clear_blob=clear_blob,
        progress_log=pl,
        abort_after_failures=abort_after,
    )
    if st.get("aborted"):
        print("migrate_mysql_blobs_to_ftp:", st, flush=True)
        print("因失败次数达到上限已中止，请修正 FTP 或网络后重试。", flush=True)
        return 1
    print("[2/4] sign_signer_stroke …", flush=True)
    st2 = mysql_store.migrate_signer_strokes_blobs_to_ftp(
        batch_size=bs,
        max_total=mt,
        clear_blob=clear_blob,
        progress_log=pl,
        abort_after_failures=abort_after,
    )
    if st2.get("aborted"):
        print("migrate_mysql_blobs_to_ftp:", st, flush=True)
        print("migrate_signer_strokes_blobs_to_ftp:", st2, flush=True)
        print("因失败次数达到上限已中止，请修正 FTP 或网络后重试。", flush=True)
        return 1
    print("[3/4] sign_stroke_item …", flush=True)
    st3 = mysql_store.migrate_stroke_items_blobs_to_ftp(
        batch_size=bs,
        max_total=mt,
        clear_blob=clear_blob,
        progress_log=pl,
        abort_after_failures=abort_after,
    )
    if st3.get("aborted"):
        print("migrate_mysql_blobs_to_ftp:", st, flush=True)
        print("migrate_signer_strokes_blobs_to_ftp:", st2, flush=True)
        print("migrate_stroke_items_blobs_to_ftp:", st3, flush=True)
        print("因失败次数达到上限已中止，请修正 FTP 或网络后重试。", flush=True)
        return 1
    print("[4/4] 校验并补传 FTP …", flush=True)
    st4 = mysql_store.verify_and_backfill_ftp_files(
        limit=bs, error_samples=20, progress_log=pl, abort_after_failures=abort_after
    )
    if st4.get("aborted"):
        print("migrate_mysql_blobs_to_ftp:", st, flush=True)
        print("migrate_signer_strokes_blobs_to_ftp:", st2, flush=True)
        print("migrate_stroke_items_blobs_to_ftp:", st3, flush=True)
        print("verify_and_backfill_ftp_files:", st4, flush=True)
        print("因失败次数达到上限已中止，请修正 FTP 或网络后重试。", flush=True)
        return 1
    print("migrate_mysql_blobs_to_ftp:", st, flush=True)
    print("migrate_signer_strokes_blobs_to_ftp:", st2, flush=True)
    print("migrate_stroke_items_blobs_to_ftp:", st3, flush=True)
    print("verify_and_backfill_ftp_files:", st4, flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
