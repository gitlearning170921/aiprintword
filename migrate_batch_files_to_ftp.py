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
from typing import Any, Dict, Optional, Tuple

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
    from ftp_store import try_upload_file, try_upload_bytes

    # 可选：若启用 MySQL，则同步回写库内 record_json，避免前端仍显示“未上传”
    try:
        import batch_history_mysql as bhm

        bhm_enabled = bool(bhm.enabled())
    except Exception:
        bhm = None
        bhm_enabled = False

    do_exports = str(os.environ.get("MIGRATE_BATCH_EXPORTS") or "1").strip() not in ("0", "false", "False")
    do_final = str(os.environ.get("MIGRATE_BATCH_FINAL_FROM_ZIP") or "1").strip() not in ("0", "false", "False")
    do_stash = str(os.environ.get("MIGRATE_BATCH_STASH") or "0").strip() not in ("0", "false", "False")
    backfill_final_paths = str(os.environ.get("MIGRATE_BACKFILL_FINAL_PATHS") or "1").strip() not in (
        "0",
        "false",
        "False",
    )
    lim_s = (os.environ.get("MIGRATE_LIMIT") or "50000").strip() or "50000"
    try:
        lim = int(lim_s)
    except Exception:
        lim = 50000
    lim = max(1, min(lim, 5000000))

    stats = {"files_seen": 0, "uploaded": 0, "failed": 0}

    exports_root = os.path.join(ROOT, "data", "batch_exports")
    hist_root = os.path.join(ROOT, "data", "batch_history")

    def _detail_is_timeout_skip(d: Any) -> bool:
        if not isinstance(d, dict):
            return False
        if d.get("timeout_skip"):
            return True
        return "【超时跳过】" in (d.get("message") or "")

    def _zip_arcname_timeout_original(original_rel: str) -> str:
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

    def _zip_arcname_with_processed_ext(original_arcname: str, processed_path: str) -> str:
        pe = os.path.splitext(processed_path or "")[1]
        if not pe:
            return (original_arcname or "").replace("\\", "/")
        norm = (original_arcname or "").replace("\\", "/")
        d, base = os.path.split(norm)
        stem = os.path.splitext(base)[0]
        new_base = stem + pe.lower()
        return f"{d}/{new_base}" if d else new_base

    # 优先以“记录”为中心迁移：这样可以在成功后回写 zip_ftp_uploaded/zip_ftp_error
    if (do_exports or do_final) and os.path.isdir(hist_root) and os.path.isdir(exports_root):
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
            if not isinstance(rec, dict):
                continue
            token = (rec.get("download_token") or "").strip()
            if not token:
                continue

            changed = False
            zp = os.path.join(exports_root, token + ".zip")
            if do_exports and os.path.isfile(zp):
                stats["files_seen"] += 1
                ftp_p, err = try_upload_file(zp, f"batch/exports/{token}.zip")
                if ftp_p:
                    stats["uploaded"] += 1
                    if rec.get("zip_ftp_uploaded") is not True:
                        rec["zip_ftp_uploaded"] = True
                        changed = True
                    # result 里也冗余一份，兼容前端/旧记录（必要时创建 result dict）
                    if not isinstance(rec.get("result"), dict):
                        rec["result"] = {}
                        changed = True
                    if rec["result"].get("zip_ftp_uploaded") is not True:
                        rec["result"]["zip_ftp_uploaded"] = True
                        changed = True
                    if rec["result"].get("zip_ftp_error"):
                        rec["result"]["zip_ftp_error"] = None
                        changed = True
                    if rec.get("zip_ftp_error"):
                        rec["zip_ftp_error"] = None
                        changed = True
                else:
                    stats["failed"] += 1
                    if err:
                        if not isinstance(rec.get("result"), dict):
                            rec["result"] = {}
                            changed = True
                        rec["result"]["zip_ftp_error"] = str(err)[:512]
                        rec["zip_ftp_error"] = str(err)[:512]
                        changed = True

            # 终版文件迁移：仅上传，不强行映射回 details（避免匹配错误）。
            # 可选：通过 details+original_names 重算 arcname，回写 final_ftp_path，让历史“单文件终版下载”可用。
            if do_final and os.path.isfile(zp):
                try:
                    with zipfile.ZipFile(zp, "r") as zf:
                        # 先全量上传（保持旧行为）
                        for zi in zf.infolist():
                            if stats["files_seen"] >= lim:
                                break
                            if zi.is_dir():
                                continue
                            arc = (zi.filename or "").replace("\\", "/").lstrip("/")
                            if not arc or arc == "修改明细.txt":
                                continue
                            stats["files_seen"] += 1
                            try:
                                data = zf.read(zi)
                                ftp_p2, _err2 = try_upload_bytes(data, f"batch/final/{hid}/{arc}")
                                if ftp_p2:
                                    stats["uploaded"] += 1
                                else:
                                    stats["failed"] += 1
                            except Exception:
                                stats["failed"] += 1

                        # 再回写 details.final_ftp_path（仅 success 且非 timeout_skip）
                        if backfill_final_paths:
                            res = rec.get("result")
                            if isinstance(res, dict):
                                details = res.get("details") or []
                                original_names = rec.get("original_names") or []
                                if isinstance(details, list) and isinstance(original_names, list):
                                    n = min(len(details), len(original_names))
                                    for i in range(n):
                                        d: Any = details[i]
                                        if not isinstance(d, dict):
                                            continue
                                        if not d.get("success"):
                                            continue
                                        if _detail_is_timeout_skip(d):
                                            continue
                                        # 兼容：即使本地 processed_path 不存在，也可用其扩展名还原 ZIP 内 arcname
                                        name = str(original_names[i] or "").replace("\\", "/")
                                        proc = str(d.get("processed_path") or "")
                                        arcname = _zip_arcname_with_processed_ext(name, proc) if proc else name
                                        arcname = (arcname or "").replace("\\", "/").lstrip("/")
                                        if not arcname:
                                            continue
                                        try:
                                            data = zf.read(arcname)
                                        except Exception:
                                            # 找不到就不回写，避免写错索引
                                            continue
                                        ftp_p3, err3 = try_upload_bytes(data, f"batch/final/{hid}/{arcname}")
                                        if ftp_p3:
                                            if d.get("final_ftp_path") != ftp_p3:
                                                d["final_ftp_path"] = ftp_p3
                                                changed = True
                                            if d.get("final_ftp_error"):
                                                d["final_ftp_error"] = None
                                                changed = True
                                        elif err3:
                                            d["final_ftp_error"] = str(err3)[:512]
                                            changed = True
                except Exception:
                    stats["failed"] += 1

            if changed:
                try:
                    with open(recp, "w", encoding="utf-8") as f:
                        json.dump(rec, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
                if bhm_enabled and bhm is not None:
                    try:
                        bhm.save_record(rec)
                    except Exception:
                        pass

    # 兼容：如果没有历史 record.json，但仍想“裸迁移”导出 ZIP（不回写记录）
    if do_exports and os.path.isdir(exports_root) and (not os.path.isdir(hist_root)):
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
            ftp_p, _err = try_upload_file(lp, f"batch/exports/{token}.zip")
            if ftp_p:
                stats["uploaded"] += 1
            else:
                stats["failed"] += 1

    if do_stash:
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
                            ftp_p, _err = try_upload_file(lp, f"batch/history_stash/{hid}/{relp}")
                            if ftp_p:
                                stats["uploaded"] += 1
                            else:
                                stats["failed"] += 1
                        except Exception:
                            stats["failed"] += 1

    # MySQL 里仅有 record_json、没有本地 data/batch_history/<id>/record.json 时，上面磁盘循环不会更新。
    # 这里按库逐条：本地有 ZIP 则上传；否则若 FTP 上已存在同名 ZIP，则只回写 zip_ftp_uploaded。
    do_mysql_bf = str(os.environ.get("MIGRATE_MYSQL_BACKFILL") or "1").strip() not in (
        "0",
        "false",
        "False",
    )
    if do_mysql_bf and bhm_enabled and bhm is not None and do_exports:
        from ftp_store import remote_file_exists
        from sign_handlers.mysql_store import _conn_commit, ensure_sign_mysql

        stats["mysql_backfill"] = {"rows": 0, "updated": 0, "skipped": 0}
        ensure_sign_mysql()
        try:
            with _conn_commit() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT id, record_json FROM app_batch_history")
                    mysql_rows = cur.fetchall() or []
        except Exception:
            mysql_rows = []
        for row in mysql_rows:
            stats["mysql_backfill"]["rows"] += 1
            try:
                rec = json.loads(row["record_json"])
            except Exception:
                stats["mysql_backfill"]["skipped"] += 1
                continue
            if not isinstance(rec, dict) or not (rec.get("id") or "").strip():
                stats["mysql_backfill"]["skipped"] += 1
                continue
            token = (rec.get("download_token") or "").strip()
            if not token:
                stats["mysql_backfill"]["skipped"] += 1
                continue
            res = rec.get("result")
            if not isinstance(res, dict):
                res = {}
            if rec.get("zip_ftp_uploaded") is True and res.get("zip_ftp_uploaded") is True:
                continue
            rel_zip = f"batch/exports/{token}.zip"
            zp = os.path.join(exports_root, token + ".zip")
            ftp_ok = False
            if os.path.isfile(zp):
                stats["files_seen"] += 1
                ftp_p, err = try_upload_file(zp, rel_zip)
                if ftp_p:
                    ftp_ok = True
                    stats["uploaded"] += 1
                elif err:
                    stats["failed"] += 1
            elif remote_file_exists(rel_zip):
                ftp_ok = True
            if not ftp_ok:
                continue
            rec["zip_ftp_uploaded"] = True
            if rec.get("zip_ftp_error"):
                rec["zip_ftp_error"] = None
            if not isinstance(rec.get("result"), dict):
                rec["result"] = {}
            rec["result"]["zip_ftp_uploaded"] = True
            if rec["result"].get("zip_ftp_error"):
                rec["result"]["zip_ftp_error"] = None
            try:
                bhm.save_record(rec)
                stats["mysql_backfill"]["updated"] += 1
            except Exception:
                pass

    print("migrate_batch_files_to_ftp:", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

