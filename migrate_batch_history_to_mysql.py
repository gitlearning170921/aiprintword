#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将 data/batch_history/<id>/record.json 导入 MySQL（需配置 MYSQL_HOST）。

用法：在项目根目录执行
  python migrate_batch_history_to_mysql.py

会跳过库中已存在的 id，导入结束后按 AIPRINTWORD_HISTORY_MAX 裁剪条数，
并对缺少 display_title 的库内记录做一次补全。
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
    import batch_history_mysql as bhm

    if not bhm.enabled():
        print("未启用 MySQL（MYSQL_HOST 为空），退出。")
        return 1
    try:
        from runtime_settings.resolve import get_setting

        max_keep = max(5, min(int(get_setting("AIPRINTWORD_HISTORY_MAX")), 500))
    except Exception:
        max_keep = 50
    hist_root = os.path.join(ROOT, "data", "batch_history")
    st = bhm.migrate_from_disk(hist_root, max_keep)
    bf = bhm.backfill_display_titles()
    print("migrate_from_disk:", st)
    print("backfill_display_titles:", bf)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
