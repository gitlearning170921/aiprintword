#!/usr/bin/env python3
"""将指定签署人「当天」录入的笔迹迁移到另一签署人（修复误选签署人）。"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

from sign_handlers import mysql_store  # noqa: E402
from sign_handlers.mysql_store import _conn_commit, _ensure_signer_tables, ensure_sign_mysql  # noqa: E402


def _resolve_signer_id(cur, name: str) -> str:
    cur.execute(
        "SELECT id, display_name FROM sign_signer WHERE display_name=%s",
        (name.strip(),),
    )
    rows = cur.fetchall() or []
    if len(rows) != 1:
        names = [r.get("display_name") for r in rows]
        raise SystemExit(f"签署人「{name}」匹配到 {len(rows)} 条：{names}")
    return str(rows[0]["id"])


def migrate(
    from_name: str,
    to_name: str,
    on_date: date,
    *,
    apply: bool,
) -> dict:
    ensure_sign_mysql()
    stats = {
        "items_moved": 0,
        "items_replaced_target": 0,
        "items_skipped": 0,
        "sets_moved": 0,
        "sets_replaced_target": 0,
        "role_rows_updated": 0,
        "legacy_merged": False,
    }
    day_str = on_date.isoformat()

    with _conn_commit() as conn:
        _ensure_signer_tables(conn)
        with conn.cursor() as cur:
            from_id = _resolve_signer_id(cur, from_name)
            to_id = _resolve_signer_id(cur, to_name)
            if from_id == to_id:
                raise SystemExit("源与目标签署人相同，无需迁移")

            cur.execute(
                """
                SELECT id, locale, kind, sha256, updated_at
                FROM sign_stroke_item
                WHERE signer_id=%s AND DATE(updated_at)=%s
                ORDER BY updated_at ASC
                """,
                (from_id, day_str),
            )
            items = list(cur.fetchall() or [])
            print(f"待迁移 stroke_item：{len(items)} 条（{from_name} -> {to_name}，日期 {day_str}）")

            for it in items:
                item_id = it["id"]
                locale = it.get("locale") or "zh"
                kind = it.get("kind") or ""
                sha = it.get("sha256") or ""
                cur.execute(
                    """
                    SELECT id FROM sign_stroke_item
                    WHERE signer_id=%s AND locale=%s AND kind=%s AND sha256=%s
                    """,
                    (to_id, locale, kind, sha),
                )
                dup = cur.fetchone()
                if dup and str(dup["id"]) != str(item_id):
                    stats["items_replaced_target"] += 1
                    if apply:
                        cur.execute(
                            "DELETE FROM sign_stroke_item WHERE id=%s",
                            (dup["id"],),
                        )
                if apply:
                    cur.execute(
                        "UPDATE sign_stroke_item SET signer_id=%s WHERE id=%s",
                        (to_id, item_id),
                    )
                    cur.execute(
                        """
                        UPDATE sign_file_role_signer
                        SET signer_id=%s
                        WHERE signer_id=%s
                          AND (sig_item_id=%s OR date_item_id=%s)
                        """,
                        (to_id, from_id, item_id, item_id),
                    )
                    stats["role_rows_updated"] += cur.rowcount
                stats["items_moved"] += 1

            cur.execute(
                """
                SELECT id, locale, sig_sha256, date_sha256, updated_at
                FROM sign_stroke_set
                WHERE signer_id=%s AND DATE(updated_at)=%s
                """,
                (from_id, day_str),
            )
            sets = list(cur.fetchall() or [])
            print(f"待迁移 stroke_set：{len(sets)} 条")
            for st in sets:
                set_id = st["id"]
                locale = st.get("locale") or "zh"
                sig_sha = st.get("sig_sha256") or ""
                date_sha = st.get("date_sha256") or ""
                cur.execute(
                    """
                    SELECT id FROM sign_stroke_set
                    WHERE signer_id=%s AND locale=%s
                      AND sig_sha256=%s AND date_sha256=%s
                    """,
                    (to_id, locale, sig_sha, date_sha),
                )
                dup = cur.fetchone()
                if dup and str(dup["id"]) != str(set_id):
                    stats["sets_replaced_target"] += 1
                    if apply:
                        cur.execute("DELETE FROM sign_stroke_set WHERE id=%s", (dup["id"],))
                if apply:
                    cur.execute(
                        "UPDATE sign_stroke_set SET signer_id=%s WHERE id=%s",
                        (to_id, set_id),
                    )
                    cur.execute(
                        """
                        UPDATE sign_file_role_signer
                        SET signer_id=%s
                        WHERE signer_id=%s AND stroke_set_id=%s
                        """,
                        (to_id, from_id, set_id),
                    )
                    stats["role_rows_updated"] += cur.rowcount
                stats["sets_moved"] += 1

            cur.execute(
                "SELECT * FROM sign_signer_stroke WHERE signer_id=%s",
                (from_id,),
            )
            leg_from = cur.fetchone()
            if leg_from and apply:
                cur.execute(
                    "SELECT signer_id FROM sign_signer_stroke WHERE signer_id=%s",
                    (to_id,),
                )
                leg_to = cur.fetchone()
                cols = [
                    c
                    for c in (
                        "sig_ftp_path",
                        "date_ftp_path",
                        "sig_size",
                        "date_size",
                        "sig_sha256",
                        "date_sha256",
                        "sig_png",
                        "date_png",
                        "ftp_last_error",
                    )
                    if c in leg_from
                ]
                if leg_to:
                    set_parts = [f"{c}=%s" for c in cols]
                    vals = [leg_from.get(c) for c in cols] + [to_id]
                    cur.execute(
                        f"UPDATE sign_signer_stroke SET {', '.join(set_parts)} WHERE signer_id=%s",
                        vals,
                    )
                else:
                    fields = ["signer_id"] + cols
                    placeholders = ", ".join(["%s"] * len(fields))
                    vals = [to_id] + [leg_from.get(c) for c in cols]
                    cur.execute(
                        f"INSERT INTO sign_signer_stroke ({', '.join(fields)}) VALUES ({placeholders})",
                        vals,
                    )
                cur.execute(
                    "DELETE FROM sign_signer_stroke WHERE signer_id=%s", (from_id,)
                )
                stats["legacy_merged"] = True
            elif leg_from and not apply:
                stats["legacy_merged"] = True

        if not apply:
            conn.rollback()
        else:
            print("已提交事务。")

    return stats


def main() -> None:
    p = argparse.ArgumentParser(description="迁移指定签署人当天笔迹到另一签署人")
    p.add_argument("--from-name", required=True, help="源签署人姓名")
    p.add_argument("--to-name", required=True, help="目标签署人姓名")
    p.add_argument(
        "--date",
        default=date.today().isoformat(),
        help="按 updated_at 日期筛选（默认今天，YYYY-MM-DD）",
    )
    p.add_argument(
        "--apply",
        action="store_true",
        help="实际写入；默认仅预览（dry-run）",
    )
    args = p.parse_args()
    on_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    stats = migrate(
        args.from_name,
        args.to_name,
        on_date,
        apply=bool(args.apply),
    )
    mode = "已执行" if args.apply else "预览"
    print(
        f"\n[{mode}] items_moved={stats['items_moved']} "
        f"items_replaced_target={stats['items_replaced_target']} "
        f"sets_moved={stats['sets_moved']} legacy_merged={stats['legacy_merged']} "
        f"role_rows_updated={stats['role_rows_updated']}"
    )
    if not args.apply:
        print("未写入数据库。确认无误后加 --apply 再执行。")


if __name__ == "__main__":
    main()
