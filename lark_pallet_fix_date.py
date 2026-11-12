#!/usr/bin/env python3
"""Fix a wrong Daily Measurement Date in the Lark Pallet record.

Yan's Lark import labelled the first day's snapshot 2026-10-30
instead of 2026-11-01 (142 rows). This script rewrites those rows'
Daily Measurement Date to 2026-11-01 and recomputes their M-key
(serial 46327 + normalised Batch#) so the date and the key agree.

Usage (GitHub Actions, secrets LARK_APP_ID / LARK_APP_SECRET /
LARK_APP_TOKEN):
    python3 lark_pallet_fix_date.py --dry-run
    python3 lark_pallet_fix_date.py            # real run
    python3 lark_pallet_fix_date.py --date-from 2026-10-30 \
        --date-to 2026-11-01 --dry-run

The script only touches rows whose Daily Measurement Date equals
--date-from. It reports how many rows it found, updated and
verified. A re-run after a successful fix finds 0 rows and
changes nothing (idempotent).
"""
import argparse
import json
import math
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

HKT = timezone(timedelta(hours=8))
LARK_API = "https://open.larksuite.com/open-apis"
PALLET_RECORD_TABLE_ID = "tblPGvy6n3bApbiN"

F_DMD = "Daily Measurement Date"
F_BATCH = ("Batch# <BU>-<Date Code>-<Packaging Type>"
           "<3 Digit Sequential Number>")
F_MKEY = ("Daily Measurement DateBatch# <BU>-<Date Code>-"
          "<Packaging Type><3 Digit Sequential Number>")

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")


def log(msg):
    print(f"[pallet-fix-date] {msg}", flush=True)


def http_json(method, url, body=None, headers=None, timeout=120):
    data = json.dumps(body).encode() if body is not None else None
    merged = {"User-Agent": USER_AGENT, "Content-Type": "application/json"}
    merged.update(headers or {})
    req = urllib.request.Request(url, method=method, data=data,
                                 headers=merged)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.load(resp)
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:2000]
        raise RuntimeError(f"HTTP {method} {url} {e.code}: {detail}")


def lark_token(app_id, app_secret):
    r = http_json("POST", LARK_API + "/auth/v3/tenant_access_token/internal",
                  {"app_id": app_id, "app_secret": app_secret})
    if r.get("code") != 0 or not r.get("tenant_access_token"):
        raise RuntimeError(f"Lark tenant token failed: {r}")
    return r["tenant_access_token"]


def norm_key(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def field_by_name(fields, name):
    if name in fields:
        return fields[name]
    want = norm_key(name)
    for k, v in fields.items():
        if norm_key(k) == want:
            return v
    return None


def plain_text(value):
    if isinstance(value, list):
        return "".join(seg.get("text", "") if isinstance(seg, dict)
                       else str(seg) for seg in value)
    if isinstance(value, dict):
        return str(value.get("text", "") or value.get("name", "") or "")
    return str(value or "")


def parse_date_ms(raw):
    """Bitable date fields arrive as epoch ms; return YYYY-MM-DD
    (HKT) or None if unparseable."""
    try:
        if isinstance(raw, (int, float)):
            return datetime.fromtimestamp(raw / 1000, tz=HKT).strftime(
                "%Y-%m-%d")
        s = str(raw or "").strip()
        if not s:
            return None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
    except (ValueError, OverflowError, OSError):
        return None
    return None


def date_to_ms(date_str):
    """YYYY-MM-DD -> epoch ms at midnight HKT (how Bitable Date
    fields store a date)."""
    dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=HKT)
    return int(dt.timestamp() * 1000)


def excel_serial(date_str):
    """Excel serial for a YYYY-MM-DD date (serial = days since
    1899-12-30). Used for the M-key column."""
    d = datetime.strptime(date_str, "%Y-%m-%d").date()
    return (d - datetime(1899, 12, 30).date()).days


def norm_batch(batch):
    b = str(batch or "").strip().upper()
    compact = re.sub(r"\s+", "", b)
    m = re.match(r"^(UPG-\d{8}-)(P?)(\d{3})$", compact)
    if m:
        return m.group(1) + "P" + m.group(3)
    return b


def list_records(token, app_token, table_id, max_pages=100):
    rows, page_token, pages = [], None, 0
    seen_ids = set()
    while True:
        qs = {"page_size": 100}
        if page_token:
            qs["page_token"] = page_token
        url = (f"{LARK_API}/bitable/v1/apps/{app_token}/tables/"
               f"{table_id}/records?{urllib.parse.urlencode(qs)}")
        r = http_json("GET", url, None,
                      {"Authorization": f"Bearer {token}"})
        if r.get("code") != 0:
            raise RuntimeError(f"Pallet record list failed: {r}")
        data = r.get("data") or {}
        items = data.get("items") or []
        if pages == 0:
            log(f"Pallet record table total rows: {data.get('total')}")
        new_this_page = 0
        for item in items:
            rid = item.get("record_id")
            if rid:
                if rid in seen_ids:
                    continue
                seen_ids.add(rid)
                new_this_page += 1
            rows.append(item)
        pages += 1
        if new_this_page == 0 and data.get("has_more"):
            log("WARNING: page returned no new rows but has_more "
                "is true — stopping the scan.")
            break
        if not data.get("has_more") or pages >= max_pages:
            break
        page_token = data.get("page_token")
        time.sleep(0.5)
    log(f"Pallet record rows read: {len(rows)} distinct "
        f"({pages} page(s))")
    return rows


def batch_update(token, app_token, table_id, records):
    """Update up to 100 records per call (Bitable batch_update)."""
    url = (f"{LARK_API}/bitable/v1/apps/{app_token}/tables/"
           f"{table_id}/records/batch_update")
    for i in range(0, len(records), 100):
        chunk = records[i:i + 100]
        r = http_json("POST", url, {"records": chunk},
                      {"Authorization": f"Bearer {token}"})
        if r.get("code") != 0:
            raise RuntimeError(f"batch_update failed: {r}")
        log(f"Updated {len(chunk)} record(s) "
            f"({i + len(chunk)}/{len(records)})")
        time.sleep(0.5)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date-from", default="2026-10-30",
                    help="Wrong date to fix (default 2026-10-30)")
    ap.add_argument("--date-to", default="2026-11-01",
                    help="Correct date (default 2026-11-01)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report only; write nothing")
    args = ap.parse_args()

    app_id = os.environ["LARK_APP_ID"]
    app_secret = os.environ["LARK_APP_SECRET"]
    app_token = os.environ["LARK_APP_TOKEN"]

    token = lark_token(app_id, app_secret)
    log("Tenant token ok")
    rows = list_records(token, app_token, PALLET_RECORD_TABLE_ID)

    target_ms = date_to_ms(args.date_to)
    target_serial = excel_serial(args.date_to)
    log(f"Fixing {args.date_from} -> {args.date_to} "
        f"(epoch ms {target_ms}, serial {target_serial})")

    to_fix = []
    for item in rows:
        fields = item.get("fields") or {}
        dmd = parse_date_ms(field_by_name(fields, F_DMD))
        if dmd != args.date_from:
            continue
        batch = norm_batch(plain_text(field_by_name(fields, F_BATCH)))
        if not batch:
            log(f"Skipping {item.get('record_id')}: blank Batch#")
            continue
        to_fix.append({
            "record_id": item["record_id"],
            "fields": {
                F_DMD: target_ms,
                F_MKEY: f"{target_serial}{batch}",
            },
        })
    log(f"Rows with Daily Measurement Date = {args.date_from}: "
        f"{len(to_fix)}")

    if args.dry_run:
        print(json.dumps({
            "dry_run": True,
            "date_from": args.date_from,
            "date_to": args.date_to,
            "rows_to_fix": len(to_fix),
        }, indent=2, ensure_ascii=False))
        return 0

    if to_fix:
        batch_update(token, app_token, PALLET_RECORD_TABLE_ID, to_fix)
    else:
        log("Nothing to fix.")

    # Verify by re-reading.
    rows2 = list_records(token, app_token, PALLET_RECORD_TABLE_ID)
    still_wrong = sum(
        1 for item in rows2
        if parse_date_ms(field_by_name(item.get("fields") or {},
                                       F_DMD)) == args.date_from)
    now_right = sum(
        1 for item in rows2
        if parse_date_ms(field_by_name(item.get("fields") or {},
                                       F_DMD)) == args.date_to)
    log(f"Verify: still {args.date_from}: {still_wrong}; "
        f"now {args.date_to}: {now_right}")
    print(json.dumps({
        "dry_run": False,
        "date_from": args.date_from,
        "date_to": args.date_to,
        "rows_fixed": len(to_fix),
        "still_wrong": still_wrong,
        "now_correct": now_right,
    }, indent=2, ensure_ascii=False))
    if still_wrong:
        log("WARNING: some rows still carry the old date.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
