#!/usr/bin/env python3
"""Delete all Pallet record rows for one Batch# (one-off cleanup).

Yan confirmed DRAFTSHELF-2 is not a real pallet: it has 30 rows in
the Lark Pallet record (one per day 2026-10-13 .. 2026-11-11) but
appears in neither the October nor the November Excel file. This
script deletes exactly those rows by record_id.

Safety:
- It only touches rows whose normalised Batch# equals --batch
  exactly (DRAFTSHELF is a different pallet and is untouched).
- It refuses to delete unless the found count equals
  --expected-count (default 30, the count Yan authorised).
  A different count means the table changed since the decision
  was made: stop and check instead of deleting the wrong set.
- Dry run reports the count and the dates; the real run verifies
  by re-reading (0 matching rows must remain).

Usage (GitHub Actions, secrets LARK_APP_ID / LARK_APP_SECRET /
LARK_APP_TOKEN):
    python3 lark_pallet_delete_batch.py --dry-run
    python3 lark_pallet_delete_batch.py            # real run
"""
import argparse
import json
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

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")


def log(msg):
    print(f"[pallet-delete-batch] {msg}", flush=True)


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


def batch_delete(token, app_token, table_id, record_ids):
    """Delete up to 100 records per call (Bitable batch_delete)."""
    url = (f"{LARK_API}/bitable/v1/apps/{app_token}/tables/"
           f"{table_id}/records/batch_delete")
    for i in range(0, len(record_ids), 100):
        chunk = record_ids[i:i + 100]
        r = http_json("POST", url, {"records": chunk},
                      {"Authorization": f"Bearer {token}"})
        if r.get("code") != 0:
            raise RuntimeError(f"batch_delete failed: {r}")
        log(f"Deleted {len(chunk)} record(s) "
            f"({i + len(chunk)}/{len(record_ids)})")
        time.sleep(0.5)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--batch", default="DRAFTSHELF-2",
                    help="Batch# whose rows to delete")
    ap.add_argument("--expected-count", type=int, default=30,
                    help="Refuse to delete unless this many rows match")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report only; delete nothing")
    args = ap.parse_args()

    app_id = os.environ["LARK_APP_ID"]
    app_secret = os.environ["LARK_APP_SECRET"]
    app_token = os.environ["LARK_APP_TOKEN"]

    target = norm_batch(args.batch)
    token = lark_token(app_id, app_secret)
    log("Tenant token ok")
    rows = list_records(token, app_token, PALLET_RECORD_TABLE_ID)

    matched = []
    for item in rows:
        fields = item.get("fields") or {}
        batch = norm_batch(plain_text(field_by_name(fields, F_BATCH)))
        if batch != target:
            continue
        matched.append({
            "record_id": item["record_id"],
            "date": parse_date_ms(field_by_name(fields, F_DMD)),
        })
    dates = sorted(m["date"] or "?" for m in matched)
    log(f"Rows with Batch# = {target}: {len(matched)} "
        f"(dates {dates[0] if dates else '-'} .. "
        f"{dates[-1] if dates else '-'})")
    if args.dry_run:
        print(json.dumps({
            "dry_run": True,
            "batch": target,
            "rows_to_delete": len(matched),
            "expected_count": args.expected_count,
            "dates": dates,
        }, indent=2, ensure_ascii=False))
        if len(matched) != args.expected_count:
            log(f"COUNT MISMATCH: found {len(matched)}, expected "
                f"{args.expected_count} — do not delete, check first.")
            return 1
        return 0

    if len(matched) != args.expected_count:
        log(f"REFUSING TO DELETE: found {len(matched)} rows, "
            f"expected {args.expected_count}. Nothing deleted.")
        return 1

    batch_delete(token, app_token, PALLET_RECORD_TABLE_ID,
                 [m["record_id"] for m in matched])

    # Verify by re-reading.
    rows2 = list_records(token, app_token, PALLET_RECORD_TABLE_ID)
    remaining = sum(
        1 for item in rows2
        if norm_batch(plain_text(field_by_name(
            item.get("fields") or {}, F_BATCH))) == target)
    log(f"Verify: remaining rows with Batch# = {target}: {remaining}")
    print(json.dumps({
        "dry_run": False,
        "batch": target,
        "rows_deleted": len(matched),
        "remaining": remaining,
    }, indent=2, ensure_ascii=False))
    if remaining:
        log("WARNING: some matching rows still remain.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
