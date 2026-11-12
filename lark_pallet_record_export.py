#!/usr/bin/env python3
"""UPG Pallet — export the Lark "Pallet record" table to
daily-record.json (repo root), which the website's 每日記錄
(Excel) view reads.

The website's Excel view used to show the CURRENT stock list
(one row per pallet, filtered by addedDate) — so filtering by
11-10 showed only the 3 pallets added that day, not the 133 that
were actually in stock that day. A daily record needs one row per
pallet PER DAY, which only exists in the Lark Pallet record table
(nightly 20:00 record sync + the Nov 1-11 backfill). This script
reads that table and writes it out as a static JSON file the
website can fetch without any new secret.

Source table: Base "UPGL inbound and cycle count"
(app_token QDo5bSOZraNgEhsRC7flOUMggPe), "Pallet record"
tblPGvy6n3bApbiN. The 4 pre-existing BLANK rows (no Batch#)
are skipped; nothing is ever written to Lark by this script.

Rules (same as the record sync / backfill):
  - Volume is RECOMPUTED as ROUNDUP(L*W*H/1e6, 1 dp) from the
    Length/Width/Height columns — the table's Volume column is
    never copied (it has wrong values like 72 / 86, and Yan's
    manual edits may have changed it).
  - Dates are written as YYYY-MM-DD strings (HKT).
  - Ship Date is a TEXT column ('YYYY-MM-DD' or '') — a row with
    a Ship Date is a departed pallet; blank means in stock that
    day. Remark is NOT stored in the table; the website derives
    it ('最後' when Ship Date is set).
  - Rows are sorted by (Daily Measurement Date, Batch#).

Output (daily-record.json, repo root):
  {"generatedAt": "<ISO UTC>", "rows": [ {dmd, location,
   dockReceiptDate, businessUnit, reference, palletId, dom,
   lengthCm, widthCm, heightCm, volumeCbm, shipDate, mkey}, ... ],
   "count": N}
  Field names match the website's pallet object where possible,
  so app.js needs only a data-source swap.

Environment (GitHub repo Secrets in the workflow):
  LARK_APP_ID, LARK_APP_SECRET, LARK_APP_TOKEN.
  --out PATH overrides the output file (default daily-record.json).
  --max-pages caps the Base scan (default SYNC_MAX_PAGES or 100).
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
F_TS = "Timestamp"
F_DOCK = "Dock Receipt Date"
F_BU = "Business Unit"
F_REF = "Reference (PO# or other Reference#)"
F_BATCH = ("Batch# <BU>-<Date Code>-<Packaging Type>"
           "<3 Digit Sequential Number>")
F_DOM = "Date of Measurement"
F_LEN = "Length (cm)"
F_WID = "Width (cm)"
F_HEI = "Height (cm)"
F_VOL = "Volume (CBM)"
F_SHIP = "Ship Date"
F_MKEY = ("Daily Measurement DateBatch# <BU>-<Date Code>-"
          "<Packaging Type><3 Digit Sequential Number>")

DEFAULT_MAX_PAGES = 100
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")


def log(msg):
    print(f"[pallet-record-export] {msg}", flush=True)


def http_json(method, url, body=None, headers=None, timeout=120):
    data = json.dumps(body).encode() if body is not None else None
    merged = {"User-Agent": USER_AGENT}
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
                  {"app_id": app_id, "app_secret": app_secret},
                  {"Content-Type": "application/json"})
    if r.get("code") != 0 or not r.get("tenant_access_token"):
        raise RuntimeError(f"Lark tenant token failed: {r}")
    return r["tenant_access_token"]


def plain_text(value):
    """Bitable text fields arrive as rich-text segment lists;
    flatten to plain text."""
    if isinstance(value, list):
        return "".join(seg.get("text", "") if isinstance(seg, dict)
                       else str(seg) for seg in value)
    if isinstance(value, dict):
        return str(value.get("text", "") or value.get("name", "") or "")
    return str(value or "")


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


def to_number(raw, default=None):
    if isinstance(raw, (int, float)):
        return raw
    try:
        return float(str(raw).strip())
    except (TypeError, ValueError, AttributeError):
        return default


def calc_volume(length_cm, width_cm, height_cm):
    """ROUNDUP(L*W*H/1e6, 1 dp)."""
    raw = (length_cm * width_cm * height_cm) / 1e6
    return math.ceil(raw * 10 - 1e-9) / 10


def list_records(token, app_token, table_id, max_pages, label):
    """All rows of a table, paginated via GET list-records
    (never POST /records/search for a full-table scan)."""
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
            raise RuntimeError(f"{label} records list failed: {r}")
        data = r.get("data") or {}
        items = data.get("items") or []
        if pages == 0:
            log(f"{label} table total rows: {data.get('total')}")
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
            log(f"WARNING: {label} page {pages} returned no new rows "
                f"but has_more is true — pagination is not advancing; "
                f"stopping the scan.")
            break
        if not data.get("has_more") or pages >= max_pages:
            break
        page_token = data.get("page_token")
        time.sleep(0.5)
    log(f"{label} rows read: {len(rows)} distinct ({pages} page(s))")
    return rows


def build_row(fields):
    """One daily-record row from one Pallet record row's fields.
    Returns (row dict, skip_reason or None)."""
    batch = plain_text(field_by_name(fields, F_BATCH)).strip()
    if not batch:
        return None, "blank Batch#"
    dmd = parse_date_ms(field_by_name(fields, F_DMD)) or ""
    length = to_number(field_by_name(fields, F_LEN), 120)
    width = to_number(field_by_name(fields, F_WID), 100)
    height = to_number(field_by_name(fields, F_HEI))
    if height is None or not (0 < height <= 1000):
        return None, f"bad Height {height!r}"
    mkey = plain_text(field_by_name(fields, F_MKEY)).strip()
    row = {
        "dmd": dmd,
        "location": plain_text(field_by_name(fields, F_TS)).strip(),
        "dockReceiptDate":
            parse_date_ms(field_by_name(fields, F_DOCK)) or "",
        "businessUnit":
            plain_text(field_by_name(fields, F_BU)).strip(),
        "reference":
            plain_text(field_by_name(fields, F_REF)).strip(),
        "palletId": batch,
        "dom": parse_date_ms(field_by_name(fields, F_DOM)) or "",
        "lengthCm": length,
        "widthCm": width,
        "heightCm": height,
        "volumeCbm": calc_volume(length, width, height),
        "shipDate":
            plain_text(field_by_name(fields, F_SHIP)).strip(),
        "mkey": mkey,
    }
    return row, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="daily-record.json",
                    help="Output JSON path (default daily-record.json)")
    ap.add_argument("--max-pages", type=int, default=None,
                    help="Max Base pages to scan (default "
                         "SYNC_MAX_PAGES or 100)")
    args = ap.parse_args()
    max_pages = args.max_pages if args.max_pages is not None else int(
        os.environ.get("SYNC_MAX_PAGES", DEFAULT_MAX_PAGES))

    app_id = os.environ["LARK_APP_ID"]
    app_secret = os.environ["LARK_APP_SECRET"]
    app_token = os.environ["LARK_APP_TOKEN"]

    token = lark_token(app_id, app_secret)
    log("Tenant token ok")
    rows = list_records(token, app_token, PALLET_RECORD_TABLE_ID,
                        max_pages, "Pallet record")

    out_rows, skipped = [], []
    for item in rows:
        row, why = build_row(item.get("fields") or {})
        if row is None:
            skipped.append(why)
            continue
        out_rows.append(row)
    out_rows.sort(key=lambda r: (r["dmd"], r["palletId"]))
    from collections import Counter
    per_day = Counter(r["dmd"] for r in out_rows)
    log(f"Exported rows: {len(out_rows)} "
        f"(skipped {len(skipped)}: {Counter(skipped)})")
    for d in sorted(per_day):
        log(f"  {d or '(no date)'}: {per_day[d]}")

    payload = {
        "generatedAt": datetime.now(timezone.utc).strftime(
            "%Y-%m-%dT%H:%M:%SZ"),
        "rows": out_rows,
        "count": len(out_rows),
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    log(f"Wrote {args.out} ({len(out_rows)} rows)")
    print(json.dumps({
        "rows": len(out_rows),
        "skipped": len(skipped),
        "per_day": dict(sorted(per_day.items())),
    }, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
