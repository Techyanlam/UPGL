#!/usr/bin/env python3
"""UPG Pallet — nightly 20:00 HKT sync: website -> Lark Base.

Direction is the REVERSE of lark_pallet_sync.py (which reads Dock
Receipt and puts new pallets in the website's pending queue). This
script copies the website's CURRENT in-stock table into the Base
every night, plus one row per pallet departed that day (with Ship
Date), plus the website's activity log.

Destination tables in Base "UPGL inbound and cycle count"
(app_token QDo5bSOZraNgEhsRC7flOUMggPe; table ids are in the Base
URLs, so they are hardcoded here — they are not secrets):
  - "Pallet record" tblPGvy6n3bApbiN — one row per in-stock pallet
    per night, plus one row per pallet departed THAT night (with
    Ship Date = the departure date, so the night's snapshot
    matches the Excel convention: a departed pallet appears on
    its departure day with Ship Date and Remark 「最後」, then
    disappears). Append-only history; the same pallet appears
    once per night, and the Daily Measurement Date differs
    between nights. The 4 pre-existing BLANK rows are left
    alone; nothing is ever deleted from this table.
  - "活動記錄" tbl3IEmUb8gRSM84 — one row per activity entry, text
    only. The 5 pre-existing BLANK rows are left alone; nothing
    is ever deleted from this table.

Semantics (Yan, 2026-11-12 00:52; departed rows added 2026-11-12):
  - Pallet record: keep old rows, add new ones. A run appends one
    row per in-stock pallet for the sync-night date, plus one row
    per pallet departed on the sync-night date (Ship Date =
    departure date).
  - 活動記錄： append new entries line by line; never duplicate
    one already in the table.

Idempotency (a manual re-run on the same night must not double-add):
  - Pallet record: a pallet is skipped if a row with the same
    normalized Batch# AND the same Daily Measurement Date already
    exists in the table.
  - 活動記錄： an entry is skipped if its exact display text is
    already a row in the table (compared stripped).

Row mapping (mirrors the website's Excel export in app.js
renderExcel/exportCsv, including its quirks):
  - Daily Measurement Date = sync-night date (not the pallet's
    addedDate — each night's snapshot is dated that night).
  - Timestamp column = the pallet's LOCATION (the website's export
    puts location in the Timestamp column; replicated here).
  - Dock Receipt Date = dockReceiptDate.
  - Business Unit / Reference / Batch# (palletId as-is).
  - Date of Measurement = addedDate, falling back to
    dockReceiptDate.
  - Length / Width / Height (cm); Volume is RECOMPUTED as
    ROUNDUP(L*W*H/1e6, 1 dp) — the Base's Volume column is never
    copied (it has wrong values like 72 / 86).
  - Ship Date = blank for in-stock rows; = the departure date
    for rows of pallets departed on the sync-night date.
  - M-key column = Excel serial of the Daily Measurement Date
    (days since 1899-12-30; verified 2026-11-12 = 46338)
    concatenated with the Batch#, as text.
  - Activity display text is composed EXACTLY as the website shows
    it (app.js renderActivity): timestamp slice(0,16) with T->space,
    full-width space U+3000 separators, label map, optional
    detail/message/reason.

Environment (all from GitHub repo Secrets in the workflow, except
WORKER_URL which is public):
  LARK_APP_ID, LARK_APP_SECRET, LARK_APP_TOKEN, WORKER_URL.
  --dry-run lists what would be appended and writes nothing.
  --sync-date YYYY-MM-DD overrides the sync-night date (default:
  today HKT); useful for a backfill or a re-run for a past night.
"""

import argparse
import http.client
import json
import math
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone

HKT = timezone(timedelta(hours=8))
LARK_API = "https://open.larksuite.com/open-apis"

# Destination tables (table ids appear in Base URLs — not secrets).
PALLET_RECORD_TABLE_ID = "tblPGvy6n3bApbiN"
ACTIVITY_TABLE_ID = "tbl3IEmUb8gRSM84"

# Pallet record field names (must match the Base table exactly).
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

# Activity table field name (single text column).
F_TEXT = "文字"

DEFAULT_MAX_PAGES = 100
BATCH_CHUNK = 100

# Cloudflare's edge in front of workers.dev blocks the default urllib
# User-Agent ("Python-urllib/x.y") with 403 — the Worker itself only
# checks Origin, and urllib sends none (which passes). Proven
# 2026-11-12: curl with Python-urllib UA -> 403, browser UA -> 200.
# Send a browser User-Agent on every request the script makes.
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")

# Website activity label map (app.js ACT_LABEL).
ACT_LABEL = {
    "depart": "出倉", "updateHeight": "改高度",
    "updateLocation": "改位置", "add": "入倉", "restore": "還原",
    "pending": "待確認", "confirm": "確認加入", "reject": "唔加入",
}
FWSP = "　"  # U+3000 full-width space, the website's separator.


def log(msg):
    print(f"[pallet-record-sync] {msg}", flush=True)


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
    # Lark rejects this endpoint with code 10003 "invalid param" if the
    # body is not sent as application/json.
    r = http_json("POST", LARK_API + "/auth/v3/tenant_access_token/internal",
                  {"app_id": app_id, "app_secret": app_secret},
                  {"Content-Type": "application/json"})
    if r.get("code") != 0 or not r.get("tenant_access_token"):
        raise RuntimeError(f"Lark tenant token failed: {r}")
    return r["tenant_access_token"]


def lark_call(method, path, body, token):
    return http_json(method, LARK_API + path, body,
                     {"Content-Type": "application/json",
                      "Authorization": f"Bearer {token}"})


def plain_text(value):
    """Bitable text fields arrive as rich-text segment lists
    ([{"text": "...", "type": "text"}]); flatten to plain text."""
    if isinstance(value, list):
        return "".join(seg.get("text", "") if isinstance(seg, dict)
                       else str(seg) for seg in value)
    if isinstance(value, dict):
        return str(value.get("text", "") or value.get("name", "") or "")
    return str(value or "")


def norm_key(s):
    """Collapse all whitespace runs to a single space (and strip)."""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def field_by_name(fields, name):
    """Look a field up by name, tolerating extra whitespace in the
    table's actual field name. PROVEN 2026-11-12 on Dock Receipt:
    its Batch# key is 'Batch#' + MANY spaces + '<BU>-...' — an exact
    lookup with the single-space display name returns None for every
    row. Exact match first, then a whitespace-collapsed match."""
    if name in fields:
        return fields[name]
    want = norm_key(name)
    for k, v in fields.items():
        if norm_key(k) == want:
            return v
    return None


def parse_date_ms(raw):
    """Bitable date fields arrive as epoch ms; return YYYY-MM-DD (HKT)
    or None if unparseable."""
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


def date_to_ms(datestr):
    """YYYY-MM-DD -> epoch ms at 00:00 HKT (Lark date fields store
    ms; midnight HKT keeps the displayed date correct in HKT)."""
    d = datetime.strptime(datestr, "%Y-%m-%d").date()
    dt = datetime(d.year, d.month, d.day, tzinfo=HKT)
    return int(dt.timestamp() * 1000)


def excel_serial(d):
    """Excel serial: days since 1899-12-30 (verified:
    2026-09-30 = 46295, 2026-11-12 = 46338)."""
    return (d - date(1899, 12, 30)).days


def calc_volume(length_cm, width_cm, height_cm):
    """ROUNDUP(L*W*H/1e6, 1 dp), same as the Worker and app.js
    (epsilon guards against float error pushing an exact value up)."""
    raw = (length_cm * width_cm * height_cm) / 1e6
    return math.ceil(raw * 10 - 1e-9) / 10


def normalize_batch(s):
    """Batch# dedupe key: strip all whitespace, uppercase."""
    return re.sub(r"\s+", "", str(s or "")).upper()


def activity_text(a):
    """Compose one activity entry's display string EXACTLY as the
    website shows it (app.js renderActivity)."""
    ts = a.get("timestamp")
    if ts:
        when = str(ts)[:16].replace("T", " ")
    else:
        when = (a.get("datetime") or a.get("time") or
                ((a.get("date") or "") +
                 (" " + a["time"] if a.get("time") else "")))
    label = (ACT_LABEL.get(a.get("type")) or a.get("type")
             or a.get("action") or "活動")
    text = (when + FWSP if when else "") + label
    if a.get("palletId"):
        text += FWSP + str(a["palletId"])
    if a.get("detail"):
        text += FWSP + str(a["detail"])
    if a.get("message"):
        text += FWSP + str(a["message"])
    if a.get("reason"):
        text += "（原因：" + str(a["reason"]) + "）"
    return text


def get_worker_state(worker_url):
    """Full Worker /state, with retries for transient read failures
    (truncated responses seen through proxies)."""
    last_err = None
    for _attempt in range(3):
        try:
            req = urllib.request.Request(
                worker_url.rstrip("/") + "/state",
                headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=60) as resp:
                d = json.load(resp)
            break
        except (urllib.error.URLError, http.client.IncompleteRead,
                ConnectionError, TimeoutError) as e:
            last_err = e
            log(f"Worker /state read failed ({e}); retrying...")
            time.sleep(2)
    else:
        raise RuntimeError(f"Worker /state read failed: {last_err}")
    log(f"Worker state: {d.get('inStockCount')} in stock, "
        f"{len(d.get('departed') or [])} departed, "
        f"{len(d.get('pending') or [])} pending, "
        f"{len(d.get('activity') or [])} activity entries")
    return d


def list_records(token, app_token, table_id, max_pages, label):
    """All rows of a table, paginated via the GET list-records
    endpoint (never POST /records/search for a full-table scan —
    its page_token does not advance and re-reads page 1 forever,
    proven 2026-11-12)."""
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
            if items:
                f0 = items[0].get("fields") or {}
                log(f"{label} sample row field keys: "
                    f"{sorted(f0.keys())}")
        new_this_page = 0
        dupes = 0
        for item in items:
            rid = item.get("record_id")
            if rid:
                if rid in seen_ids:
                    dupes += 1
                    continue
                seen_ids.add(rid)
                new_this_page += 1
            rows.append(item)
        if dupes:
            log(f"{label} page {pages + 1}: dropped {dupes} "
                f"duplicate row(s)")
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
    if pages >= max_pages:
        log(f"WARNING: hit MAX_PAGES={max_pages} — older rows were not "
            f"scanned. Raise SYNC_MAX_PAGES if the table keeps growing.")
    return rows


def batch_create(token, app_token, table_id, records, label):
    """Append records via /records/batch_create, in chunks of
    BATCH_CHUNK. Any API failure raises (exit non-zero)."""
    total = 0
    for start in range(0, len(records), BATCH_CHUNK):
        chunk = records[start:start + BATCH_CHUNK]
        r = lark_call(
            "POST",
            f"/bitable/v1/apps/{app_token}/tables/{table_id}"
            f"/records/batch_create",
            {"records": [{"fields": f} for f in chunk]}, token)
        if r.get("code") != 0:
            raise RuntimeError(
                f"{label} batch_create failed "
                f"(records {start + 1}-{start + len(chunk)}): {r}")
        total += len(chunk)
        log(f"{label}: appended {len(chunk)} row(s) "
            f"({total}/{len(records)})")
    return total


def build_pallet_row(pallet, sync_date, ship_date=None):
    """One Pallet record row for one pallet on the sync night.
    In-stock pallets get a blank Ship Date; a pallet departed on
    the sync night gets Ship Date = its departure date.
    Returns (fields dict, skip_reason or None)."""
    pallet_id = str(pallet.get("palletId") or "").strip()
    if not pallet_id:
        return None, "empty palletId"
    height = pallet.get("heightCm")
    if (not isinstance(height, (int, float))
            or not (0 < height <= 1000)):
        return None, f"bad Height {height!r}"
    length = pallet.get("lengthCm")
    if (not isinstance(length, (int, float))
            or not (0 < length <= 1000)):
        length = 120
    width = pallet.get("widthCm")
    if (not isinstance(width, (int, float))
            or not (0 < width <= 1000)):
        width = 100
    fields = {
        F_DMD: date_to_ms(sync_date),
        F_TS: str(pallet.get("location") or ""),
        F_BU: str(pallet.get("businessUnit") or ""),
        F_REF: str(pallet.get("reference") or ""),
        F_BATCH: pallet_id,
        F_LEN: length,
        F_WID: width,
        F_HEI: height,
        F_VOL: calc_volume(length, width, height),
        F_MKEY: f"{excel_serial(datetime.strptime(sync_date, '%Y-%m-%d').date())}{pallet_id}",
    }
    dock = parse_date_ms(pallet.get("dockReceiptDate"))
    if dock:
        fields[F_DOCK] = date_to_ms(dock)
    dom_src = pallet.get("addedDate") or pallet.get("dockReceiptDate")
    dom = parse_date_ms(dom_src)
    if dom:
        fields[F_DOM] = date_to_ms(dom)
    # Ship Date: blank for in-stock pallets, the departure date for
    # pallets departed on the sync night. Empty-valued fields are
    # omitted (Lark treats an omitted field as blank).
    if ship_date:
        fields[F_SHIP] = date_to_ms(ship_date)
    fields = {k: v for k, v in fields.items()
              if v != "" and v is not None}
    return fields, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="List what would be appended, write nothing")
    ap.add_argument("--sync-date", default=None,
                    help="Sync-night date YYYY-MM-DD (default: today "
                         "HKT)")
    ap.add_argument("--max-pages", type=int, default=None,
                    help="Max Base pages to scan per table (default "
                         "SYNC_MAX_PAGES or 100)")
    args = ap.parse_args()

    max_pages = args.max_pages if args.max_pages is not None else int(
        os.environ.get("SYNC_MAX_PAGES", DEFAULT_MAX_PAGES))
    if args.sync_date:
        try:
            sync_d = datetime.strptime(
                args.sync_date, "%Y-%m-%d").date()
        except ValueError:
            raise RuntimeError(
                f"Bad --sync-date {args.sync_date!r} (want YYYY-MM-DD)")
    else:
        sync_d = datetime.now(HKT).date()
    sync_date = sync_d.strftime("%Y-%m-%d")

    app_id = os.environ["LARK_APP_ID"]
    app_secret = os.environ["LARK_APP_SECRET"]
    app_token = os.environ["LARK_APP_TOKEN"]
    worker_url = os.environ["WORKER_URL"]

    token = lark_token(app_id, app_secret)
    log("Tenant token ok")
    log(f"Sync date (night): {sync_date}")

    # --- Worker state -------------------------------------------------
    state = get_worker_state(worker_url)
    in_stock = state.get("inStock") or []
    activity = state.get("activity") or []

    # --- Pallet record: existing (Batch#, date) pairs -----------------
    existing_rows = list_records(token, app_token,
                                 PALLET_RECORD_TABLE_ID, max_pages,
                                 "Pallet record")
    existing_pairs = set()
    for row in existing_rows:
        f = row.get("fields") or {}
        batch = normalize_batch(
            plain_text(field_by_name(f, F_BATCH)).strip())
        dmd = parse_date_ms(field_by_name(f, F_DMD))
        if batch and dmd:
            existing_pairs.add((batch, dmd))
    log(f"Pallet record existing (Batch#, date) pairs: "
        f"{len(existing_pairs)}")

    to_append, skipped_pallets = [], []
    for pallet in in_stock:
        pallet_id = str(pallet.get("palletId") or "").strip()
        pair = (normalize_batch(pallet_id), sync_date)
        if pair in existing_pairs:
            skipped_pallets.append(
                (pallet_id, f"already present for {sync_date}"))
            continue
        fields, why = build_pallet_row(pallet, sync_date)
        if fields is None:
            skipped_pallets.append((pallet_id, why))
            continue
        to_append.append(fields)
    for batch, why in skipped_pallets:
        log(f"Pallet record skipped: {batch or '(no Batch#)'} — {why}")
    # Pallets departed ON the sync night: one row each, with Ship
    # Date = the departure date (Excel convention: a departed
    # pallet appears on its departure day, then disappears).
    # Anything departed on an earlier night already has its row
    # from that night's run (or the backfill) and is not re-added.
    departed = state.get("departed") or []
    departed_today = 0
    for pallet in departed:
        dep_date = str(pallet.get("departedDate") or "").strip()
        if dep_date != sync_date:
            continue
        pallet_id = str(pallet.get("palletId") or "").strip()
        pair = (normalize_batch(pallet_id), sync_date)
        if pair in existing_pairs:
            skipped_pallets.append(
                (pallet_id, f"already present for {sync_date}"))
            continue
        fields, why = build_pallet_row(pallet, sync_date,
                                       ship_date=dep_date)
        if fields is None:
            skipped_pallets.append((pallet_id, why))
            continue
        to_append.append(fields)
        departed_today += 1
    log(f"Pallet record rows to append: {len(to_append)} "
        f"({[f[F_BATCH] for f in to_append]})")
    log(f"  of which departed on {sync_date}: {departed_today}")

    # --- 活動記錄: existing texts --------------------------------------
    activity_rows = list_records(token, app_token, ACTIVITY_TABLE_ID,
                                 max_pages, "活動記錄")
    existing_texts = set()
    for row in activity_rows:
        f = row.get("fields") or {}
        t = plain_text(f.get(F_TEXT)).strip()
        if t:
            existing_texts.add(t)
    log(f"活動記錄 existing texts: {len(existing_texts)}")

    to_append_act, skipped_act = [], []
    seen_this_run = set()
    for entry in activity:
        text = activity_text(entry)
        key = text.strip()
        if not key:
            skipped_act.append((text, "empty text"))
            continue
        if key in existing_texts or key in seen_this_run:
            skipped_act.append((text, "already in table"))
            continue
        seen_this_run.add(key)
        to_append_act.append({F_TEXT: text})
    for text, why in skipped_act:
        log(f"Activity skipped: {text!r} — {why}")
    log(f"Activity rows to append: {len(to_append_act)}")
    for rec in to_append_act:
        log(f"Activity to append: {rec[F_TEXT]!r}")

    # --- Write (unless dry run) ---------------------------------------
    appended_pallets = 0
    appended_act = 0
    if args.dry_run:
        log("Dry run: nothing written to Lark.")
    else:
        if to_append:
            appended_pallets = batch_create(
                token, app_token, PALLET_RECORD_TABLE_ID, to_append,
                "Pallet record")
        if to_append_act:
            appended_act = batch_create(
                token, app_token, ACTIVITY_TABLE_ID, to_append_act,
                "活動記錄")

    print(json.dumps({
        "dry_run": args.dry_run,
        "sync_date": sync_date,
        "pallet_rows_appended": appended_pallets if not args.dry_run
                                 else len(to_append),
        "pallet_rows_skipped": len(skipped_pallets),
        "activity_appended": appended_act if not args.dry_run
                             else len(to_append_act),
        "activity_skipped": len(skipped_act),
    }, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
