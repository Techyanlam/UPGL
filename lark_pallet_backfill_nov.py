#!/usr/bin/env python3
"""UPG Pallet — November backfill: Excel cycle-count history -> Lark Base.

One-off backfill (Yan, 2026-11-12 12:03): the file
DOCK_RECEIPT_CYCLE_COUNT_11.xlsx holds the November daily
cycle-count snapshots, and Lark's "Pallet record" table is the
daily-history store (one snapshot set per day, keyed by the M-key
column). This script writes the November snapshots into Pallet
record so the history is complete and searchable by day in Lark.

Source file (sheet '2026', 13 columns, same layout as the website
Excel export / Pallet record):
  Daily Measurement Date, Timestamp (location), Dock Receipt Date,
  Business Unit, Reference, Batch#, Date of Measurement, Length,
  Width, Height, Volume, Ship Date, M-key.

Measured days (9): 2026-11-01, 11-02, 11-03, 11-04, 11-05, 11-06,
11-09, 11-10, 11-11. MISSING days: 11-07 (Sat) and 11-08 (Sun) —
no snapshots; Yan's rule: carry the previous day's records forward
(11-07 and 11-08 are both a copy of the 11-06 snapshot, with the
date and M-key changed to the carried day). Nov 1 (Sun) HAS a
snapshot — use it as-is. 11-12 is NOT written (already in Lark from
the 2026-11-12 02:30 upload; the nightly 20:00 sync completes it).

Dedupe (per day): DRAFT-PALLET appears twice identically — count
once. The UPG-20260827-P002/P003/P004 trio appears twice per day
(one row's Date of Measurement typo'd 2028-08-27, one correct
2026-08-31) — always use the 2026 row.

The file's M-key column is STALE AND WRONG (off by 2 days: it says
46325 for 2026-11-01, but the correct serial is 46327). NEVER copy
the file's M-key. It is recomputed as the Excel serial of the
snapshot date + normalized Batch#, as text. Serial =
(date - 1899-12-30).days; verified 11-01=46327 ... 11-11=46337,
11-12=46338.

The 11-11 snapshot's location / height / volume are ALL BROKEN:
397 cells are '#REF!' (Timestamp, Height, Volume). Other days are
clean. For 11-11 rows the file's location/height/volume are NOT
used. Instead:
  - the 129 in-stock pallets take height/location from the LIVE
    website /state (in-stock list, matched by normalized Batch#),
    falling back to the file's 11-10 row for the same Batch# if the
    pallet is not in live stock (reported per row);
  - the 4 departed on 11-11 (UPG-20261014-P012/P013,
    UPG-20261027-P006, UPG-20261105-P001) take height/location
    from the file's 11-10 snapshot rows for the same Batch#.
  - if a height/location still cannot be resolved, the row is
    SKIPPED and reported (never write '#REF!' or a blank height).

Normalization (same rules as the website's earlier corrections):
  Batch# -> strip spaces, uppercase, add a missing P
  (UPG-YYYYMMDD-NNN -> UPG-YYYYMMDD-PNNN). Dock Receipt Date for
  UPG- pallets = the Batch# date code (ground truth — the file's
  dock dates are raw/wrong for many rows). For non-UPG pallets
  (DRAFT-PALLET, DRAFTSHELF, PACKAGING MATERIALS-2) the file's dock
  date is used as-is (blank stays blank). Volume =
  ROUNDUP(L*W*H/1e6, 1 decimal) recomputed — never copy the file's
  Volume.

BOTH in-stock rows (Ship Date blank) AND departed rows (Ship Date
= the file's ship date) are written — Pallet record mirrors the
website Excel export, which includes departed rows. NOTE: this
differs from the nightly record-sync script's in-stock-only scope;
that is intentional for this backfill (the file's records include
departed pallets).

Column mapping (same as the nightly script / website Excel export):
  col1 Daily Measurement Date = snapshot date; col2 Timestamp =
  location; col3 Dock Receipt Date; col4 Business Unit; col5
  Reference; col6 Batch# (normalized); col7 Date of Measurement =
  snapshot date; col8-10 Length/Width/Height (cm); col11 Volume
  (recomputed ROUNDUP 1dp); col12 Ship Date ('' in stock, else the
  file's ship date — the column is a TEXT field, so the date is
  written as a 'YYYY-MM-DD' string); col13 M-key = serial(snapshot
  date) + normalized Batch# as text.

Idempotency: a row is skipped if a Pallet record row with the same
normalized Batch# AND the same Daily Measurement Date already
exists (same dedupe as the nightly script). A re-run appends
nothing.

Environment (from GitHub repo Secrets in the workflow):
  LARK_APP_ID, LARK_APP_SECRET, LARK_APP_TOKEN, WORKER_URL
  (WORKER_URL defaults to the public Worker URL if unset).
  --dry-run lists what would be appended and writes nothing.
  --offline builds the rows without contacting Lark (dedupe is
  then unverified — for local testing only; the script fails
  LOUDLY if Lark credentials are missing and --offline is not set).
"""

import argparse
import http.client
import json
import math
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone

HKT = timezone(timedelta(hours=8))
LARK_API = "https://open.larksuite.com/open-apis"
DEFAULT_WORKER_URL = "https://upg-pallet-stock.yan-lam.workers.dev"
DEFAULT_EXCEL = "DOCK_RECEIPT_CYCLE_COUNT_11.xlsx"
SHEET_NAME = "2026"

# Destination table (table id appears in the Base URL — not a secret).
PALLET_RECORD_TABLE_ID = "tblPGvy6n3bApbiN"

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

DEFAULT_MAX_PAGES = 100
BATCH_CHUNK = 100

# Cloudflare's edge in front of workers.dev blocks the default urllib
# User-Agent ("Python-urllib/x.y") with 403 — send a browser
# User-Agent on every request the script makes.
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")

# The backfill covers 2026-11-01 .. 2026-11-11 only. 11-12 is already
# in Lark; anything later must never be written by this script.
BACKFILL_START = date(2026, 11, 1)
BACKFILL_END = date(2026, 11, 11)

# Expected measured days (verified from the file 2026-11-12). The
# script fails loudly if the file's days differ — a one-off backfill
# must not silently write the wrong history.
EXPECTED_DAYS = [
    date(2026, 11, 1), date(2026, 11, 2), date(2026, 11, 3),
    date(2026, 11, 4), date(2026, 11, 5), date(2026, 11, 6),
    date(2026, 11, 9), date(2026, 11, 10), date(2026, 11, 11),
]
# Expected (deduped rows, departed rows) per measured day, verified
# from the file 2026-11-12. Carried days (11-07, 11-08) are copies of
# 11-06, so they inherit its counts.
EXPECTED_COUNTS = {
    date(2026, 11, 1): (142, 3),
    date(2026, 11, 2): (144, 8),
    date(2026, 11, 3): (140, 1),
    date(2026, 11, 4): (139, 11),
    date(2026, 11, 5): (129, 5),
    date(2026, 11, 6): (127, 6),
    date(2026, 11, 9): (133, 1),
    date(2026, 11, 10): (133, 1),
    date(2026, 11, 11): (133, 4),
}

EXPECTED_HEADERS = [
    "Daily Measurement Date", "Timestamp", "Dock Receipt Date",
    "Business Unit", "Reference (PO# or other Reference#)",
    "Batch# <BU>-<Date Code>-<Packaging Type>"
    "<3 Digit Sequential Number>",
    "Date of Measurement", "Length (cm)", "Width (cm)",
    "Height (cm)", "Volume (CBM)", "Ship Date",
]


def log(msg):
    print(f"[pallet-backfill-nov] {msg}", flush=True)


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
    table's actual field name (the Batch# key carries many spaces)."""
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
    2026-09-30 = 46295, 2026-11-01 = 46327, 2026-11-12 = 46338)."""
    return (d - date(1899, 12, 30)).days


def calc_volume(length_cm, width_cm, height_cm):
    """ROUNDUP(L*W*H/1e6, 1 dp), same as the Worker and app.js
    (epsilon guards against float error pushing an exact value up)."""
    raw = (length_cm * width_cm * height_cm) / 1e6
    return math.ceil(raw * 10 - 1e-9) / 10


def normalize_batch(s):
    """Canonical Batch#: strip ALL whitespace, uppercase, and add a
    missing P (UPG-YYYYMMDD-NNN -> UPG-YYYYMMDD-PNNN). The file has
    'UPG-20220525- 024' (space, no P); the website stores
    'UPG-20220525-P024'. Both must map to the same key."""
    b = re.sub(r"\s+", "", str(s or "")).upper()
    m = re.match(r"^(UPG-\d{8}-)(\d{3})$", b)
    if m:
        b = m.group(1) + "P" + m.group(2)
    return b


def cell_to_date(value):
    """An Excel cell -> datetime.date or None. Handles datetime,
    date, epoch numbers, and date strings; error strings ('#REF!')
    and blanks -> None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return datetime.fromtimestamp(value / 1000,
                                          tz=HKT).date()
        except (ValueError, OverflowError, OSError):
            return None
    s = str(value).strip()
    if not s or s.startswith("#") or s.startswith("="):
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def cell_is_error(value):
    """True for cached formula-error strings ('#REF!', '#VALUE!', ...)
    and unevaluated formula strings."""
    return isinstance(value, str) and (
        value.startswith("#") or value.startswith("="))


def dock_date_for(batch_norm, file_dock):
    """Dock Receipt Date: for UPG- pallets the Batch# date code is
    ground truth; for non-UPG pallets the file's dock date is used
    as-is (blank stays blank)."""
    m = re.match(r"^UPG-(\d{4})(\d{2})(\d{2})-P\d+$", batch_norm)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return cell_to_date(file_dock)


def get_worker_state(worker_url):
    """Full Worker /state. urllib first (3 attempts); the egress
    proxy sometimes truncates urllib responses (IncompleteRead), so
    fall back to curl (2 attempts) before failing loudly."""
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
        for _attempt in range(2):
            try:
                out = subprocess.run(
                    ["curl", "-s", "-A", USER_AGENT,
                     worker_url.rstrip("/") + "/state"],
                    capture_output=True, timeout=60, check=True)
                d = json.loads(out.stdout.decode("utf-8"))
                break
            except (subprocess.SubprocessError, ValueError) as e:
                last_err = e
                log(f"Worker /state curl fallback failed ({e}); "
                    f"retrying...")
                time.sleep(2)
        else:
            raise RuntimeError(
                f"Worker /state read failed: {last_err}")
    log(f"Worker state: {d.get('inStockCount')} in stock, "
        f"{len(d.get('departed') or [])} departed, "
        f"{len(d.get('pending') or [])} pending")
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


# ---------------------------------------------------------------------
# Excel reading
# ---------------------------------------------------------------------

def read_excel(path):
    """Read the '2026' sheet (data_only=True: cached values, so
    11-11's broken formulas arrive as '#REF!' strings). Returns
    {snapshot_date: [row dicts]} with raw cell values, plus the
    header row for validation."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(
            f"Sheet {SHEET_NAME!r} not in {path} "
            f"(sheets: {wb.sheetnames})")
    ws = wb[SHEET_NAME]
    headers = [ws.cell(1, c).value for c in range(1, 14)]
    for i, want in enumerate(EXPECTED_HEADERS):
        got = headers[i]
        # col 13's header is a formula (=A1&F1); with data_only=True
        # it evaluates to the concatenated text — checked loosely.
        if i < len(EXPECTED_HEADERS) and norm_key(got) != norm_key(want):
            raise RuntimeError(
                f"Header mismatch col {i + 1}: got {got!r}, "
                f"want {want!r}")
    mkey_header = str(headers[12] or "")
    if "Daily Measurement Date" not in mkey_header or "Batch#" not in mkey_header:
        raise RuntimeError(
            f"Header mismatch col 13 (M-key): got {headers[12]!r}")
    rows_by_day = {}
    n_rows = 0
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 14)]
        if all(v is None for v in vals):
            continue
        n_rows += 1
        snap = cell_to_date(vals[0])
        if snap is None:
            raise RuntimeError(
                f"Row {r}: Daily Measurement Date unparseable: "
                f"{vals[0]!r}")
        rows_by_day.setdefault(snap, []).append({
            "row": r,
            "location": vals[1],
            "dock_raw": vals[2],
            "business_unit": vals[3],
            "reference": vals[4],
            "batch_raw": vals[5],
            "dom_raw": vals[6],
            "length": vals[7],
            "width": vals[8],
            "height": vals[9],
            "volume": vals[10],
            "ship_raw": vals[11],
        })
    log(f"Excel rows read: {n_rows} data rows over "
        f"{len(rows_by_day)} day(s): "
        f"{sorted(d.isoformat() for d in rows_by_day)}")
    return rows_by_day


def dedupe_day(rows):
    """Dedupe one day's rows by normalized Batch#. DRAFT-PALLET
    appears twice identically — count once. The UPG-20260827 trio
    appears twice (one row's Date of Measurement typo'd 2028-08-27,
    one correct 2026-08-31) — always keep the 2026 row. Returns
    {normalized Batch#: row dict} and the number of dupes dropped."""
    groups = {}
    dropped = 0
    for row in rows:
        batch = normalize_batch(row["batch_raw"])
        if not batch:
            raise RuntimeError(
                f"Row {row['row']}: empty Batch#")
        dom = cell_to_date(row["dom_raw"])
        dom_year = dom.year if dom else None
        if batch not in groups:
            groups[batch] = (row, dom_year)
        else:
            dropped += 1
            old_row, old_year = groups[batch]
            if old_year == 2028 and dom_year == 2026:
                groups[batch] = (row, dom_year)
    return {b: r for b, (r, _y) in groups.items()}, dropped


# ---------------------------------------------------------------------
# Row building
# ---------------------------------------------------------------------

def build_row_fields(batch, snap_date, location, dock, business_unit,
                     reference, length, width, height, ship):
    """One Pallet record field dict for one pallet on one snapshot
    day. Returns (fields, skip_reason or None)."""
    if not isinstance(height, (int, float)) or isinstance(height, bool) \
            or not (0 < height <= 1000):
        return None, f"bad Height {height!r}"
    if not isinstance(length, (int, float)) or isinstance(length, bool) \
            or not (0 < length <= 1000):
        return None, f"bad Length {length!r}"
    if not isinstance(width, (int, float)) or isinstance(width, bool) \
            or not (0 < width <= 1000):
        return None, f"bad Width {width!r}"
    snap_str = snap_date.isoformat()
    fields = {
        F_DMD: date_to_ms(snap_str),
        F_TS: str(location or ""),
        F_BU: str(business_unit or ""),
        F_REF: str(reference or ""),
        F_BATCH: batch,
        F_DOM: date_to_ms(snap_str),
        F_LEN: length,
        F_WID: width,
        F_HEI: height,
        F_VOL: calc_volume(length, width, height),
        F_MKEY: f"{excel_serial(snap_date)}{batch}",
    }
    if dock:
        fields[F_DOCK] = date_to_ms(dock.isoformat())
    if ship:
        # Ship Date is a TEXT field in Pallet record — write the
        # date as a 'YYYY-MM-DD' string, not epoch ms.
        fields[F_SHIP] = ship.isoformat()
    fields = {k: v for k, v in fields.items()
              if v != "" and v is not None}
    return fields, None


def resolve_11_11(rows_11, live_by_batch, rows_10):
    """Resolve location/height for the 11-11 snapshot, whose file
    Timestamp/Height/Volume cells are '#REF!'. Returns
    (resolved, skipped, sources) where resolved maps Batch# ->
    (location, height) and sources maps Batch# -> 'live' /
    'fallback_11_10'. Departed rows resolve from the file's 11-10
    rows only; in-stock rows resolve from live /state first, then
    fall back to the 11-10 row. Unresolvable rows are skipped."""
    resolved, skipped, sources = {}, [], {}
    for batch, row in rows_11.items():
        ship = cell_to_date(row["ship_raw"])
        if ship is not None:
            # Departed on 11-11: height/location from the file's
            # 11-10 row for the same Batch#.
            prev = rows_10.get(batch)
            if prev is None:
                skipped.append((batch, "departed, no 11-10 row"))
                continue
            loc, hei = prev["location"], prev["height"]
            if cell_is_error(loc) or cell_is_error(hei):
                skipped.append((batch, "departed, 11-10 row broken"))
                continue
            resolved[batch] = (loc, hei)
            sources[batch] = "fallback_11_10"
            continue
        # In-stock on 11-11: live /state first.
        live = live_by_batch.get(batch)
        if live is not None:
            hei = live.get("heightCm")
            loc = live.get("location")
            if (isinstance(hei, (int, float)) and not isinstance(hei, bool)
                    and 0 < hei <= 1000
                    and isinstance(loc, str) and loc.strip()):
                resolved[batch] = (loc, hei)
                sources[batch] = "live"
                continue
        # Fallback: the file's 11-10 row (last known-good values).
        prev = rows_10.get(batch)
        if prev is not None:
            loc, hei = prev["location"], prev["height"]
            if (not cell_is_error(loc) and not cell_is_error(hei)
                    and isinstance(hei, (int, float))
                    and not isinstance(hei, bool)
                    and 0 < hei <= 1000
                    and isinstance(loc, str) and loc.strip()):
                resolved[batch] = (loc, hei)
                sources[batch] = "fallback_11_10"
                continue
        skipped.append((batch, "in-stock, no live match and no "
                               "usable 11-10 row"))
    return resolved, skipped, sources


def build_day_fields(snap_date, deduped, resolved_11_11=None):
    """Build the Lark field dicts for one snapshot day. For 11-11,
    resolved_11_11 maps Batch# -> (location, height) and overrides
    the file's broken cells. Returns (fields_list, skipped_list)
    where skipped_list is [(batch, reason)]."""
    out, skipped = [], []
    for batch, row in sorted(deduped.items()):
        if resolved_11_11 is not None:
            if batch not in resolved_11_11:
                # resolve_11_11 already reported this row as skipped.
                continue
            location, height = resolved_11_11[batch]
        else:
            location, height = row["location"], row["height"]
            if cell_is_error(location) or cell_is_error(height):
                skipped.append((batch, "broken location/height cell"))
                continue
        dock = dock_date_for(batch, row["dock_raw"])
        ship = cell_to_date(row["ship_raw"])
        fields, why = build_row_fields(
            batch, snap_date, location, dock,
            row["business_unit"], row["reference"],
            row["length"], row["width"], height, ship)
        if fields is None:
            skipped.append((batch, why))
            continue
        out.append(fields)
    return out, skipped


def carry_fields(src_fields, day):
    """Clone one day's built field dicts for a carried day,
    re-dating Daily Measurement Date / Date of Measurement / M-key.
    Everything else (location, dock date, dimensions, ship date)
    is copied unchanged."""
    cloned = []
    for f in src_fields:
        g = dict(f)
        g[F_DMD] = date_to_ms(day.isoformat())
        g[F_DOM] = date_to_ms(day.isoformat())
        g[F_MKEY] = f"{excel_serial(day)}{f[F_BATCH]}"
        cloned.append(g)
    return cloned


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="List what would be appended, write nothing")
    ap.add_argument("--offline", action="store_true",
                    help="Build rows without contacting Lark (dedupe "
                         "unverified — local testing only)")
    ap.add_argument("--excel-path", default=DEFAULT_EXCEL,
                    help="Cycle-count file (default: "
                         f"{DEFAULT_EXCEL} in the working directory)")
    ap.add_argument("--max-pages", type=int, default=None,
                    help="Max Base pages to scan (default "
                         "SYNC_MAX_PAGES or 100)")
    args = ap.parse_args()

    max_pages = args.max_pages if args.max_pages is not None else int(
        os.environ.get("SYNC_MAX_PAGES", DEFAULT_MAX_PAGES))

    if not args.offline:
        app_id = os.environ.get("LARK_APP_ID")
        app_secret = os.environ.get("LARK_APP_SECRET")
        app_token = os.environ.get("LARK_APP_TOKEN")
        missing = [n for n, v in (
            ("LARK_APP_ID", app_id), ("LARK_APP_SECRET", app_secret),
            ("LARK_APP_TOKEN", app_token)) if not v]
        if missing:
            raise RuntimeError(
                "Missing Lark credentials in environment: "
                + ", ".join(missing) +
                " (pass --offline to build rows without contacting "
                "Lark; dedupe will then be unverified)")
    worker_url = os.environ.get("WORKER_URL", DEFAULT_WORKER_URL)

    # --- Excel: read + verify days + dedupe --------------------------
    rows_by_day = read_excel(args.excel_path)
    found_days = sorted(rows_by_day)
    if found_days != EXPECTED_DAYS:
        raise RuntimeError(
            "Measured days differ from expected: found "
            f"{[d.isoformat() for d in found_days]}, expected "
            f"{[d.isoformat() for d in EXPECTED_DAYS]}")
    if any(d > BACKFILL_END for d in found_days):
        raise RuntimeError(
            "File contains days after the backfill end "
            f"{BACKFILL_END.isoformat()} — refusing to write them")
    deduped = {}
    for day in EXPECTED_DAYS:
        deduped[day], dropped = dedupe_day(rows_by_day[day])
        n_dep = sum(
            1 for row in deduped[day].values()
            if cell_to_date(row["ship_raw"]) is not None)
        want_total, want_dep = EXPECTED_COUNTS[day]
        if (len(deduped[day]), n_dep) != (want_total, want_dep):
            raise RuntimeError(
                f"{day.isoformat()}: deduped "
                f"({len(deduped[day])}, departed {n_dep}) != expected "
                f"({want_total}, departed {want_dep}) — refusing to "
                f"write a history that does not match the verified "
                f"counts")
        log(f"{day.isoformat()}: raw {len(rows_by_day[day])} -> "
            f"deduped {len(deduped[day])} "
            f"({len(deduped[day]) - n_dep} in-stock, "
            f"{n_dep} departed; dropped {dropped} dupes)")

    # --- Carried days: 11-07 and 11-08 copy 11-06 -------------------
    carry_src = date(2026, 11, 6)
    carried_days = []
    for d in [date(2026, 11, 7), date(2026, 11, 8)]:
        if d.weekday() < 5:
            raise RuntimeError(
                f"Carried day {d.isoformat()} is a weekday "
                f"({d.strftime('%A')}) — the carry rule covers "
                f"weekends only; refusing to guess")
        carried_days.append(d)
    log(f"Carried days (copy of {carry_src.isoformat()}): "
        f"{[d.isoformat() for d in carried_days]}")

    # --- 11-11 #REF! resolution (needs live /state) ------------------
    state = get_worker_state(worker_url)
    live_by_batch = {
        normalize_batch(p.get("palletId")): p
        for p in (state.get("inStock") or [])
        if p.get("palletId")
    }
    day_11 = date(2026, 11, 11)
    day_10 = date(2026, 11, 10)
    resolved_11, skipped_11, sources_11 = resolve_11_11(
        deduped[day_11], live_by_batch, deduped[day_10])
    from collections import Counter
    src_counts = Counter(sources_11.values())
    log(f"11-11 resolution: {len(resolved_11)} resolved "
        f"({dict(src_counts)}), {len(skipped_11)} skipped")
    for batch, why in skipped_11:
        log(f"11-11 skipped: {batch} — {why}")

    # --- Build all snapshot rows -------------------------------------
    all_days = EXPECTED_DAYS + carried_days
    built = {}   # date -> [fields]
    skipped_rows = []  # (date, batch, reason)
    for day in EXPECTED_DAYS:
        if day == day_11:
            fields, skipped = build_day_fields(
                day, deduped[day], resolved_11)
            # Rows resolve_11_11 already skipped are reported there;
            # build_day_fields only reports rows that failed field
            # validation after resolution.
            skipped_rows += [(day, b, w) for b, w in skipped]
            skipped_rows += [(day, b, w) for b, w in skipped_11]
        else:
            fields, skipped = build_day_fields(day, deduped[day])
            skipped_rows += [(day, b, w) for b, w in skipped]
        built[day] = fields
        log(f"{day.isoformat()}: built {len(fields)} row(s)")
    # Carried days: clone the 11-06 field dicts, re-dating them.
    src_fields = built[carry_src]
    for day in carried_days:
        built[day] = carry_fields(src_fields, day)
        log(f"{day.isoformat()}: carried {len(built[day])} row(s) "
            f"from {carry_src.isoformat()}")
    total_built = sum(len(v) for v in built.values())
    log(f"Total rows built: {total_built} over {len(all_days)} days")

    # --- Lark dedupe (unless offline) --------------------------------
    existing_pairs = set()
    lark_dedupe_verified = False
    if args.offline:
        log("OFFLINE mode: Lark dedupe NOT verified — "
            "rows_appended below is the would-append count assuming "
            "no rows exist yet in Pallet record.")
    else:
        token = lark_token(app_id, app_secret)
        log("Tenant token ok")
        existing_rows = list_records(
            token, app_token, PALLET_RECORD_TABLE_ID, max_pages,
            "Pallet record")
        for row in existing_rows:
            f = row.get("fields") or {}
            batch = normalize_batch(
                plain_text(field_by_name(f, F_BATCH)).strip())
            dmd = parse_date_ms(field_by_name(f, F_DMD))
            if batch and dmd:
                existing_pairs.add((batch, dmd))
        log(f"Pallet record existing (Batch#, date) pairs: "
            f"{len(existing_pairs)}")
        lark_dedupe_verified = True

    # --- Decide what to append ---------------------------------------
    to_append = []
    per_day = {}
    rows_skipped = 0
    for day in sorted(built):
        day_str = day.isoformat()
        n = 0
        for fields in built[day]:
            pair = (normalize_batch(fields[F_BATCH]), day_str)
            if pair in existing_pairs:
                rows_skipped += 1
                continue
            to_append.append(fields)
            n += 1
        per_day[day_str] = n
        log(f"{day_str}: to append {n} "
            f"(skipped {len(built[day]) - n} already present)")
    log(f"Rows to append: {len(to_append)}; "
        f"rows skipped (already present): {rows_skipped}; "
        f"rows skipped (unresolvable): {len(skipped_rows)}")
    for day, batch, why in skipped_rows:
        log(f"Skipped (unresolvable): {day.isoformat()} {batch} — {why}")
    if to_append:
        sample = to_append[0]
        log(f"Sample row to append: {json.dumps(sample, ensure_ascii=False)}")
        dep_samples = [f for f in to_append if F_SHIP in f]
        if dep_samples:
            log(f"Sample departed row: "
                f"{json.dumps(dep_samples[0], ensure_ascii=False)}")

    # --- Write (unless dry run) --------------------------------------
    appended = 0
    if args.dry_run:
        log("Dry run: nothing written to Lark.")
        appended = len(to_append)
    elif args.offline:
        raise RuntimeError(
            "--offline without --dry-run would write to Lark with "
            "unverified dedupe — refusing. Run with --dry-run, or "
            "without --offline once Lark credentials are set.")
    else:
        if to_append:
            appended = batch_create(
                token, app_token, PALLET_RECORD_TABLE_ID, to_append,
                "Pallet record")

    print(json.dumps({
        "dry_run": args.dry_run,
        "offline": args.offline,
        "lark_dedupe_verified": lark_dedupe_verified,
        "days": [d.isoformat() for d in sorted(built)],
        "rows_appended": appended,
        "rows_skipped": rows_skipped,
        "rows_skipped_unresolvable": len(skipped_rows),
        "skipped_unresolvable": [
            {"date": d.isoformat(), "batch": b, "reason": w}
            for d, b, w in skipped_rows
        ],
        "per_day": per_day,
    }, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
