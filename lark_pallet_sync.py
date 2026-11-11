#!/usr/bin/env python3
"""UPG Pallet — Lark Dock Receipt -> Cloudflare Worker pending queue.

Two-step sync (Yan, 2026-11-11 20:17):
  1. A new row lands in Lark Base "UPGL inbound and cycle count",
     table Dock Receipt (via the "UPGL" form). A Lark automation fires
     repository_dispatch (event type `lark-new-pallet`) to this repo.
  2. This workflow script reads recent Dock Receipt rows via the Lark
     API, keeps only Business Unit = Labware AND Batch# P-type (pallet),
     and POSTs them to the Cloudflare Worker /syncBatch endpoint — which
     puts them in the PENDING queue, NOT in stock.
  3. A colleague confirms each pending pallet in the website
     ("確認加入") or rejects it ("唔加入"); only then does it join the
     in-stock table. Confirm/reject are logged in the activity record.

Rules (same as the Muse site):
  - Batch# is normalized (strip spaces, uppercase) for dedupe; the
    Worker's state (in-stock + pending) is the source of truth —
    already-present pallets are skipped, never duplicated.
  - Volume is recomputed by the Worker from L x W x H (ROUNDUP to
    1 dp); the Base's Volume column is never copied (it has wrong
    values like 72 / 86).
  - Only rows with Dock Receipt Date within +/- DAYS_BACK days of
    today are considered (default 14) — pre-entered future rows count.
  - Type must be "Dock Receipt" or blank (cycle-count / transfer rows
    are out).
  - --dry-run lists what would be synced and writes nothing.

Environment (all from GitHub repo Secrets in the workflow, except
WORKER_URL which is public):
  LARK_APP_ID, LARK_APP_SECRET, LARK_APP_TOKEN, LARK_TABLE_ID,
  WORKER_URL, WORKER_SYNC_SECRET (optional until Yan sets the matching
  Worker secret SYNC_SECRET), SYNC_DAYS_BACK (optional, default 14),
  SYNC_MAX_PAGES (optional, default 50).
"""

import argparse
import http.client
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

HKT = timezone(timedelta(hours=8))
LARK_API = "https://open.larksuite.com/open-apis"

# Dock Receipt field names (must match the Base table exactly).
F_BU = "Business Unit"
F_BATCH = "Batch#"
F_LENGTH = "Length (cm)"
F_WIDTH = "Width (cm)"
F_HEIGHT = "Height (cm)"
F_DOCK_DATE = "Dock Receipt Date"
F_LOCATION = "Location"
F_REFERENCE = "Reference"
F_TYPE = "Type"

DEFAULT_DAYS_BACK = 14
DEFAULT_MAX_PAGES = 50

# Cloudflare's edge in front of workers.dev blocks the default urllib
# User-Agent ("Python-urllib/x.y") with 403 — the Worker itself only
# checks Origin, and urllib sends none (which passes). Proven
# 2026-11-12: curl with Python-urllib UA -> 403, browser UA -> 200.
# Send a browser User-Agent on every request the script makes.
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/120.0.0.0 Safari/537.36")


def log(msg):
    print(f"[pallet-sync] {msg}", flush=True)


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


def normalize_batch(s):
    """Batch# dedupe key: strip all whitespace, uppercase."""
    return re.sub(r"\s+", "", str(s or "")).upper()


def is_pallet_batch(norm):
    """P-type pallet: normalized Batch# ends with -P + digits
    (e.g. UPG-20261109-P001). Cartons (-C...) and odd legacy codes
    (UPG-20220525-024, ...-P001(DRAFT)) do not match."""
    return re.search(r"-P\d+$", norm) is not None


def read_rows(token, app_token, table_id, max_pages):
    """All Dock Receipt rows, paginated. Filtering happens client-side
    (same lesson as the reimbursement sync: do not rely on the search
    API's field-operator support)."""
    rows, page_token, pages = [], None, 0
    seen_ids = set()
    while True:
        body = {"page_size": 100}
        if page_token:
            body["page_token"] = page_token
        r = lark_call("POST",
                      f"/bitable/v1/apps/{app_token}/tables/{table_id}"
                      f"/records/search", body, token)
        if r.get("code") != 0:
            raise RuntimeError(f"Dock Receipt records search failed: {r}")
        data = r.get("data") or {}
        items = data.get("items") or []
        dupes = 0
        for item in items:
            rid = item.get("record_id")
            if rid:
                if rid in seen_ids:
                    dupes += 1
                    continue
                seen_ids.add(rid)
            rows.append(item)
        if dupes:
            log(f"Page {pages + 1}: dropped {dupes} duplicate row(s)")
        pages += 1
        if not data.get("has_more") or pages >= max_pages:
            break
        page_token = data.get("page_token")
        time.sleep(0.5)
    log(f"Dock Receipt rows read: {len(rows)} distinct "
        f"({pages} page(s))")
    if pages >= max_pages:
        log(f"WARNING: hit MAX_PAGES={max_pages} — older rows were not "
            f"scanned. Raise SYNC_MAX_PAGES if the table keeps growing.")
    return rows


def get_worker_keys(worker_url):
    """Normalized Batch# keys already in stock or pending (the Worker's
    state is the dedupe source of truth). Retries transient read
    failures (truncated responses seen through proxies)."""
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
    keys = set()
    for p in (d.get("inStock") or []) + (d.get("pending") or []):
        k = p.get("key") or normalize_batch(p.get("palletId"))
        if k:
            keys.add(k)
    log(f"Worker state: {d.get('inStockCount')} in stock, "
        f"{len(d.get('pending') or [])} pending")
    return keys


def select_candidates(rows, worker_keys, today, days_back):
    """Keep only new Labware P-type rows with a usable height and a
    Dock Receipt Date within +/- days_back of today.
    Returns (candidates, skipped) where skipped is [(batch, reason)]."""
    candidates, skipped = [], []
    for row in rows:
        f = row.get("fields") or {}
        rid = row.get("record_id")
        batch_raw = plain_text(f.get(F_BATCH)).strip()
        if not batch_raw:
            skipped.append(("", f"row {rid}: empty Batch#"))
            continue
        norm = normalize_batch(batch_raw)
        if plain_text(f.get(F_BU)).strip() != "Labware":
            skipped.append((batch_raw, "BU is not Labware"))
            continue
        if not is_pallet_batch(norm):
            skipped.append((batch_raw, "not P-type Batch#"))
            continue
        typ = plain_text(f.get(F_TYPE)).strip()
        if typ and typ != "Dock Receipt":
            skipped.append((batch_raw, f"Type={typ!r} (not Dock Receipt)"))
            continue
        dock = parse_date_ms(f.get(F_DOCK_DATE))
        if not dock:
            skipped.append((batch_raw, "no Dock Receipt Date"))
            continue
        try:
            dock_d = datetime.strptime(dock, "%Y-%m-%d").date()
        except ValueError:
            skipped.append((batch_raw, f"bad Dock Receipt Date {dock!r}"))
            continue
        if abs((dock_d - today).days) > days_back:
            skipped.append((batch_raw,
                            f"Dock Receipt Date {dock} outside "
                            f"+/-{days_back}d window"))
            continue
        height = f.get(F_HEIGHT)
        if not isinstance(height, (int, float)) or not (0 < height <= 1000):
            skipped.append((batch_raw, f"bad Height {height!r}"))
            continue
        if norm in worker_keys:
            skipped.append((batch_raw, "already in stock / pending"))
            continue
        length = f.get(F_LENGTH)
        width = f.get(F_WIDTH)
        ref = plain_text(f.get(F_REFERENCE)).strip()
        candidates.append({
            "palletId": batch_raw,
            "heightCm": height,
            "date": dock,
            "dockReceiptDate": dock,
            "location": plain_text(f.get(F_LOCATION)).strip(),
            "businessUnit": "Labware",
            "reference": "" if ref in ("", "N/A") else ref,
            "lengthCm": length if isinstance(length, (int, float))
                        and 0 < length <= 1000 else 120,
            "widthCm": width if isinstance(width, (int, float))
                       and 0 < width <= 1000 else 100,
        })
    return candidates, skipped


def post_sync_batch(worker_url, secret, pallets):
    """One atomic POST: the Worker adds all non-duplicates to pending
    in a single read-modify-write (avoids the KV cache race that
    per-pallet POSTs would hit)."""
    body = {"action": "syncBatch", "pallets": pallets}
    if secret:
        body["secret"] = secret
    r = http_json("POST", worker_url.rstrip("/") + "/", body,
                  {"Content-Type": "application/json"})
    if not isinstance(r, dict) or not r.get("ok"):
        raise RuntimeError(f"Worker syncBatch failed: {r}")
    return r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="List what would sync, write nothing")
    ap.add_argument("--days-back", type=int, default=None,
                    help="Dock Receipt Date window +/- days (default "
                         "SYNC_DAYS_BACK or 14)")
    ap.add_argument("--max-pages", type=int, default=None,
                    help="Max Base pages to scan (default SYNC_MAX_PAGES "
                         "or 50)")
    args = ap.parse_args()

    days_back = args.days_back if args.days_back is not None else int(
        os.environ.get("SYNC_DAYS_BACK", DEFAULT_DAYS_BACK))
    max_pages = args.max_pages if args.max_pages is not None else int(
        os.environ.get("SYNC_MAX_PAGES", DEFAULT_MAX_PAGES))

    app_id = os.environ["LARK_APP_ID"]
    app_secret = os.environ["LARK_APP_SECRET"]
    app_token = os.environ["LARK_APP_TOKEN"]
    table_id = os.environ["LARK_TABLE_ID"]
    worker_url = os.environ["WORKER_URL"]
    worker_secret = os.environ.get("WORKER_SYNC_SECRET") or ""

    token = lark_token(app_id, app_secret)
    log("Tenant token ok")

    rows = read_rows(token, app_token, table_id, max_pages)
    worker_keys = get_worker_keys(worker_url)
    today = datetime.now(HKT).date()
    candidates, skipped = select_candidates(rows, worker_keys, today,
                                            days_back)
    for batch, why in skipped:
        log(f"Skipped: {batch or '(no Batch#)'} — {why}")
    log(f"Candidates to sync: {len(candidates)} "
        f"({[c['palletId'] for c in candidates]})")

    if not candidates:
        print(json.dumps({"dry_run": args.dry_run, "fetched": len(rows),
                          "added": [], "skipped": len(skipped)}, indent=2))
        return 0

    if args.dry_run:
        log("Dry run: nothing written to the Worker.")
        print(json.dumps({"dry_run": True, "fetched": len(rows),
                          "would_add": [c["palletId"] for c in candidates],
                          "skipped": len(skipped)}, indent=2))
        return 0

    result = post_sync_batch(worker_url, worker_secret, candidates)
    log(f"Worker syncBatch: added={result.get('added')}, "
        f"skipped={result.get('skipped')}")
    print(json.dumps({"dry_run": False, "fetched": len(rows),
                      "added": result.get("added"),
                      "skipped": result.get("skipped")}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
