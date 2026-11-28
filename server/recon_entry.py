#!/usr/bin/env python3
"""Recon data-entry server: lets the GitHub Pages form read and write the
DataInput workbooks (in a Team library or Becca's OneDrive) through Microsoft Graph.

Runs on the lab Mac mini under Becca's account and uses her delegated
"E5-Nap API" token (the MSAL cache from ~/Code/admin/general/api-access).
The token never leaves this machine: the form talks to this server, and this
server only reads and updates building rows in the workbooks named in the
config, and adds Street View pictures and DebrisTech 360 stills (with the citation
burned in) that students save into a building's photo folder. It never deletes anything.

Config (outside the repo): ~/.config/recon-entry/config.json
  {
    "lab_code": "...",                      # students type this once in the form
    "allowed_origins": ["https://rkn2.github.io"],
    "datasets": {
      "stlouis": {"drive": "<driveId of a Team library, optional>",
                  "path": "tornadoREUs/.../StLouis_Tornado_DataInput.xlsx",
                  "sheet": "BuidlingAttributes", "header_row": 2},
      ...
    }
  }

Run:  uv run --with requests --with msal python server/recon_entry.py [--port 8787]

Street View saves need "google_key" in the config (a Street View Static API
key, never sent to the browser) and are capped per day by
"streetview_daily_cap" (default 200), counted in streetview_calls.jsonl.
A dataset's "storm_month" (e.g. "2025-05") labels each saved picture as
before or after the storm.

Every write is appended to ~/.local/state/recon-entry/writes.jsonl with the
old and new value of each cell, so any save can be traced and undone.
"""
import argparse
import hashlib
import hmac
import json
import logging
import math
import re
import sys
import threading
import time
import traceback
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import msal
import requests

sys.path.insert(0, "/Users/becca/Code/admin/general/api-access")
import test_graph_login as tgl  # noqa: E402

CONFIG_PATH = Path.home() / ".config/recon-entry/config.json"
STATE_DIR = Path.home() / ".local/state/recon-entry"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Files.ReadWrite.All"]
CACHE_SECONDS = 30
MAX_BODY = 512 * 1024

log = logging.getLogger("recon-entry")


# ── Graph ───────────────────────────────────────────────────────────

_token_lock = threading.Lock()
_token = {"value": None, "expires": 0.0}


def _save_cache_atomic(cache):
    """Like tgl.save_cache, but write-then-rename so Becca's other jobs sharing this
    cache file never read a half-written file."""
    if cache.has_state_changed:
        tmp = tgl.CACHE_PATH.with_suffix(".tmp-recon-entry")
        tmp.write_text(cache.serialize())
        tmp.chmod(0o600)
        tmp.replace(tgl.CACHE_PATH)


def token():
    """Silent token from Becca's MSAL cache, kept in memory until 5 min before it
    expires. Never falls back to device code."""
    with _token_lock:
        if _token["value"] and time.time() < _token["expires"] - 300:
            return _token["value"]
        cache = tgl.load_cache()
        app = msal.PublicClientApplication(
            tgl.CLIENT_ID, authority=f"https://login.microsoftonline.com/{tgl.TENANT_ID}", token_cache=cache)
        accounts = app.get_accounts()
        if not accounts:
            raise RuntimeError("No signed-in account in the MSAL cache; run test_graph_login.py interactively.")
        r = app.acquire_token_silent_with_error(SCOPES, account=accounts[0])
        _save_cache_atomic(cache)
        if not r or "access_token" not in r:
            raise RuntimeError(f"Silent token failed: {r}")
        _token["value"] = r["access_token"]
        _token["expires"] = time.time() + float(r.get("expires_in", 600))
        return _token["value"]


class GraphError(RuntimeError):
    def __init__(self, msg, status, code):
        super().__init__(msg)
        self.status, self.code = status, code


def graph(method, url, session=None, headers_extra=None, **kw):
    headers = {"Authorization": "Bearer " + token(), **(headers_extra or {})}
    if session:
        headers["workbook-session-id"] = session
    r = requests.request(method, GRAPH + url, headers=headers, timeout=60, **kw)
    if r.status_code >= 400:
        try:
            code = r.json()["error"]["code"]
        except Exception:
            code = ""
        raise GraphError(f"Graph {method} {url} -> {r.status_code}: {r.text[:2000]}", r.status_code, code)
    return r.json() if r.content else {}


def col_letter(n):  # 1-based
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


# ── Workbook model ──────────────────────────────────────────────────

def row_hash(values):
    return hashlib.sha256(json.dumps(values, separators=(",", ":"), default=str).encode()).hexdigest()[:16]


class Dataset:
    def __init__(self, key, cfg):
        self.key = key
        self.path = cfg["path"]
        self.sheet = cfg.get("sheet", "BuidlingAttributes")
        self.header_row = int(cfg.get("header_row", 1))
        self.photos = cfg.get("photos")  # OneDrive folder of per-building photo folders (optional)
        # More read-only roots laid out the same way (Marshall keeps "after" apart from "before").
        # The Building Photos tab merges them; saves still go to `photos`.
        self.photos_more = list(cfg.get("photos_more", []))
        # Optional driveId of a Team's document library; paths are then relative to that library's root.
        self.drive = f"/drives/{cfg['drive']}" if cfg.get("drive") else "/me/drive"
        self._photo_folders = None       # (time, [names])
        # Month of the storm ("2025-05" for St. Louis); Street View saves are labelled before/after it.
        self.storm = cfg.get("storm_month")
        # Base URL of a DebrisTech 360 viewer's tile index (".../player2/tiles-NNN"), optional.
        self.dt360_tiles = cfg.get("dt360_tiles")
        self.dt360_label = cfg.get("dt360_label", "")  # event name for the citation on saved 360 stills
        self.lock = threading.Lock()  # serializes writes to this workbook
        self._item = None
        self._session = None  # persistent workbook session: reads see earlier writes
        self._session_lock = threading.Lock()
        self._cache = None  # (time, headers, rows)

    def item_url(self):
        if self._item is None:
            it = graph("GET", f"{self.drive}/root:/{quote(self.path)}?$select=id,parentReference")
            self._item = f"/drives/{it['parentReference']['driveId']}/items/{it['id']}"
        return self._item

    def ws_url(self):
        return f"{self.item_url()}/workbook/worksheets/{quote(self.sheet)}"

    def wb(self, method, url, **kw):
        """Workbook call inside this dataset's session; recreates an expired session once.
        Sessionless calls can read stale data right after a write (seen 2026-10-16)."""
        for attempt in (1, 2):
            with self._session_lock:
                if self._session is None:
                    self._session = graph("POST", f"{self.item_url()}/workbook/createSession",
                                          json={"persistChanges": True})["id"]
                    log.info("%s: new workbook session", self.key)
            try:
                return graph(method, url, session=self._session, **kw)
            except GraphError as ex:
                expired = ex.code in ("InvalidSessionReCreatable", "InvalidSession", "SessionNotFound") or (
                    ex.status == 404 and "session" in str(ex).lower())
                if attempt == 1 and expired:
                    log.warning("%s: session expired (%s), recreating", self.key, ex.code)
                    self._session = None
                    continue
                raise

    def load(self, fresh=False):
        """Returns (headers, rows) where rows is a list of (row_number, values)."""
        if not fresh and self._cache and time.time() - self._cache[0] < CACHE_SECONDS:
            return self._cache[1], self._cache[2]
        used = self.wb("GET", f"{self.ws_url()}/usedRange(valuesOnly=true)?$select=address,values")
        values = used["values"]
        # usedRange may not start at A1
        m = re.search(r"!([A-Z]+)(\d+)", used["address"])
        if m.group(1) != "A":
            raise RuntimeError(f"{self.key}: used range starts at column {m.group(1)}, expected A")
        first_row = int(m.group(2))
        hi = self.header_row - first_row
        headers = [str(h).strip() for h in values[hi]]
        rows = []
        for i, vals in enumerate(values[hi + 1:], start=self.header_row + 1):
            rows.append((i, vals))
        self._cache = (time.time(), headers, rows)
        return headers, rows

    def invalidate(self):
        self._cache = None

    @staticmethod
    def first_index(headers):
        """Column name -> first column index (duplicate names map to the first one)."""
        idx = {}
        for i, h in enumerate(headers):
            if h and h not in idx:
                idx[h] = i
        return idx

    def schema(self):
        headers, _ = self.load()
        dups = sorted({h for h in headers if h and headers.count(h) > 1})
        return {"columns": [h for h in headers if h], "duplicate_columns": dups}

    def row_dict(self, headers, rownum, vals):
        d = {"_row": rownum, "_hash": row_hash(vals)}
        for h, i in self.first_index(headers).items():
            v = vals[i] if i < len(vals) else ""
            if v != "":
                d[h] = v
        return d

    def all_rows(self):
        headers, rows = self.load()
        out = []
        for n, vals in rows:
            if any(str(v).strip() for v in vals):
                out.append(self.row_dict(headers, n, vals))
        return out

    def read_row(self, headers, rownum):
        last = col_letter(len(headers))
        rng = self.wb("GET", f"{self.ws_url()}/range(address='A{rownum}:{last}{rownum}')?$select=values")
        vals = rng["values"][0]
        return vals + [""] * (len(headers) - len(vals))

    def write_cells(self, rownum, old, new):
        """PATCH only the cells that changed, grouped into contiguous runs."""
        changed = [i for i, (a, b) in enumerate(zip(old, new)) if a != b and str(a) != str(b)]
        runs = []
        for i in changed:
            if runs and runs[-1][-1] == i - 1:
                runs[-1].append(i)
            else:
                runs.append([i])
        for run in runs:
            a, b = col_letter(run[0] + 1), col_letter(run[-1] + 1)
            self.wb("PATCH", f"{self.ws_url()}/range(address='{a}{rownum}:{b}{rownum}')",
                    json={"values": [[new[i] for i in run]]})

    def update(self, rownum, expected_hash, fields, who):
        with self.lock:
            headers, _ = self.load(fresh=True)
            if rownum <= self.header_row:
                raise BadRequest("row is in the header")
            current = self.read_row(headers, rownum)
            if row_hash(current) != expected_hash:
                self.invalidate()
                return 409, {"error": "conflict", "message": "This building was changed by someone else since you opened it.",
                             "current": self.row_dict(headers, rownum, current)}
            new = self._apply(headers, list(current), fields)
            changes = self._diff(headers, current, new)
            if changes:
                self.write_cells(rownum, current, new)
                current = self.read_row(headers, rownum)  # what Excel actually stored
                audit({"dataset": self.key, "op": "update", "row": rownum, "who": who, "changes": changes})
            self.invalidate()
            return 200, {"row": self.row_dict(headers, rownum, current), "changed": len(changes)}

    def append(self, fields, who):
        with self.lock:
            headers, rows = self.load(fresh=True)
            last_used = max([n for n, v in rows if any(str(x).strip() for x in v)] or [self.header_row])
            rownum = last_used + 1
            current = self.read_row(headers, rownum)
            if any(str(v).strip() for v in current):
                raise RuntimeError(f"{self.key}: row {rownum} expected empty but has data")
            new = self._apply(headers, list(current), fields)
            changes = self._diff(headers, current, new)
            self.write_cells(rownum, current, new)
            stored = self.read_row(headers, rownum)
            audit({"dataset": self.key, "op": "append", "row": rownum, "who": who, "changes": changes})
            self.invalidate()
            return 201, {"row": self.row_dict(headers, rownum, stored)}

    def _apply(self, headers, vals, fields):
        idx = self.first_index(headers)
        unknown = [k for k in fields if k not in idx]
        if unknown:
            raise BadRequest(f"unknown columns: {unknown}")
        locked = [k for k in fields if k.startswith("nws_")]
        if locked:
            raise BadRequest(f"NWS survey columns are filled by nws_join.py and can't be edited: {locked}")
        vals += [""] * (len(headers) - len(vals))
        for k, v in fields.items():
            if isinstance(v, bool):
                v = "Y" if v else ""
            elif v is None:
                v = ""
            elif not isinstance(v, (int, float, str)):
                raise BadRequest(f"bad value for {k}")
            if isinstance(v, str) and v and v[0] in "=+-@" and not re.fullmatch(r"[-+]?\d+(\.\d+)?", v):
                v = "'" + v  # keep text as text, never a formula
            vals[idx[k]] = v
        return vals

    @staticmethod
    def _diff(headers, old, new):
        out = []
        for i, (a, b) in enumerate(zip(old, new)):
            if a != b and str(a) != str(b):
                out.append({"col": headers[i], "old": a, "new": b})
        return out


def photo_folders(ds):
    """Names of the per-building photo folders, cached 10 minutes."""
    if not ds.photos:
        raise BadRequest("no photo folder for this dataset")
    if ds._photo_folders and time.time() - ds._photo_folders[0] < 600:
        return ds._photo_folders[1]
    roots = []
    for path in [ds.photos] + ds.photos_more:
        root = graph("GET", f"{ds.drive}/root:/{quote(path)}?$select=id,webUrl")
        names, url = [], f"{ds.drive}/items/{root['id']}/children?$select=name,folder&$top=999"
        while url:
            page = graph("GET", url)
            names += [c["name"] for c in page["value"] if "folder" in c]
            nxt = page.get("@odata.nextLink")
            url = nxt[len(GRAPH):] if nxt else None
        roots.append({"path": path, "webUrl": root["webUrl"], "folders": names})
    merged = list(dict.fromkeys(n for r in roots for n in r["folders"]))
    ds._photo_folders = (time.time(), {"folders": merged, "webUrl": roots[0]["webUrl"], "_roots": roots})
    return ds._photo_folders[1]


def photos_in(ds, folder):
    """Photos in one building folder, with short-lived thumbnail links."""
    pf = photo_folders(ds)
    if folder not in pf["folders"]:
        raise BadRequest("unknown photo folder")
    out, folder_url = [], None
    for r in pf["_roots"]:
        if folder not in r["folders"]:
            continue
        it = graph("GET", f"{ds.drive}/root:/{quote(r['path'] + '/' + folder)}:/children"
                          "?$select=name,webUrl,file&$expand=thumbnails&$top=200")
        for f in it["value"]:
            if "file" not in f:
                continue
            th = (f.get("thumbnails") or [{}])[0]
            when = r["path"].rstrip("/").rsplit("/", 1)[-1].lower()  # root named before/after tags its photos
            out.append({"file": f["name"], "web": f["webUrl"], "when": when if when in ("before", "after") else None,
                        "thumb": (th.get("medium") or {}).get("url"), "large": (th.get("large") or {}).get("url")})
        folder_url = folder_url or r["webUrl"] + "/" + quote(folder)
    return {"folder": folder, "webUrl": folder_url, "photos": out}


# ── Street View stills ("Save this view") ──────────────────────────
SV_META = "https://maps.googleapis.com/maps/api/streetview/metadata"
SV_IMG = "https://maps.googleapis.com/maps/api/streetview"
_sv_lock = threading.Lock()


def _google_get(url, params, timeout):
    """GET a Google endpoint; errors never carry the URL (it holds the key)."""
    try:
        return requests.get(url, params=params, timeout=timeout)
    except requests.RequestException as ex:
        raise RuntimeError(f"Google Street View request failed ({type(ex).__name__})") from None


def _sv_count_today():
    f = STATE_DIR / "streetview_calls.jsonl"
    today = datetime.now(timezone.utc).date().isoformat()
    if not f.exists():
        return 0
    return sum(1 for line in f.read_text().splitlines() if line.startswith('{"day": "' + today))


def _sv_record(rec):
    rec = {"day": datetime.now(timezone.utc).date().isoformat(), **rec}
    with open(STATE_DIR / "streetview_calls.jsonl", "a") as f:
        f.write(json.dumps(rec) + "\n")


def safe_folder_name(addr):
    name = re.sub(r'[\\/:*?"<>|#%]', " ", str(addr)).strip().rstrip(".")
    return re.sub(r"\s+", " ", name)[:200]


def save_streetview(ds, b, who):
    """Fetch one Street View still by pano id and file it in the building's photo folder."""
    key = CFG.get("google_key", "").strip()
    if not key or not ds.photos:
        raise BadRequest("Street View saving is not set up for this dataset")
    pano = str(b.get("pano", ""))
    if not re.fullmatch(r"[A-Za-z0-9_-]{10,64}", pano):
        raise BadRequest("bad pano id")
    try:
        heading = float(b.get("heading", 0)) % 360
        pitch = max(-90.0, min(90.0, float(b.get("pitch", 0))))
        fov = max(10.0, min(120.0, float(b.get("fov", 90))))
        row = int(b.get("row"))
    except (TypeError, ValueError):
        raise BadRequest("bad view numbers")
    rec = next((r for r in ds.all_rows() if r.get("_row") == row), None)
    if rec is None:
        raise BadRequest("unknown row")
    folders = photo_folders(ds)["folders"]
    folder = b.get("folder") or ""
    if folder and folder not in folders:
        raise BadRequest("unknown photo folder")
    if not folder:
        folder = safe_folder_name(rec.get("complete_address") or rec.get("complete address") or f"row {row}")
        if not folder:
            raise BadRequest("this building has no address for a photo folder")
    with _sv_lock:
        cap = int(CFG.get("streetview_daily_cap", 200))
        if _sv_count_today() >= cap:
            return 429, {"error": "daily_cap", "message": f"Street View saves are capped at {cap} a day. Try again tomorrow."}
        meta = _google_get(SV_META, {"pano": pano, "key": key}, 30).json()
        if meta.get("status") != "OK":
            _sv_record({"kind": "meta", "pano": pano, "status": meta.get("status")})
            raise BadRequest(f"Google has no Street View image for this view ({meta.get('status')})")
        img = _google_get(SV_IMG, {"size": "640x640", "pano": pano, "heading": f"{heading:.1f}",
                                     "pitch": f"{pitch:.1f}", "fov": f"{fov:.0f}", "key": key,
                                     "return_error_code": "true"}, 60)
        _sv_record({"kind": "image", "pano": pano, "status": img.status_code, "who": who, "row": row})
    if img.status_code != 200 or not img.headers.get("Content-Type", "").startswith("image/"):
        raise GraphError(f"Street View image request failed: {img.status_code}", img.status_code, "")
    date = str(meta.get("date") or "undated")
    if not ds.storm or date == "undated":
        when = ""
    else:
        when = "before" if date < ds.storm else "after" if date > ds.storm else "stormmonth"
    fname = f"GSV_{when + '_' if when else ''}{date}_h{heading:03.0f}.jpg"
    path = quote(f"{ds.photos}/{folder}/{fname}")
    it = graph("PUT", f"{ds.drive}/root:/{path}:/content?@microsoft.graph.conflictBehavior=rename",
               data=img.content, headers_extra={"Content-Type": "image/jpeg"})
    ds._photo_folders = None  # a new folder may exist now
    audit({"dataset": ds.key, "op": "streetview", "row": row, "who": who, "folder": folder,
           "file": it.get("name"), "pano": pano, "date": date, "heading": round(heading, 1),
           "pitch": round(pitch, 1), "fov": round(fov), "at": datetime.now(timezone.utc).isoformat()})
    return 200, {"file": it.get("name"), "webUrl": it.get("webUrl"), "folder": folder, "date": date, "when": when}


# ── DebrisTech 360 drive (after the storm) ──────────────────────────
# The viewer's frame index lives in zoom-15 GeoJSON tiles; frames carry id/lat/lon/heading/date
# in their properties. The browser can't read them (no CORS), so the server finds frames.
# Each frame's panorama is an equirectangular tile pyramid at <tour>/panorama-tiles/<id>/L_col_row.jpg
# (1-based; level L is 2^L x 2^(L-1) tiles of 1024 px). Pano column at the centre is the camera
# heading, so a compass bearing b is the viewer's hlookat b - heading (checked against the viewer
# 2026-11-25). The lab may use the imagery with a citation, so every saved still has it burned in.
_dt_tiles = {}  # (x, y) -> list of frame dicts
DT_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{8,100}")


def _dt_tile(base, x, y):
    if (x, y) not in _dt_tiles:
        r = requests.get(f"{base}/15/{x}/{y}.geojson", timeout=30)
        pts = []
        if r.status_code == 200:
            for f in r.json().get("features", []):
                pr = f.get("properties") or {}
                if "lat" in pr and "lon" in pr:
                    pts.append({"id": str(pr.get("id", "")), "lat": float(pr["lat"]), "lon": float(pr["lon"]),
                                "heading": float(pr.get("heading") or 0), "date": str(pr.get("date", ""))[:10]})
        elif r.status_code != 404:
            raise RuntimeError(f"360 tile {x}/{y}: HTTP {r.status_code}")
        _dt_tiles[(x, y)] = pts
    return _dt_tiles[(x, y)]


def _dist_m(lat1, lon1, lat2, lon2):
    return 6371000 * math.hypot(math.radians(lat2 - lat1), math.radians(lon2 - lon1) * math.cos(math.radians(lat1)))


def _bearing(lat1, lon1, lat2, lon2):
    p1, p2, dl = math.radians(lat1), math.radians(lat2), math.radians(lon2 - lon1)
    y = math.sin(dl) * math.cos(p2)
    x = math.cos(p1) * math.sin(p2) - math.sin(p1) * math.cos(p2) * math.cos(dl)
    return math.degrees(math.atan2(y, x)) % 360


def _dt_frames_near(ds, lat, lon):
    base = ds.dt360_tiles
    if not base:
        raise BadRequest("no 360 drive for this dataset")
    if not (-90 < lat < 90 and -180 < lon < 180):
        raise BadRequest("bad coordinates")
    n = 2 ** 15
    x = int((lon + 180) / 360 * n)
    y = int((1 - math.asinh(math.tan(math.radians(lat))) / math.pi) / 2 * n)
    out = []
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            out.extend(_dt_tile(base, x + dx, y + dy))
    return out


def dt360_nearest(ds, lat, lon):
    best = min(_dt_frames_near(ds, lat, lon), key=lambda f: _dist_m(lat, lon, f["lat"], f["lon"]), default=None)
    if best is None:
        return {"distance_m": None, "date": None}
    return {"distance_m": round(_dist_m(lat, lon, best["lat"], best["lon"])), "date": best["date"]}


def _msgpack(buf):
    """Just enough msgpack to read the viewer's share links (arrays, maps, strings, numbers, nil/bool)."""
    import struct
    pos = 0

    def take(n):
        nonlocal pos
        if pos + n > len(buf):
            raise ValueError("short")
        v = buf[pos:pos + n]
        pos += n
        return v

    def one(depth=0):
        if depth > 8:
            raise ValueError("too deep")
        t = take(1)[0]
        if t <= 0x7f: return t
        if t >= 0xe0: return t - 256
        if 0x90 <= t <= 0x9f: return [one(depth + 1) for _ in range(t & 0x0f)]
        if 0x80 <= t <= 0x8f: return {one(depth + 1): one(depth + 1) for _ in range(t & 0x0f)}
        if 0xa0 <= t <= 0xbf: return take(t & 0x1f).decode("utf-8", "replace")
        simple = {0xc0: None, 0xc2: False, 0xc3: True}
        if t in simple: return simple[t]
        fmts = {0xca: ">f", 0xcb: ">d", 0xcc: ">B", 0xcd: ">H", 0xce: ">I", 0xcf: ">Q",
                0xd0: ">b", 0xd1: ">h", 0xd2: ">i", 0xd3: ">q"}
        if t in fmts:
            f = fmts[t]
            return struct.unpack(f, take(struct.calcsize(f)))[0]
        if t in (0xd9, 0xda, 0xdb):
            n = struct.unpack({0xd9: ">B", 0xda: ">H", 0xdb: ">I"}[t], take({0xd9: 1, 0xda: 2, 0xdb: 4}[t]))[0]
            return take(n).decode("utf-8", "replace")
        if t in (0xdc, 0xdd):
            n = struct.unpack(">H" if t == 0xdc else ">I", take(2 if t == 0xdc else 4))[0]
            if n > 64:
                raise ValueError("array too long")
            return [one(depth + 1) for _ in range(n)]
        raise ValueError(f"msgpack type {t:#x}")
    return one()


def dt360_parse_share(link):
    """A viewer Share link -> (frame id, hlookat, vlookat, fov). The share value is base64 msgpack
    [version, [id, heading, height, lat, lon, date, ?, tileFormat], [hlookat, vlookat, fov], ...]."""
    import base64
    q = parse_qs(urlparse(str(link).strip()).query)
    val = (q.get("share") or [""])[0]
    if not val:
        raise BadRequest("That isn't a 360 viewer Share link (it has no share= part).")
    try:
        data = _msgpack(base64.b64decode(val + "=" * (-len(val) % 4)))
        fid = str(data[1][0])
        h, v, fov = (float(x) for x in data[2][:3])
    except (ValueError, TypeError, IndexError, KeyError) as e:
        raise BadRequest(f"Couldn't read that Share link ({e}).")
    if not DT_ID.fullmatch(fid):
        raise BadRequest("Share link has a bad frame id.")
    return fid, h, v, fov


def _dt_render(pano_base, fid, h, v, hfov, W=1600, H=1000):
    import io
    from concurrent.futures import ThreadPoolExecutor
    import numpy as np
    from PIL import Image
    # Pick the pyramid level whose resolution covers the output (level 4 = 16384 px around).
    level = 3 if (360 / hfov) * W <= 8192 * 1.25 else 4
    T, cols, rows = 1024, 2 ** level, 2 ** (level - 1)
    PW, PH = cols * T, rows * T
    f = (W / 2) / math.tan(math.radians(hfov) / 2)
    X, Y = np.meshgrid(np.arange(W) - (W - 1) / 2, np.arange(H) - (H - 1) / 2)
    Z = np.full_like(X, f)
    p = math.radians(v)  # krpano vlookat: positive looks down
    Y2, Z2 = Y * math.cos(p) + Z * math.sin(p), -Y * math.sin(p) + Z * math.cos(p)
    yaw = np.arctan2(X, Z2) + math.radians(h)
    pitch = np.arctan2(Y2, np.hypot(X, Z2))
    u = np.clip(((yaw / (2 * math.pi) + 0.5) % 1.0) * PW, 0, PW - 1.001)
    w = np.clip((pitch / math.pi + 0.5) * PH, 0, PH - 1.001)
    tc, tr = (u // T).astype(int), (w // T).astype(int)
    need = sorted(set(zip(tc.ravel().tolist(), tr.ravel().tolist())))

    def get(cr):
        r = requests.get(f"{pano_base}/{fid}/{level}_{cr[0] + 1}_{cr[1] + 1}.jpg", timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"360 picture tile {level}_{cr[0] + 1}_{cr[1] + 1}: HTTP {r.status_code}")
        return cr, np.asarray(Image.open(io.BytesIO(r.content)).convert("RGB"))
    with ThreadPoolExecutor(8) as ex:
        tiles = dict(ex.map(get, need))
    out = np.zeros((H, W, 3), np.uint8)
    for (c, r_), img in tiles.items():
        m = (tc == c) & (tr == r_)
        out[m] = img[(w[m] - r_ * T).astype(int), (u[m] - c * T).astype(int)]
    return Image.fromarray(out)


def _burn_citation(img, lines):
    from PIL import ImageDraw, ImageFont
    d = ImageDraw.Draw(img)
    size = max(14, img.width // 64)
    font = None
    for fp in ("/System/Library/Fonts/Supplemental/Arial Bold.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
               "/Library/Fonts/Arial.ttf"):
        try:
            font = ImageFont.truetype(fp, size)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default(size=size)
    pad, gap = size // 2, size // 4
    boxes = [d.textbbox((0, 0), t, font=font) for t in lines]
    tw = max(b[2] - b[0] for b in boxes)
    th = sum(b[3] - b[1] for b in boxes) + gap * (len(lines) - 1)
    x1, y1 = img.width - size // 2, img.height - size // 2
    x0, y0 = x1 - tw - 2 * pad, y1 - th - 2 * pad
    d.rounded_rectangle((x0, y0, x1, y1), radius=size // 3, fill=(26, 26, 26))
    y = y0 + pad
    for t, b in zip(lines, boxes):
        d.text((x1 - pad - (b[2] - b[0]), y - b[1]), t, font=font, fill=(255, 255, 255))
        y += (b[3] - b[1]) + gap
    return img


def save_dt360(ds, b, who):
    """Save a still from the DebrisTech 360 drive into the building's photo folder, with the
    citation burned into the bottom-right corner. Without a share link it faces the building
    from the nearest frame; with one it uses that frame and angle (the frame must be near)."""
    import io
    if not ds.dt360_tiles or not ds.photos:
        raise BadRequest("360 saving is not set up for this dataset")
    try:
        row = int(b.get("row"))
    except (TypeError, ValueError):
        raise BadRequest("bad row")
    rec = next((r for r in ds.all_rows() if r.get("_row") == row), None)
    if rec is None:
        raise BadRequest("unknown row")
    try:
        blat, blon = float(rec.get("latitude")), float(rec.get("longitude"))
    except (TypeError, ValueError):
        raise BadRequest("this building has no coordinates")
    folders = photo_folders(ds)["folders"]
    folder = b.get("folder") or ""
    if folder and folder not in folders:
        raise BadRequest("unknown photo folder")
    if not folder:
        folder = safe_folder_name(rec.get("complete_address") or rec.get("complete address") or f"row {row}")
        if not folder:
            raise BadRequest("this building has no address for a photo folder")
    frames = _dt_frames_near(ds, blat, blon)
    near_m = float(CFG.get("dt360_near_m", 150))
    if b.get("share"):
        fid, h, v, fov = dt360_parse_share(b["share"])
        fr = next((f for f in frames if f["id"] == fid), None)
        if fr is None or _dist_m(blat, blon, fr["lat"], fr["lon"]) > near_m:
            raise BadRequest(f"That Share link is for a spot more than {near_m:.0f} m from this building. "
                             "Check you have the right building open.")
        how = "share"
    else:
        fr = min(frames, key=lambda f: _dist_m(blat, blon, f["lat"], f["lon"]), default=None)
        if fr is None or _dist_m(blat, blon, fr["lat"], fr["lon"]) > near_m:
            raise BadRequest(f"The 360 drive didn't pass within {near_m:.0f} m of this building.")
        h, v, fov = (_bearing(fr["lat"], fr["lon"], blat, blon) - fr["heading"]) % 360, -8.0, 100.0
        how = "facing"
    if not DT_ID.fullmatch(fr["id"]):
        raise BadRequest("bad frame id in the 360 index")
    v = max(-90.0, min(90.0, v))
    fov = max(30.0, min(120.0, fov))
    pano_base = ds.dt360_tiles.rstrip("/").rsplit("/", 2)[0] + "/panorama-tiles"
    img = _dt_render(pano_base, fr["id"], h, v, fov)
    date = fr["date"] or "undated"
    label = ds.dt360_label
    _burn_citation(img, ["Imagery © DebrisTech 360° viewer",
                         f"{label + ', ' if label else ''}captured {date}"])
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=90)
    bearing = (h + fr["heading"]) % 360
    fname = f"DT360_after_{date}_h{bearing:03.0f}.jpg"
    path = quote(f"{ds.photos}/{folder}/{fname}")
    it = graph("PUT", f"{ds.drive}/root:/{path}:/content?@microsoft.graph.conflictBehavior=rename",
               data=buf.getvalue(), headers_extra={"Content-Type": "image/jpeg"})
    ds._photo_folders = None
    dist = round(_dist_m(blat, blon, fr["lat"], fr["lon"]))
    audit({"dataset": ds.key, "op": "dt360", "row": row, "who": who, "folder": folder, "file": it.get("name"),
           "frame": fr["id"], "how": how, "bearing": round(bearing, 1), "pitch": round(v, 1), "fov": round(fov),
           "distance_m": dist})
    return 200, {"file": it.get("name"), "webUrl": it.get("webUrl"), "folder": folder, "date": date,
                 "bearing": round(bearing), "distance_m": dist}


class BadRequest(Exception):
    pass


def audit(rec):
    rec["at"] = datetime.now(timezone.utc).isoformat()
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    with open(STATE_DIR / "writes.jsonl", "a") as fh:
        fh.write(json.dumps(rec, default=str) + "\n")


# ── HTTP ────────────────────────────────────────────────────────────

CFG = {}
DATASETS = {}


class Handler(BaseHTTPRequestHandler):
    server_version = "recon-entry/1"

    def log_message(self, fmt, *args):
        log.info("%s %s", self.headers.get("X-Forwarded-For", self.client_address[0]), fmt % args)

    def cors(self):
        origin = self.headers.get("Origin", "")
        if origin in CFG.get("allowed_origins", []):
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, PATCH, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Lab-Code, X-Entered-By, ngrok-skip-browser-warning")
            self.send_header("Access-Control-Max-Age", "600")

    def send(self, code, obj):
        body = json.dumps(obj, default=str).encode()
        self.send_response(code)
        self.cors()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.cors()
        self.end_headers()

    def authed(self):
        code = self.headers.get("X-Lab-Code", "").strip().lower()  # students may type H3D
        if hmac.compare_digest(code.encode(), CFG["lab_code"].strip().lower().encode()):
            return True
        time.sleep(1)
        self.send(401, {"error": "bad_code", "message": "Lab code not recognized."})
        return False

    def route(self):
        parts = [p for p in self.path.split("?")[0].split("/") if p]
        if parts[:1] == ["recon"]:  # public path on the Tailscale funnel, whether or not it is stripped
            parts = parts[1:]
        if parts[:1] != ["api"]:
            return None, None, None
        ds = DATASETS.get(parts[1]) if len(parts) > 1 else None
        return parts, ds, parts[2:] if len(parts) > 2 else []

    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > MAX_BODY:
            raise BadRequest("body too large")
        return json.loads(self.rfile.read(n) or b"{}")

    def handle_safely(self, fn):
        try:
            fn()
        except BadRequest as ex:
            self.send(400, {"error": "bad_request", "message": str(ex)})
        except Exception as ex:
            log.error("request failed: %s %s\n%s", self.command, self.path, traceback.format_exc())
            self.send(502, {"error": "server_error", "message": str(ex)[:500]})

    def do_GET(self):
        def go():
            parts, ds, rest = self.route()
            if parts == ["api", "health"]:
                return self.send(200, {"ok": True, "datasets": sorted(DATASETS)})
            if parts is None or not self.authed():
                return parts is None and self.send(404, {"error": "not_found"})
            if ds is None:
                return self.send(404, {"error": "unknown_dataset"})
            if rest == ["schema"]:
                return self.send(200, ds.schema())
            if rest == ["rows"]:
                return self.send(200, {"rows": ds.all_rows()})
            if rest == ["dt360"]:
                q = parse_qs(urlparse(self.path).query)
                try:
                    lat, lon = float(q["lat"][0]), float(q["lon"][0])
                except (KeyError, ValueError):
                    raise BadRequest("lat and lon required")
                return self.send(200, dt360_nearest(ds, lat, lon))
            if rest == ["photo-folders"]:
                pf = photo_folders(ds)
                return self.send(200, {"folders": pf["folders"], "webUrl": pf["webUrl"]})
            if rest == ["photos"]:
                q = parse_qs(urlparse(self.path).query).get("folder", [""])[0]
                return self.send(200, photos_in(ds, q))
            self.send(404, {"error": "not_found"})
        self.handle_safely(go)

    def do_POST(self):
        def go():
            parts, ds, rest = self.route()
            if parts is None or not self.authed():
                return parts is None and self.send(404, {"error": "not_found"})
            if ds is not None and rest == ["dt360-save"]:
                code, out = save_dt360(ds, self.body(), self.headers.get("X-Entered-By", ""))
                return self.send(code, out)
            if ds is not None and rest == ["streetview-save"]:
                code, out = save_streetview(ds, self.body(), self.headers.get("X-Entered-By", ""))
                return self.send(code, out)
            if ds is None or rest != ["rows"]:
                return self.send(404, {"error": "not_found"})
            b = self.body()
            code, out = ds.append(b.get("fields", {}), self.headers.get("X-Entered-By", ""))
            self.send(code, out)
        self.handle_safely(go)

    def do_PATCH(self):
        def go():
            parts, ds, rest = self.route()
            if parts is None or not self.authed():
                return parts is None and self.send(404, {"error": "not_found"})
            if ds is None or len(rest) != 2 or rest[0] != "rows" or not rest[1].isdigit():
                return self.send(404, {"error": "not_found"})
            b = self.body()
            if not b.get("hash"):
                raise BadRequest("missing hash")
            code, out = ds.update(int(rest[1]), b["hash"], b.get("fields", {}), self.headers.get("X-Entered-By", ""))
            self.send(code, out)
        self.handle_safely(go)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8787)
    ap.add_argument("--config", type=Path, default=CONFIG_PATH)
    args = ap.parse_args()
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(), logging.FileHandler(STATE_DIR / "server.log")])
    CFG.update(json.loads(args.config.read_text()))
    if not CFG.get("lab_code", "").strip():
        sys.exit("config lab_code missing")
    for k, c in CFG["datasets"].items():
        DATASETS[k] = Dataset(k, c)
    token()  # fail at startup, not on the first student request
    log.info("listening on 127.0.0.1:%d, datasets %s", args.port, sorted(DATASETS))
    ThreadingHTTPServer(("127.0.0.1", args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
