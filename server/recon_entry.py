#!/usr/bin/env python3
"""Recon data-entry server: lets the GitHub Pages form read and write the
OneDrive DataInput workbooks through Microsoft Graph.

Runs on the lab Mac mini under Becca's account and uses her delegated
"E5-Nap API" token (the MSAL cache from ~/Code/admin/general/api-access).
The token never leaves this machine: the form talks to this server, and this
server only reads and updates building rows in the workbooks named in the
config. No delete, no other files.

Config (outside the repo): ~/.config/recon-entry/config.json
  {
    "lab_code": "...",                      # students type this once in the form
    "allowed_origins": ["https://rkn2.github.io"],
    "datasets": {
      "stlouis": {"path": "students/undergrads/.../StLouis_Tornado_DataInput.xlsx",
                  "sheet": "BuidlingAttributes", "header_row": 2},
      ...
    }
  }

Run:  uv run --with requests --with msal python server/recon_entry.py [--port 8787]

Every write is appended to ~/.local/state/recon-entry/writes.jsonl with the
old and new value of each cell, so any save can be traced and undone.
"""
import argparse
import hashlib
import hmac
import json
import logging
import re
import sys
import threading
import time
import traceback
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote

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


def graph(method, url, session=None, **kw):
    headers = {"Authorization": "Bearer " + token()}
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
        self.lock = threading.Lock()  # serializes writes to this workbook
        self._item = None
        self._session = None  # persistent workbook session: reads see earlier writes
        self._session_lock = threading.Lock()
        self._cache = None  # (time, headers, rows)

    def item_url(self):
        if self._item is None:
            it = graph("GET", f"/me/drive/root:/{quote(self.path)}?$select=id,parentReference")
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
            self.send(404, {"error": "not_found"})
        self.handle_safely(go)

    def do_POST(self):
        def go():
            parts, ds, rest = self.route()
            if parts is None or not self.authed():
                return parts is None and self.send(404, {"error": "not_found"})
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
