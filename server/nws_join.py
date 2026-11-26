"""Fill NWS damage survey values into a recon workbook.

For each building row with coordinates, finds the nearest building point in the NWS Damage
Assessment Toolkit (DAT) survey for the storm and writes it into columns that start with
"nws_". Each DAT point goes to at most one building (closest pairs first), so a point on one
house isn't copied onto its neighbours. Tree and other non-building points are skipped.
Only the nws_ columns are written. They're added at the right end of the sheet if missing.

  uv run --with requests --with msal server/nws_join.py --dataset stlouis \\
      --start 2025-05-16 --end 2025-05-18 [--path <test copy> --drive me] [--write] [--csv out.csv]

Without --write it only prints what it would do. --csv writes the rows where the student
marked the building undamaged but NWS rated it damaged (addresses only, no names).
"""
import argparse, csv, json, math, os, sys, time
from datetime import date

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import recon_entry as R  # noqa: E402

DAT = ("https://services.dat.noaa.gov/arcgis/rest/services/nws_damageassessmenttoolkit/"
       "DamageViewer/FeatureServer/0/query")
COLS = ["nws_match", "nws_distance_m", "nws_damage_indicator", "nws_degree_of_damage",
        "nws_degree_of_damage_text", "nws_ef", "nws_wind_mph", "nws_point_id"]
MAX_M = 20.0
NOT_BUILDINGS = ("Trees", "Transmission", "Electrical", "Free-Standing", "Free Standing")


def fetch_points(start, end, bbox):
    """All DAT damage points in bbox (minlon, minlat, maxlon, maxlat) between start and end dates."""
    out, offset = [], 0
    while True:
        r = requests.get(DAT, params={
            "where": f"stormdate >= DATE '{start}' AND stormdate < DATE '{end}'",
            "geometry": ",".join(map(str, bbox)), "geometryType": "esriGeometryEnvelope", "inSR": 4326,
            "spatialRel": "esriSpatialRelIntersects", "outFields": "*", "returnGeometry": "false",
            "resultOffset": offset, "resultRecordCount": 2000, "f": "json"}, timeout=60)
        r.raise_for_status()
        j = r.json()
        if "error" in j:
            raise RuntimeError(f"DAT query failed: {j['error']}")
        out += [f["attributes"] for f in j["features"]]
        offset += len(j["features"])
        if not j.get("exceededTransferLimit") or not j["features"]:
            return out


def is_building(p):
    t = str(p.get("damage_txt") or "")
    return t and not t.startswith(NOT_BUILDINGS) and p.get("lat") and p.get("lon")


def dist_m(a, b, c, d):
    k = math.cos(math.radians(a))
    return math.hypot((a - c) * 111320, (b - d) * 111320 * k)


def match(buildings, points):
    """buildings: {row: (lat, lon)}. Returns {row: (point, distance)} and the set of rows whose
    near point was taken by a closer building."""
    cell = 0.0005  # ~50 m grid for the candidate search
    grid = {}
    for p in points:
        grid.setdefault((int(p["lat"] / cell), int(p["lon"] / cell)), []).append(p)
    pairs = []
    for row, (la, lo) in buildings.items():
        gi, gj = int(la / cell), int(lo / cell)
        for di in (-1, 0, 1):
            for dj in (-1, 0, 1):
                for p in grid.get((gi + di, gj + dj), []):
                    d = dist_m(la, lo, p["lat"], p["lon"])
                    if d <= MAX_M:
                        pairs.append((d, row, p["objectid"], p))
    pairs.sort(key=lambda x: (x[0], x[1], x[2]))
    got, used, near = {}, set(), set()
    for d, row, oid, p in pairs:
        near.add(row)
        if row in got or oid in used:
            continue
        got[row] = (p, d)
        used.add(oid)
    return got, near - set(got)


def num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError):
        return v if v not in (None, "") else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.path.expanduser("~/.config/recon-entry/config.json"))
    ap.add_argument("--dataset", required=True)
    ap.add_argument("--start", required=True, help="first storm date, YYYY-MM-DD")
    ap.add_argument("--end", required=True, help="day after the last storm date, YYYY-MM-DD")
    ap.add_argument("--path", help="workbook path instead of the dataset's (for a test copy)")
    ap.add_argument("--drive", help="'me' for your own OneDrive, else the dataset's drive")
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--csv")
    ap.add_argument("--state-dir", help="where the audit line goes (default: the server's state folder)")
    a = ap.parse_args()

    cfg = json.load(open(a.config))
    R.CFG.update(cfg)
    dcfg = dict(cfg["datasets"][a.dataset])
    if a.path:
        dcfg["path"] = a.path
    if a.drive == "me":
        dcfg.pop("drive", None)
    if a.state_dir:
        R.STATE_DIR = R.Path(a.state_dir)
    ds = R.Dataset(a.dataset, dcfg)
    print("workbook:", dcfg.get("drive", "me"), dcfg["path"])

    headers, rows = ds.load(fresh=True)
    idx = ds.first_index(headers)
    buildings = {}
    for n, vals in rows:
        try:
            la, lo = float(vals[idx["latitude"]]), float(vals[idx["longitude"]])
        except (ValueError, TypeError, IndexError):
            continue
        if -90 < la < 90 and -180 < lo < 180 and (la, lo) != (0, 0):
            buildings[n] = (la, lo)
    if not buildings:
        sys.exit("no rows with coordinates")
    lats = [v[0] for v in buildings.values()]
    lons = [v[1] for v in buildings.values()]
    bbox = (min(lons) - 0.002, min(lats) - 0.002, max(lons) + 0.002, max(lats) + 0.002)
    pts = fetch_points(a.start, a.end, bbox)
    bpts = [p for p in pts if is_building(p)]
    print(f"DAT points in the area: {len(pts)}, building points: {len(bpts)}")
    got, taken = match(buildings, bpts)

    # Column positions: reuse existing nws_ columns, else append after the last header.
    last = max(i for i, h in enumerate(headers) if h) + 1
    pos, add = {}, []
    for c in COLS:
        if c in idx:
            pos[c] = idx[c]
        else:
            pos[c] = last + len(add)
            add.append(c)
    first, lastc = min(pos.values()), max(pos.values())
    if lastc - first + 1 != len(COLS) or sorted(pos.values()) != [pos[c] for c in COLS]:
        sys.exit(f"nws_ columns aren't together in order: {pos}")
    for c in add:  # new columns must be empty in every row
        for n, vals in rows:
            if pos[c] < len(vals) and str(vals[pos[c]]).strip():
                sys.exit(f"column {R.col_letter(pos[c] + 1)} (for {c}) has data in row {n}")

    last_row = max(n for n, v in rows if any(str(x).strip() for x in v))
    today = date.today().isoformat()
    block, counts, flag = [], {}, []
    for n in range(ds.header_row + 1, last_row + 1):
        if n in got:
            p, d = got[n]
            line = ["matched", round(d, 1), p.get("damage_txt") or "", num(p.get("dod")), p.get("dod_txt") or "",
                    p.get("efscale") or "", num(p.get("windspeed")), num(p.get("objectid"))]
        elif n in taken:
            line = ["nearest NWS point belongs to another building"] + [""] * 7
        elif n in buildings:
            line = [f"no NWS point within {MAX_M:.0f} m"] + [""] * 7
        else:
            line = ["no coordinates"] + [""] * 7
        counts[line[0]] = counts.get(line[0], 0) + 1
        block.append(line)
    print("rows:", counts)

    if a.csv:
        st = idx.get("status_u")
        ad = idx.get("complete_address")
        byrow = dict(rows)
        with open(a.csv, "w", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["row", "address", "student overall status", "NWS degree of damage", "NWS description",
                        "NWS EF", "NWS wind mph", "distance m", "latitude", "longitude"])
            for n, (p, d) in sorted(got.items()):
                s = str(byrow[n][st]).strip() if st is not None else ""
                if s == "undamaged" and (num(p.get("dod")) or 0) >= 2:
                    w.writerow([n, byrow[n][ad] if ad is not None else "", s, num(p.get("dod")), p.get("dod_txt"),
                                p.get("efscale"), num(p.get("windspeed")), round(d, 1), *buildings[n]])
                    flag.append(n)
        print(f"wrote {len(flag)} disagreements to {a.csv}")

    a1, a2 = R.col_letter(first + 1), R.col_letter(lastc + 1)
    print(f"would write columns {a1}:{a2}, rows {ds.header_row}-{last_row}" + (f" (new: {add})" if add else ""))
    if not a.write:
        print("dry run, nothing written")
        return
    ds.wb("PATCH", f"{ds.ws_url()}/range(address='{a1}{ds.header_row}:{a2}{ds.header_row}')", json={"values": [COLS]})
    if ds.header_row > 1:
        above = ds.wb("GET", f"{ds.ws_url()}/range(address='{a1}{ds.header_row - 1}:{a2}{ds.header_row - 1}')?$select=values")
        if any(str(v).strip() and not str(v).startswith("NWS DAT survey") for v in above["values"][0]):
            sys.exit(f"row {ds.header_row - 1} above the nws_ columns isn't empty")
        note = [f"NWS DAT survey, filled by nws_join.py on {today}. Don't edit."] + [""] * (len(COLS) - 1)
        ds.wb("PATCH", f"{ds.ws_url()}/range(address='{a1}{ds.header_row - 1}:{a2}{ds.header_row - 1}')",
              json={"values": [note]})
    for i in range(0, len(block), 400):
        r0 = ds.header_row + 1 + i
        chunk = block[i:i + 400]
        ds.wb("PATCH", f"{ds.ws_url()}/range(address='{a1}{r0}:{a2}{r0 + len(chunk) - 1}')", json={"values": chunk})
    R.audit({"dataset": ds.key, "op": "nws_join", "who": "nws_join.py", "path": dcfg["path"], "columns": f"{a1}:{a2}",
             "rows": counts, "dat_points": len(bpts), "storm": [a.start, a.end]})
    print("written")


if __name__ == "__main__":
    main()
