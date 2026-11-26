# recon-entry server

Lets the data-entry form (`tornado.html`) read and write the OneDrive DataInput
workbooks. Students don't sign in to Microsoft; they type a lab code once.

```
form (GitHub Pages) ──HTTPS──> Tailscale funnel ──> recon_entry.py on the lab Mac mini
                                                       └─ Microsoft Graph (Becca's E5-Nap token) ──> OneDrive .xlsx
```

- Runs as Becca on the Mac mini (launchd: `~/Library/LaunchAgents/com.becca.recon-entry.plist`,
  port 8787 on localhost only). Her token stays on that machine.
- Public address: https://becca-mcp.tail3a44c1.ts.net/recon (a path on Becca's Tailscale funnel, the same one the
  HARB Sanborn checking page uses at `/sanborn`). The server accepts paths with or without the `/recon` prefix.
  Chrome on a computer that is itself on Becca's tailnet can't reach this address; students' computers are fine.
- The launchd job sets `UV_CACHE_DIR` and `UV_LINK_MODE=copy` to match Becca's shell. Without them uv fails to start
  the server whenever it has to download a package (the default cache links onto the Samsung drive; "Cross-device
  link", seen 2026-11-22). So the server needs the Samsung drive mounted when it starts.
- The launchd job runs `uv run --with requests --with msal --with numpy --with pillow` (numpy and Pillow cut the 360
  stills). After changing the plist's arguments, `kickstart` isn't enough; use `launchctl bootout` then `bootstrap`.
- It can only list rows, update cells in an existing row, and add a row, in the
  workbooks named in `~/.config/recon-entry/config.json`. No deletes.
- Each row is found by its row number. Every save carries a fingerprint of the row
  as the student loaded it; if the row changed since (someone edited Excel, sorted,
  inserted rows), the save is refused (409) and the form reloads that building.
- Only changed cells are written. Text starting with `= + - @` is stored as text,
  never a formula.
- Every write is logged with old and new values in
  `~/.local/state/recon-entry/writes.jsonl` (to trace or undo a save).
  Server log: `~/.local/state/recon-entry/server.log`.

## Endpoints

All but `/api/health` need header `X-Lab-Code`.

| | |
|---|---|
| `GET /api/health` | up check |
| `GET /api/<dataset>/schema` | column names |
| `GET /api/<dataset>/rows` | all non-empty rows, each with `_row` and `_hash` |
| `PATCH /api/<dataset>/rows/<n>` | `{"hash": ..., "fields": {...}}` |
| `POST /api/<dataset>/rows` | `{"fields": {...}}` adds a row after the last used one |
| `POST /api/<dataset>/streetview-save` | `{"row", "pano", "heading", "pitch", "fov", "folder"?}` saves a Google Street View still to the building's photo folder |
| `GET /api/<dataset>/dt360?lat=&lon=` | nearest frame of the DebrisTech 360 drive: `{"distance_m", "date"}` |
| `POST /api/<dataset>/dt360-save` | `{"row", "share"?, "folder"?}` saves a 360 still with the DebrisTech citation burned in. Without `share` it faces the building from the nearest frame; with a viewer Share link it uses that frame and angle (the frame must be within 150 m of the building) |

## Operating

```
launchctl kickstart -k gui/$(id -u)/com.becca.recon-entry   # restart after editing config
tail -f ~/.local/state/recon-entry/server.log
# funnel route (Becca's own tailscaled, userspace):
/opt/homebrew/opt/tailscale/bin/tailscale --socket=$HOME/.local/state/becca-tailscale/tailscaled.sock funnel status
```

Testing the form locally: serve the repo on port 8790 and open
`http://localhost:8790/tornado.html?api=http://localhost:8787`.

## NWS damage survey columns (`nws_join.py`)

`nws_join.py` fills eight `nws_` columns (match status, distance, damage indicator, degree of damage and its description, EF, wind speed, DAT point id) from the NWS Damage Assessment Toolkit survey points for the storm dates. Each building gets the closest building point within 20 m, and each point goes to one building only, closest pairs first. Trees, power lines and poles are skipped. Rows without a match say why in `nws_match`. The columns are added at the right end of the sheet the first time, and the server refuses edits to any `nws_` column; the form shows them read-only in the Damage Assessment section.

    uv run --with requests --with msal server/nws_join.py --dataset stlouis --start 2025-05-16 --end 2025-05-18 \
        [--path <test copy> --drive me --state-dir <scratch>] [--csv disagreements.csv] [--write]

It's a dry run without `--write`. Rerunning it rewrites the same columns, so it can be run again when NWS updates the survey.
