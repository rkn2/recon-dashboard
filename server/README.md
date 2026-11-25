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

## Operating

```
launchctl kickstart -k gui/$(id -u)/com.becca.recon-entry   # restart after editing config
tail -f ~/.local/state/recon-entry/server.log
# funnel route (Becca's own tailscaled, userspace):
/opt/homebrew/opt/tailscale/bin/tailscale --socket=$HOME/.local/state/becca-tailscale/tailscaled.sock funnel status
```

Testing the form locally: serve the repo on port 8790 and open
`http://localhost:8790/tornado.html?api=http://localhost:8787`.
