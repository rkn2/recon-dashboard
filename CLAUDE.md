# recon-dashboard

Web form and dashboard for lab students to enter and review building damage data
(St. Louis tornado, Marshall and Montpelier floods). Static pages on GitHub Pages
(`main` is live for every student); data lives in OneDrive Excel workbooks, reached
through `server/recon_entry.py` on Becca's Mac mini. See `server/README.md`.

The person working in this repo may be an undergraduate learning to code
(see ONBOARDING.md). Explain what you change and why, in plain words, and keep
changes small enough for them to review.

## Rules

- **Never commit secrets.** This repo is public. No lab code, tokens, passwords,
  or anything from `~/.config/recon-entry/`. If you see one in a diff, stop and say so.
- **Never push to `main` or merge.** Work on a branch and open a pull request for
  Becca (`gh pr create`).
- **Don't change `server/`** unless the user is Becca. Suggest server changes in the
  pull request description instead.
- **Test copies only.** Develop against the test datasets/workbooks, never the live
  ones, unless Becca says otherwise.
- **No student personal information** in code, commits, or test data.
- **Verify in a browser** before calling a change done: serve the repo with
  `python3 -m http.server 8790` and load the page. Say what you checked.

## Code conventions

- Plain HTML, CSS, and JavaScript in single files; no build step, no frameworks.
- Match the existing style in the file you're editing. Colors are CSS variables on
  `:root` with a dark-mode block; reuse them.
- Field keys must match the workbook column names exactly (including the
  spreadsheet's spellings, e.g. `buidling_height_m`). The server rejects unknown
  columns.
- Errors must be visible to the student (status bar or alert), never swallowed.
