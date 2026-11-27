# Onboarding: recon data entry interface

Welcome! This is the web form and dashboard that lab students use to enter and
review building damage data for three disasters:

- **St. Louis** (2025 tornado)
- **Marshall, NC** (2024 Hurricane Helene flood)
- **Montpelier, VT** (2023 flood)

Your job is to make the interface good for all three. The St. Louis form works
today. Marshall and Montpelier need forms, and the dashboard needs to show live
data. Your full project scope is in Becca's notes; the short version is at the
bottom of this page.

You'll work with three tools: **GitHub** (where the code lives), **Claude Code**
(an AI assistant that runs in your terminal and can read and edit the code), and
**herdr** (a terminal app that keeps your Claude sessions organized). Set them up
in this order.

---

## 1. Accounts (day 1)

- [ ] **GitHub account.** Make one at github.com if you don't have one. Send Becca
      your username so she can give you access to `rkn2/recon-dashboard`.
- [ ] **Claude account.** Sign in at claude.ai. Becca will tell you which plan or
      account to use for Claude Code.
- [ ] **OneDrive access.** Becca shares the data folders with your PSU account.
      You only need to look at the spreadsheets, not edit them.
- [ ] **Lab code.** Becca gives you the code the form asks for. Treat it like a
      password: don't post it, don't put it in code, don't commit it.

## 2. Install the tools (Mac)

Open **Terminal** (Cmd+Space, type "Terminal").

```bash
# Homebrew (a Mac package installer), if you don't have it
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Git, the GitHub command line tool, and herdr
brew install git gh herdr

# Claude Code
curl -fsSL https://claude.ai/install.sh | bash
```

Then tell git who you are and sign in to GitHub:

```bash
git config --global user.name "Your Name"
git config --global user.email "jbl6699@psu.edu"
gh auth login        # pick GitHub.com, HTTPS, and "Login with a web browser"
```

On Windows, ask Becca before installing; the steps are different.

## 3. Get the code

```bash
mkdir -p ~/Code && cd ~/Code
gh repo clone rkn2/recon-dashboard
cd recon-dashboard
```

## 4. herdr: your workspace

herdr keeps terminal sessions running in the background, even if you close the
window, and shows you which Claude session is working, waiting on you, or done.

```bash
cd ~/Code/recon-dashboard
herdr
```

- First run shows a short tour.
- Everything is clickable: click panes and tabs, drag borders, right-click for menus.
- Split a pane: right-click, or `ctrl+b` then `v` (side by side) / `ctrl+b` then `-` (stacked).
- Leave herdr without stopping anything: `ctrl+b` then `q`. Come back with `herdr`.
- Show all shortcuts: `ctrl+b` then `?`.

A good layout: **Claude** in the left pane, and in the right pane a local web
server so you can see your changes (step 6).

## 5. Claude Code

In a herdr pane, inside the repo folder:

```bash
claude
```

The first time, it asks you to sign in and whether you trust the folder (yes).
Claude reads this repo's `CLAUDE.md` automatically, which tells it the project's
rules. Good first prompts:

- "Explain how tornado.html saves a building, step by step."
- "What would a Marshall form need that the St. Louis form doesn't have? Compare
  the fields."
- "Make a branch and add X. Show me the diff before committing."

How to work with it well:

- **Read what it changes.** You are responsible for every line you commit, even if
  Claude wrote it. If you can't explain a change, ask Claude to explain it, or
  don't commit it.
- **Small steps.** One feature per branch. Test it in the browser before moving on.
- **Say what you want, not how.** "The status bar should say who last edited the
  building" works better than guessing at code.
- **It can be wrong.** If something looks off, say so. Ask it to check its work
  ("open the page and confirm it saves").
- Press `Esc` to stop Claude mid-task. Type `/help` to see commands.

## 6. Run the form on your computer

In the right-hand pane:

```bash
cd ~/Code/recon-dashboard
python3 -m http.server 8790
```

Open http://localhost:8790/tornado.html in your browser. Becca will tell you
which test address and test data to use while you develop. **You always work
against test copies of the spreadsheets, never the real ones.**

## 7. The GitHub workflow

`main` is the live site that every student uses. Pushing to `main` changes it for
everyone immediately, so you never push to `main` directly.

```bash
git checkout main && git pull            # start from the latest
git checkout -b marshall-form            # one branch per feature
# ... work, test in the browser ...
git add tornado.html                     # add the files you changed
git commit -m "Marshall form: flood height section"
git push -u origin marshall-form
gh pr create                             # opens a pull request for Becca to review
```

Becca reviews the pull request and merges it when it's ready. Claude can run all
of these for you; ask it to "commit this and open a pull request."

More on git: the lab guide at `lab-research-toolkit/git-and-github.md`.

## 8. Rules that matter

1. **Never commit secrets.** No lab code, no passwords, no tokens. The repo is
   public.
2. **Test copies only** until Becca says a change is ready.
3. **Don't edit the server** (`server/`) or anything on Becca's Mac mini. If the
   server needs a change, tell Becca.
4. **Don't push to `main`.** Branches and pull requests only.
5. **No student information** in the repo (grades, schedules, anything personal).

## 9. How it works (big picture)

```
form in the browser ──> lab server on Becca's Mac mini ──> OneDrive Excel workbook
```

- The **Excel workbooks on OneDrive are the record**. The form is a nicer way to
  fill them in.
- The form finds a building by its row number. Every save checks that the row
  hasn't changed since the form loaded it. If someone else changed it, the save is
  refused and the form reloads their version, so nobody's work gets silently
  overwritten.
- Details: `server/README.md`.

## 10. Your project

1. **Marshall and Montpelier forms.** Ideally one form page with an event picker,
   built from each workbook's columns and the flood data input guide.
2. **Live dashboard.** `index.html` shows real data from the server, not the old
   copy pasted into the file.
3. **Data quality views.** Missing fields, inconsistent names ("Sam" vs "Samantha"),
   duplicate addresses, values that aren't allowed codes, progress per student.
4. **A one-page guide** for students entering data.

**First week:** get everything above running, have Claude walk you through
`tornado.html`, compare the Marshall spreadsheet's columns with the flood data
input guide, and show Becca a sketch of the Marshall form before building it.
