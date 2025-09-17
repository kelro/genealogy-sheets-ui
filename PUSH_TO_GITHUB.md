# Push a Project to GitHub — **Command‑Line First** (Forget‑Proof Guide)

This guide assumes you have **Git** and **GitHub CLI (`gh`)** installed and authenticated (`gh auth login`).  
It **forces** the command‑line workflow so you don’t have to remember web steps.

> Keep this next to you. Follow it top‑to‑bottom every time.

---

## 0) One‑time setup checks (skip if already done)

```bash
# Identify yourself for commits (run once per machine)
git config --global user.name "Ronald Kelley"
git config --global user.email "kelro.privet@gmail.com"

# Make sure gh is logged in
gh auth status
# If not authenticated:
gh auth login
```

---

## 1) Create or open your project folder

```bash
# Example: use your existing folder
cd ~/Documents/"Genealogy Sheets"

# (Optional) If starting a brand new folder:
# mkdir -p ~/Documents/"My New Project"
# cd ~/Documents/"My New Project"
```

Check that the files you want to publish are present (README.md, .gitignore, LICENSE, Code.gs, index.html, sidebar_*.html, etc.).

---

## 2) Initialize the local Git repo and make the FIRST commit

> You **must** have at least one commit before you can push.

```bash
git init -b main        # create a new repo with 'main' as default branch
git status              # see untracked files

git add .               # stage everything in the folder
git commit -m "Initial commit: Genealogy Sheets UI (Apps Script)"
```

If you forget a file:
```bash
git add <path/to/file>
git commit -m "Add missing file <name>"
```

---

## 3) Create the GitHub repo FROM THE COMMAND LINE (no web)

Choose visibility: `--public` **or** `--private` (pick one). This also sets `origin` for you.

```bash
# PUBLIC repo
gh repo create genealogy-sheets-ui --source=. --public

# or PRIVATE repo
# gh repo create genealogy-sheets-ui --source=. --private
```

> `gh` will show the new repo URL and confirm that it added the `origin` remote.

---

## 4) Push the local `main` branch to GitHub

```bash
git push -u origin main
```

You’re live. Visit the printed URL (e.g., https://github.com/kelro/genealogy-sheets-ui).

**If you see** `error: src refspec main does not match any`  
→ You skipped **Step 2** (no first commit). Go back, commit, then push again.

---

# Everyday Updates — Cheat Sheet

## A) When you change a code file

```bash
git status                         # see changes
git add <file1> <file2>            # stage specific files (or 'git add -A' for all changes)
git commit -m "Explain what changed briefly"
git pull --rebase origin main      # sync with remote (replays your commits on top)
git push
```

Common variations:
```bash
git diff                           # see what changed (unstaged)
git diff --staged                  # see what's staged
```

## B) When you add a NEW file

```bash
git add <newfile>
git commit -m "Add <newfile> and explain purpose"
git push
```

## C) When you rename or move a file

```bash
git mv old/path/File.gs new/path/File.gs
git commit -m "Rename/move File.gs to new/path"
git push
```

## D) When you delete a file

```bash
git rm <file>
git commit -m "Remove <file> (reason)"
git push
```

## E) When you want a clean “add everything” update

```bash
git add -A
git commit -m "Update: short summary of changes"
git pull --rebase origin main
git push
```

---

# Branching (good hygiene for bigger edits)

```bash
git checkout -b feature/timeline-print-fix
# ...edit files...
git add -A
git commit -m "Fix timeline print scaling"
git push -u origin feature/timeline-print-fix
```
Then open a Pull Request on GitHub from `feature/timeline-print-fix` → `main`. After merge, delete the branch.

---

# Version Tags & Releases

```bash
git tag -a v1.0.0 -m "Initial public release"
git push origin v1.0.0
```
On GitHub → Releases → “Draft a new release” → choose `v1.0.0` → add notes → publish.

---

# Quick Fixes (Copy/Paste)

**Remote already set to a different URL**  
```bash
git remote -v
git remote set-url origin https://github.com/<USER>/<REPO>.git
# or SSH:
# git remote set-url origin git@github.com:<USER>/<REPO>.git
```

**“src refspec main does not match any”** (no commits on main)  
```bash
git add .
git commit -m "Initial commit"
git push -u origin main
```

**“Updates were rejected because the remote contains work…”** (remote has commits you don’t)  
```bash
git pull --rebase origin main
# resolve any conflicts if asked:
#   edit files to resolve, then
git add <resolved-files>
git rebase --continue
git push
```

**Change default branch name to `main` if your Git created `master`**  
```bash
git branch -M main
git push -u origin main
```

**Store HTTPS credentials to avoid repeated prompts**  
```bash
git config --global credential.helper store
# Next push will prompt once; credentials are then remembered.
```

**See recent history and branches at a glance**  
```bash
git log --oneline --graph --decorate --all
```

---

# Recommended Repo Layout

```
/ (repo root)
  README.md
  LICENSE
  .gitignore
  Code.gs
  index.html
  sidebar_person.html
  sidebar_marriage.html
  sidebar_child.html
  sidebar_search.html
  sidebar_bootstrap.html
  ui_common.html
  importer_Prod.gs
  NormalizeDates.gs
  exportNamedRanges.gs
  debugCellType.gs
```

---

# Optional: Make creation + push a single command

You can let `gh` push for you automatically right after creating the repo:

```bash
# After you did Step 2 (first commit):
gh repo create genealogy-sheets-ui --source=. --public --push
```

If this fails with a push error, just run:
```bash
git push -u origin main
```

---

## TL;DR core loop (for updates)

```bash
git status
git add -A
git commit -m "Your message"
git pull --rebase origin main
git push
```

---

# Using clasp (Command Line Apps Script Projects)

If you want to **sync code directly between Google Sheets Apps Script and GitHub**, use `clasp`.

## 1) Install clasp (one time)
```bash
npm install -g @google/clasp
clasp --version
```

## 2) Authenticate with Google
```bash
clasp login
```
Opens a browser — log in with your Google account.

## 3) Initialize a local project
In your repo folder:
```bash
cd ~/Documents/"Genealogy Sheets"

# If creating a new Apps Script project tied to a Sheet:
clasp create --title "Genealogy Sheets UI" --type sheets

# If you already have a script in the Sheet, clone it with Script ID:
clasp clone <SCRIPT_ID>
```
- Find the **Script ID** in Google Sheets: Extensions → Apps Script → Project Settings.

## 4) Pull existing Apps Script code from Google
```bash
clasp pull
```
This downloads `.gs` and `.html` files into your repo.

## 5) Push local changes up to Google
```bash
clasp push
```
This uploads your repo’s `.gs` and `.html` files into the bound Apps Script project.

## 6) Typical workflow
- Run `clasp pull` before editing to sync down changes from the online editor.
- Edit files locally → commit → push to GitHub.
- Run `clasp push` to sync changes back to Google.

## 7) Git ignore
Add `.clasp.json` to your `.gitignore` (already included) so project IDs are not shared.  
Keep `appsscript.json` in GitHub — it defines project settings/scopes.

---


---

# Google Apps Script Sync (clasp) — Pull from Sheets, Push to GitHub, and Back

Use **clasp** (official Google CLI) to sync your Apps Script project (the code behind your Google Sheet) with your local repo.

## A) Install once
```bash
# Requires Node.js and npm
npm install -g @google/clasp
clasp --version
```

## B) Authenticate once (per machine)
```bash
clasp login
# A browser window opens to authorize your Google account
```

## C) Initialize in your repo folder
Go to your repo folder (where README.md is):
```bash
cd ~/Documents/"Genealogy Sheets"
```

### If you already have a script attached to your Sheet (container-bound)
1) Find the **Script ID** in **Extensions → Apps Script → Project Settings → Script ID**.  
2) Clone it locally:
```bash
clasp clone <SCRIPT_ID>
```
This creates `.clasp.json` (project metadata) and pulls code.\
*(Leave `appsscript.json` tracked in Git — it defines manifest/scopes.)*

### If you want to create a brand-new script project
```bash
clasp create --title "Genealogy Sheets UI" --type sheets
```
This sets up a new Apps Script container for Sheets and links it to your folder.

## D) Pull code down from Google (overwrite local)
```bash
clasp pull
```
> Use when you edited code in the Apps Script editor and want it locally.

## E) Push local code up to Google (overwrite remote)
```bash
clasp push
# If Google warns about file diffs you want to override:
# clasp push --force
```
> Use after local edits (and after pulling latest from GitHub).

## F) Typical day-to-day loop with GitHub + clasp
```bash
# 1) Get latest from Google (if you edited online)
clasp pull

# 2) Commit to GitHub
git add -A
git commit -m "Update: <what changed>"
git push

# 3) (Optional) Push back to Google
clasp push
```

## G) Helpful commands
```bash
clasp open          # opens the Apps Script editor in your browser
clasp status        # shows local vs remote file differences
clasp logs --watch  # tail runtime logs for executions
```

## H) .gitignore note
Keep the manifest **tracked**; hide project identifiers:
```
# .gitignore recommendations
.clasp.json
# appsscript.json  <-- DO NOT ignore; keep this tracked
```
If your current `.gitignore` ignores `appsscript.json`, remove that line so the manifest is versioned.

