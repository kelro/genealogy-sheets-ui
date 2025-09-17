# Push to GitHub Guide

---

# Google Apps Script Sync (clasp) — **Complete, Forget‑Proof Guide**

Use **clasp** (official Google CLI) to pull your Apps Script code **from Google Sheets to your computer**, commit it to **Git/GitHub**, and push edits **back to Google**.

> Follow these steps **in order**. Keep this section handy.

## 0) One‑time prerequisites

### 0.1 Install/upgrade Node with nvm (avoids permissions issues)
`clasp` requires modern Node (v20+). Use **nvm** so global installs don’t need sudo.

```bash
# Install nvm
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.7/install.sh | bash
# Reload your shell
source ~/.bashrc

# Install and use Node 20 (or latest LTS)
nvm install 20
nvm use 20
nvm alias default 20

# Verify
node -v   # should print v20.x.x
npm -v
```

### 0.2 Install clasp
```bash
npm install -g @google/clasp
clasp --version
```

### 0.3 Log in to Google
```bash
clasp login
```
A browser window opens—authorize your Google account.

### 0.4 Enable the **Apps Script API** (only once per Google account)
1) Open: https://script.google.com/home/usersettings  
2) Turn **“Google Apps Script API”** to **ON**.  
3) Wait ~1–2 minutes, then continue.

---

## 1) Choose your starting point

### Case A — You ALREADY have a script attached to an existing Sheet (container‑bound)
1) In the Sheet: **Extensions → Apps Script → Project Settings → Script ID** (copy it).  
2) In your repo folder:
```bash
cd ~/Documents/"Genealogy Sheets"            # your project folder
clasp clone <SCRIPT_ID>                       # pulls the existing script down
```
This creates a `.clasp.json` (metadata) and downloads your `.gs` / `.html` files.

> If you get a 404/permission error, verify you’re logged into the correct Google account with `clasp login` and that the Script ID is correct.

### Case B — You want to CREATE a new Sheet + script from scratch
```bash
cd ~/Documents/"Genealogy Sheets"            # your project folder
clasp create --title "Genealogy Sheets UI" --type sheets
```
`clasp` will create a **new Spreadsheet** with a **bound Apps Script** and link it to this folder.

> To open the script later:
```bash
clasp open
```

---

## 2) Sync commands you will actually use

### Pull from Google → to your computer (overwrite local)
```bash
clasp pull
```
Use this after editing in the online Apps Script editor to refresh local files.

### Push from your computer → to Google (overwrite remote)
```bash
clasp push
# If Google shows conflict warnings you want to override:
# clasp push --force
```
Use this after local edits (and after syncing from GitHub if you collaborate).

### See differences
```bash
clasp status    # shows which files differ between local and remote
```

### Open the Apps Script project in the browser
```bash
clasp open
```

### Tail runtime logs
```bash
clasp logs --watch
```

---

## 3) Git/GitHub + clasp daily loop

```bash
# If you edited online, pull first:
clasp pull

# Save to GitHub
git add -A
git commit -m "Update: <what changed>"
git push

# Optionally push back to Google right away
clasp push
```

---

## 4) Files to track vs ignore

Keep the manifest under version control (important), but hide the project link file:

```
# .gitignore recommendation
.clasp.json        # contains project identifiers (keep private)
# appsscript.json  # DO NOT ignore; keep this tracked in Git
```

If your .gitignore currently ignores `appsscript.json`, **remove that line** so the manifest is versioned.

---

## 5) Common errors & fixes (copy/paste)

**A) “User has not enabled the Apps Script API.”**  
→ Enable it here, wait a minute, retry: https://script.google.com/home/usersettings

**B) “EBADENGINE … requires node >=20.0.0”**  
→ You’re on an older Node. Install nvm and run:
```bash
nvm install 20
nvm use 20
npm install -g @google/clasp
```

**C) “EACCES: permission denied … /usr/local/lib/node_modules”** (global install)  
→ Don’t use `sudo npm install -g`. Install with nvm (above) so global installs go to your home directory.

**D) 404 / permission when cloning**  
- Confirm `clasp login` is using the **same Google account** that owns the script.  
- Double‑check the **Script ID** (Apps Script → Project Settings).

**E) Local/remote differences warnings on push/pull**  
- Use `clasp status` to see which files differ.  
- Decide your direction: `clasp pull` (take Google’s copy) or `clasp push` (send your local copy).  
- If you *intend* to overwrite, use `--force` with care.

**F) Wrong project linked (pulls the wrong code)**  
- Edit `.clasp.json` to change the `scriptId`, or re‑run:
```bash
rm .clasp.json
clasp clone <SCRIPT_ID>
```

**G) “Manifest scopes” or authorization prompts**  
- Pushing a new `appsscript.json` may change scopes; the next run of your script will prompt you to authorize. This is normal—review and allow.

---

## 6) Optional: Web app deploy (read‑only timeline page)
If you add a `doGet(e)` that serves `index.html`, you can deploy a web app:

```bash
clasp deploy --description "Initial web app"
# Later, to list versions/deployments:
clasp deployments
```
Then set permissions in the Apps Script UI for who can access the web app link.


---

## 📄 Reference: Recommended .gitignore

Always include these lines in your `.gitignore`:

```
# macOS
.DS_Store

# Node (if you later add build tooling)
node_modules/
dist/

# Apps Script clasp (safe defaults)
.clasp.json

# Logs / temp exports
*.log
*.tmp
*.bak
```

> Note: Do **NOT** ignore `appsscript.json`. Keep it versioned in GitHub — it defines project scopes and settings.


---

# 🔧 Helper Scripts (Forget‑Proof Automation)

Two Bash scripts are provided to make syncing easier.  
👉 **Always keep these scripts INSIDE your project folder** (where `.git`, `.clasp.json`, and your code files live).

### 📂 Recommended layout
```
~/Documents/Genealogy Sheets/
  README.md
  LICENSE
  .gitignore
  PUSH_TO_GITHUB.md
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
  .clasp.json
  appsscript.json
  sync_apps_script_to_github.sh
  push_local_to_github_and_google.sh
```

### 1) sync_apps_script_to_github.sh
Pull the latest from **Google Apps Script** → commit/push to **GitHub**.  
Optional flag `--push-google` also pushes back to Google.

```bash
cd ~/Documents/"Genealogy Sheets"
chmod +x sync_apps_script_to_github.sh
./sync_apps_script_to_github.sh
./sync_apps_script_to_github.sh -m "Custom commit message"
./sync_apps_script_to_github.sh --push-google
```

### 2) push_local_to_github_and_google.sh
Commit **local edits** → push to **GitHub** → push to **Google**.  
Optional flag `--pull-first` grabs online edits before committing.

```bash
cd ~/Documents/"Genealogy Sheets"
chmod +x push_local_to_github_and_google.sh
./push_local_to_github_and_google.sh
./push_local_to_github_and_google.sh -m "Commit message"
./push_local_to_github_and_google.sh --pull-first
```

### 🛠 Troubleshooting
- If scripts fail with “missing clasp” → install Node 20+, then `npm install -g @google/clasp` and `clasp login`  
- If scripts fail with “not a git repo” → run `git init -b main` inside the folder  
- If scripts fail with “no origin remote” → add it:  
  ```bash
  git remote add origin https://github.com/<USER>/<REPO>.git
  ```
