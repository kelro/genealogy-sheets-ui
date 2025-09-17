# Genealogy Sheets UI (Apps Script)

A no‑Forms workflow for entering and exploring family data **directly in Google Sheets**.  
Includes sidebars for **Add Person / Add Marriage / Add Child**, a **Search & Edit** panel (with relink & delete), and a **D3-based Family Timeline**.

---

## 🚀 Highlights

- **Rename‑proof sheet access**  
  Resolves sheets by cached `sheetId` → name → header signature, so renames don’t break code.
- **Rich data entry UI**  
  Sidebars for adding People, Marriages, and Children with consistent styling (see `ui_common.html`).
- **Search & Edit panel**  
  Edit people, marriages, and children; **relink** a child to a different marriage (or none); delete child; **delete person with cascade & impact preview**.
- **Interactive Timeline (Dialog & Sidebar)**  
  Generation coloring, spouse/children details, **birth‑state callouts**, **dark mode**, zoom/brush, and **print**.
- **Production CSV Importer (with backups)**  
  One‑time / periodic import from a Drive CSV ID or URL with **automatic spreadsheet backup**, truncation & rewrite of People/Marriages/Children, and progress logging to `Import_Log`. (See `importer_Prod.gs`)
- **Date normalization & repair suite**  
  Tools to **normalize legacy dates to ISO** (`YYYY-MM-DD`) with **STRICT** or **APPROX** modes, produce **conversion reports**, **highlight issues**, and **force‑convert** ISO‑like text to real Dates. (See `NormalizeDates.gs` and GX* utilities in `Code.gs`)
- **Named Range utilities**  
  Export named ranges (`exportNamedRanges.gs`) and a **Named Range Renamer** in `Code.gs` that can preview/apply updates across formulas, validation rules, conditional formatting, filters, slicers, and (optionally) **Named Functions**.
- **Setup & repair utilities**  
  Create/verify tables, verify required HTML files, rebuild menus, install the **onOpen** trigger if needed.

---

## 📦 Repository Contents

```
Code.gs
ui_common.html
sidebar_person.html
sidebar_marriage.html
sidebar_child.html
sidebar_search.html
sidebar_bootstrap.html
index.html
importer_Prod.gs
NormalizeDates.gs
exportNamedRanges.gs
debugCellType.gs
```

**What each file does**

- **`Code.gs`** — Menus, rename‑proof helpers, table creation/verification, data APIs (search/edit/relink/delete), timeline API (`getPeopleData()`), maintenance & date utilities, **Named Range Renamer**, missing‑by‑surname finder, generation recalculation, and **import Missing → People**.
- **`index.html`** — D3 **Family Timeline** UI (dialog/sidebar and `doGet`), with filters, zoom/brush, dark mode, birth‑state callouts, and print.
- **`sidebar_person.html` / `sidebar_marriage.html` / `sidebar_child.html`** — Data‑entry sidebars with consistent UI and success/error toasts.
- **`sidebar_search.html`** — **Search & Edit** with person bundle view; edit fields; **relink child**; **delete child**; **delete person (cascade)** with impact preview.
- **`sidebar_bootstrap.html`** — Simple launcher for Search & Edit.
- **`ui_common.html`** — Shared CSS/UI components for consistent styling, inputs, buttons, statuses, and toasts.
- **`importer_Prod.gs`** — **Production CSV importer** with Drive file prompt, **automatic backup**, sheet truncation/rewrites, parsing for meta/spouses/children, and `Import_Log`.
- **`NormalizeDates.gs`** — **ISO date normalization** tools with `NORMALIZE_MODE = 'STRICT' | 'APPROX'` and a detailed `Normalization_Log`.
- **`exportNamedRanges.gs`** — Exports all named ranges to a new `NamedRangesExport` sheet.
- **`debugCellType.gs`** — Utility to quickly log value types (e.g., confirm Date objects).

---

## 🧭 Menus (as rendered by `onOpen()`)

**Genealogy**  
- 📋 **Data Entry**  
  - Add Person…  
  - Add Marriage…  
  - Add Child…  
  - Search & Edit…  
  - Timeline… (dialog)
- 🛠 **Utilities**  
  - Create/Verify Tables  
  - Verify Setup  
  - Find Missing by Surname…  
  - Recalculate Generations…  
  - Import Missing → People…
- ⚙️ **Maintenance**  
  - Rebuild Menu Now  
  - Repair Menu (install trigger)

> The menu uses section headers for readability; actions are grouped by task.

---

## 🔧 Setup

1. Open (or create) your Genealogy Google Sheet.  
2. **Extensions → Apps Script.**  
3. Add the files above to the project and paste contents.  
4. **Save** and reload the Sheet → **Genealogy** menu appears.

**Optional but recommended**  
- **Advanced Google Services → Google Sheets API** (toggle on)  
- **Google Cloud Console → enable “Google Sheets API”**  
  > Required if you use the **Named Range Renamer** with **Named Functions**.

---

## 🛠 Usage Guide

### 1) Create or verify base tables
- Run **Create/Verify Tables** once to create:
  - `People (headers: PersonID, FullName, DateOfBirth, PlaceOfBirth, DateOfDeath, PlaceOfDeath, Parents, Notes, Timestamp, Generation)`
  - `Marriages (headers: MarriageID, PersonID, SpouseName, MarriageDate, MarriagePlace, Status, Timestamp)`
  - `Children (headers: ChildID, PersonID, MarriageID, ChildName, BornDate, BornPlace, DiedDate, DiedPlace, Notes, Timestamp)`

### 2) Enter data via sidebars
- **Add Person / Add Marriage / Add Child** sidebars write rows with UUIDs and timestamps.  
- Child `MarriageID` is optional for out‑of‑wedlock cases.

### 3) Search & Edit bundle
- Search by **name or PersonID**.  
- View a **person bundle** (person + marriages + children).  
- Edit values inline; **relink child** between marriages (or none).  
- **Delete child**, or **delete person** with cascade (children → marriages → person) after an **impact preview**.

### 4) Interactive Timeline
- Launch **Timeline…** (dialog).  
- Filter by family, toggle **birth‑state callouts**, **dark mode**, zoom/brush to focus years, and **print**.

### 5) Production CSV Importer
- In Apps Script editor: run **`runProductionImportFromPrompt`**.  
- Paste a Drive **CSV URL or file ID** when prompted.  
- The importer will:
  - Make an **automatic backup** of the current spreadsheet.
  - Clear & **rewrite People / Marriages / Children**.
  - Parse **Born/Died meta**, spouses, and children lists.
  - Log progress to **`Import_Log`**.

### 6) Date normalization & repair (legacy data)
- **Normalize to ISO**: run `normalizeAllDates` (see `NORMALIZE_MODE` in `NormalizeDates.gs`).  
  - `STRICT` leaves year‑only as blank.  
  - `APPROX` converts `YYYY` → `YYYY-01-01` and logs the approximation.
- **Spot issues / force conversions** (from `Code.gs`):
  - `GX_showDateConversionIssues()` and `GX_showDateConversionIssues_REAL()` — report type issues.
  - `GX_highlightDateIssues()` / `GX_clearDateIssueHighlights()` — visual highlights.
  - `GX_forceConvertIsoDates()` — converts ISO‑like text to real Dates with formatting.
  - `GX_maintenanceDates()` — convenience wrapper for routine cleanup.
  - `debugCellType()` (separate file) to confirm whether a cell holds a **Date** vs **text**.

### 7) Named Range helpers
- **Named Range Renamer** (in `Code.gs`) can **preview & apply** renames across:
  - Formulas, data validation (custom formulas), **conditional formatting**, **filters**, **slicers**  
  - (Optional) **Named Functions** when Sheets API is enabled.
- **Export Named Ranges**: run `exportNamedRanges()` to generate a `NamedRangesExport` sheet.

### 8) Finding & importing missing People by surname
- **Find Missing by Surname…** creates a `Missing_<Surname>` sheet by comparing **Children** vs **People**.  
- **Import Missing → People…** (menu) or `GX_importMissingToPeople('<sheet>')` to bring them into **People** with auto `PersonID`.

### 9) Recalculating Generations
- **Recalculate Generations…** fills/updates the `Generation` column, inferring from parents when possible.

---

## 📑 Data Schema

**People**  
`PersonID`, `FullName`, `DateOfBirth`, `PlaceOfBirth`, `DateOfDeath`, `PlaceOfDeath`, `Parents`, `Notes`, `Timestamp`, `Generation` (optional)

**Marriages**  
`MarriageID`, `PersonID`, `SpouseName`, `MarriageDate`, `MarriagePlace`, `Status`, `Timestamp`

**Children**  
`ChildID`, `PersonID`, `MarriageID`, `ChildName`, `BornDate`, `BornPlace`, `DiedDate`, `DiedPlace`, `Notes`, `Timestamp`

**Notes**
- IDs are **UUIDs**.  
- Out‑of‑wedlock children have blank `MarriageID`.  
- `Generation` can be numeric or inferred from Notes (e.g., `Generation: N`).

---

## 🧰 Troubleshooting

- **Menus missing** → run **Repair Menu (install trigger)**, then reload the Sheet.  
- **Missing HTML warning** → run **Verify Setup** to see which files are absent.  
- **Headers missing** → run **Create/Verify Tables**.  
- **Renamer + Named Functions** → enable **Sheets API** as noted above.  
- **Type vs display confusion** → use `debugCellType()` to log underlying types; use `GX_forceConvertIsoDates()` to coerce ISO‑like text to real Dates.

---

## 🔁 Changelog (summary of recent additions)

- **Search & Edit** with **child relink** and **delete**; **delete person (cascade)** with impact preview.  
- **Timeline dialog** with generation colors, spouse/children details, birth‑state callouts, dark mode, zoom/brush, print.  
- **Rename‑proof resolvers** and setup utilities.  
- **Named Range Renamer** (preview/apply across formulas, validation, conditional formatting, filters, slicers; optional Named Functions).  
- **Production CSV Importer** with backup & logging.  
- **Date normalization suite** + conversion reports, highlights, and force‑convert helpers.

---

## 📜 License

MIT (or your preferred license).

---

## ☁️ Deploying as a Web App (optional)

- Add a `doGet(e)` that serves `index.html` (timeline) if you want a read‑only web view.  
- Publish → Deploy as web app → Execute as **Me**, Who has access **Anyone with the link** (or your domain).

---

## 🧩 GitHub tips

- Include this `README.md` in the repo root.  
- Consider a `.gitignore` (see below).  
- Tag releases after major schema or importer changes.

**Suggested `.gitignore`**

```
# macOS
.DS_Store

# Node (if you later add build tooling)
node_modules/
dist/

# Apps Script clasp (if adopted later)
.clasp.json
appsscript.json
```

---

## 🛠 Contributing & Development

This project uses GitHub CLI (`gh`) and Git for version control.  
If you need to push changes or update the repo, see the step-by-step guide:

➡️ [PUSH_TO_GITHUB.md](PUSH_TO_GITHUB.md)

That document covers:
- Creating the repo from the command line (no GitHub website required)
- Making the first commit and initial push
- Everyday updates (edit/add/delete files)
- Branching, pull requests, tags/releases
- Common fixes for Git errors

### 🔄 Google Apps Script Sync (clasp)

If you want to sync your Apps Script project (from Google Sheets) with this repo,  
see the **clasp workflow** section in [PUSH_TO_GITHUB.md](PUSH_TO_GITHUB.md).  
This explains how to:

- Install and log in to `clasp`
- Pull code down from Google
- Push local edits back to Google
- Keep `appsscript.json` versioned in GitHub

