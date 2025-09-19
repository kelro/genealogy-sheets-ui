/**
 * Genealogy Sheets UI — Robust + Child Relinking & Delete + Named Range Renamer merged
 * (Rename-proof: resolves sheets by cached sheetId → name → header signature)
 */

/** === CONFIG === */
const SHEET_PEOPLE   = 'People';
const SHEET_MARRIAGE = 'Marriages';
const SHEET_CHILDREN = 'Children';

const HEADERS_PEOPLE   = ['PersonID','FullName','DateOfBirth','PlaceOfBirth','DateOfDeath','PlaceOfDeath','Parents','Notes','Timestamp'];
const HEADERS_MARRIAGE = ['MarriageID','PersonID','SpouseName','MarriageDate','MarriagePlace','Status','Timestamp'];
const HEADERS_CHILDREN = ['ChildID','PersonID','MarriageID','ChildName','BornDate','BornPlace','DiedDate','DiedPlace','Notes','Timestamp'];

/** =====================================================================================
 *  MENUS (Merged + Robust)
 * ===================================================================================== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    // --- Genealogy menu ---
    ui.createMenu('Genealogy')
      // Section: Data Entry
      .addItem("📋 Data Entry", "GX_menuHeader")
      .addItem("Add Person…", "showAddPerson")
      .addItem("Add Marriage…", "showAddMarriage")
      .addItem("Add Child…", "showAddChild")
      .addItem("Search & Edit…", "showSearchEdit")
      .addItem("Timeline…", "showTimelineDialog")

      // Section: Utilities
      .addItem("🛠 Utilities", "GX_menuHeader")
      .addItem("Create/Verify Tables", "ensureAllTables")
      .addItem("Verify Setup", "GX_verifySetup")
      .addItem("Find Missing by Surname…", "findMissingAnySurname")
      .addItem("Recalculate Generations…", "recalcGenerationsPrompt")
      .addItem("Import Missing → People…", "GX_importMissingPrompt")

      // Section: Maintenance
      .addItem("⚙️ Maintenance", "GX_menuHeader")
      .addItem("Rebuild Menu Now", "GX_rebuildMenuNow")
      .addItem("Repair Menu (install trigger)", "GX_installOpenTrigger")

      .addToUi();

    // --- Named Range Renamer menu ---
    ui.createMenu('Named Range Renamer')
      .addItem('Preview changes…', 'NRR_preview')
      .addItem('Apply changes…', 'NRR_apply')
      .addToUi();
  } catch (e) {
    try {
      ui.createMenu('Genealogy')
        .addItem('Repair Menu (install trigger)', 'GX_installOpenTrigger')
        .addItem('Verify Setup', 'GX_verifySetup')
        .addToUi();
    } catch (_) {}
    console.error('onOpen failed:', e);
  }
}

/** Force-rebuild the menus immediately (no reload needed). */
function GX_rebuildMenuNow() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menus rebuilt.\nIf they still do not appear after reload, run "Repair Menu (install trigger)".');
}

/** Install an installable onOpen trigger for this spreadsheet. */
function GX_installOpenTrigger() {
  const ssId = SpreadsheetApp.getActive().getId();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onOpenHandler_' || t.getHandlerFunction() === 'onOpen')
    .forEach(ScriptApp.deleteTrigger);

  ScriptApp.newTrigger('onOpenHandler_')
    .forSpreadsheet(ssId)
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('Installable onOpen trigger created.\nClose & reopen the spreadsheet to test.');
}

/** Installable-trigger entry that safely calls onOpen. */
function onOpenHandler_() {
  try { onOpen(); } catch (e) { console.error('onOpenHandler_ failed:', e); }
}

/** Check expected sheets + HTML files exist; create/fix sheets/headers. */
function GX_verifySetup() {
  const missingBits = [];
  try {
    ensureAllTables();
  } catch (e) {
    missingBits.push('⚠️ ensureAllTables error: ' + (e && e.message));
  }

  const htmlFiles = ['sidebar_person','sidebar_marriage','sidebar_child','sidebar_bootstrap','sidebar_search','index'];
  const missingHtml = [];
  htmlFiles.forEach(name => {
    try { HtmlService.createTemplateFromFile(name); }
    catch (e) { missingHtml.push(name); }
  });

  const msg = [
    'Setup verification complete.',
    missingBits.length ? missingBits.join('\\n') : 'Sheets/headers OK.',
    missingHtml.length ? ('Missing HTML files: ' + missingHtml.join(', ')) : 'All HTML files present.'
  ].join('\\n');

  SpreadsheetApp.getUi().alert(msg);
}

/** =====================================================================================
 *  RENAME-PROOF SHEET RESOLVERS
 * ===================================================================================== */

const SP_ = PropertiesService.getScriptProperties();

function peopleSheet_()   { return resolveSheet_('people',    SHEET_PEOPLE,   HEADERS_PEOPLE); }
function marriageSheet_() { return resolveSheet_('marriages', SHEET_MARRIAGE, HEADERS_MARRIAGE); }
function childrenSheet_() { return resolveSheet_('children',  SHEET_CHILDREN, HEADERS_CHILDREN); }

function resolveSheet_(key, defaultName, headers) {
  const ss = SpreadsheetApp.getActive();

  const cached = Number(SP_.getProperty('sheetId:' + key)) || null;
  if (cached) {
    const byId = ss.getSheets().find(s => s.getSheetId() === cached);
    if (byId) {
      ensureHeaders_(byId, headers);
      return byId;
    }
  }

  let sh = ss.getSheetByName(defaultName);

  if (!sh) {
    sh = ss.getSheets().find(s => headersMatch_(s, headers));
  }

  if (!sh) {
    sh = ss.insertSheet(defaultName);
    sh.appendRow(headers);
  } else {
    ensureHeaders_(sh, headers);
  }

  SP_.setProperty('sheetId:' + key, String(sh.getSheetId()));
  return sh;
}

function headersMatch_(sh, headers) {
  if (sh.getLastRow() === 0) return false;
  const row = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
                .getDisplayValues()[0]
                .map(v => String(v).trim());
  return headers.every(h => row.indexOf(h) !== -1);
}

function ensureHeaders_(sh, headers) {
  const current = (sh.getLastRow() === 0)
    ? []
    : sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
        .getDisplayValues()[0]
        .map(v => String(v).trim());
  const missing = headers.filter(h => current.indexOf(h) === -1);
  if (missing.length) {
    sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
    sh.getRange(1, sh.getLastColumn() - missing.length + 1, 1, missing.length).setValues([missing]);
  }
}

/** =====================================================================================
 *  UI LAUNCHERS
 * ===================================================================================== */
function showAddPerson() {
  ensureAllTables();
  const html = HtmlService.createTemplateFromFile('sidebar_person').evaluate().setTitle('Add Person').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showAddMarriage() {
  ensureAllTables();
  const html = HtmlService.createTemplateFromFile('sidebar_marriage').evaluate().setTitle('Add Marriage').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showAddChild() {
  ensureAllTables();
  const html = HtmlService.createTemplateFromFile('sidebar_child').evaluate().setTitle('Add Child').setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showSearchEdit() {
  ensureAllTables();
  const html = HtmlService.createTemplateFromFile('sidebar_bootstrap')
    .evaluate()
    .setTitle('Search & Edit')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}
function openSearchEditAt(widthPx) {
  ensureAllTables();
  const w = Math.max(560, Math.min(900, Math.floor(Number(widthPx) || 670)));
  const html = HtmlService.createTemplateFromFile('sidebar_search')
    .evaluate()
    .setTitle('Search & Edit')
    .setWidth(w);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Timeline launcher (dialog only) */
function showTimelineDialog() {
  const t = HtmlService.createTemplateFromFile('index');
  t.family = 'All';
  const html = t.evaluate().setTitle('Timeline').setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Family Timeline');
}

/** =====================================================================================
 *  TABLE MANAGEMENT
 * ===================================================================================== */
function ensureAllTables() {
  peopleSheet_();
  marriageSheet_();
  childrenSheet_();
}

function ensureSheetWithHeaders(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else {
    const current = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0] || [];
    const missing = headers.filter(h => current.indexOf(h) === -1);
    if (missing.length) {
      sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
      sh.getRange(1, sh.getLastColumn() - missing.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sh;
}

/** =====================================================================================
 *  UTILITIES
 * ===================================================================================== */
function uuid() {
  const chars = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.split('');
  for (let i=0; i<chars.length; i++) {
    const c = chars[i];
    if (c === 'x' || c === 'y') {
      const r = Math.random() * 16 | 0;
      const v = (c === 'x') ? r : (r & 0x3 | 0x8);
      chars[i] = v.toString(16);
    }
  }
  return chars.join('');
}
function nowString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Chicago', "yyyy-MM-dd'T'HH:mm:ss");
}
function getHeaders_(sh) {
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] || [];
  return hdrs.map(h => String(h).trim());
}
function headerIndex_(sh, headerName) {
  const hdrs = getHeaders_(sh);
  const want = String(headerName).trim().toLowerCase();
  return hdrs.findIndex(h => String(h).trim().toLowerCase() === want);
}
function findRowById_(sh, idHeader, idValue) {
  const idIdx = headerIndex_(sh, idHeader);
  if (idIdx === -1) return null;
  const nRows = Math.max(0, sh.getLastRow() - 1);
  if (nRows === 0) return null;
  const col = sh.getRange(2, idIdx + 1, nRows, 1).getDisplayValues();
  const needle = String(idValue).trim().toLowerCase();
  for (let i=0;i<col.length;i++) {
    const val = String(col[i][0]).trim().toLowerCase();
    if (val === needle) return i + 2; // sheet row
  }
  return null;
}
function findRowsWhere_(sh, header, value) {
  const idx = headerIndex_(sh, header);
  if (idx === -1) return [];
  const nRows = Math.max(0, sh.getLastRow() - 1);
  if (nRows === 0) return [];
  const col = sh.getRange(2, idx + 1, nRows, 1).getDisplayValues();
  const needle = String(value).trim().toLowerCase();
  const rows = [];
  for (let i=0;i<col.length;i++) {
    const val = String(col[i][0]).trim().toLowerCase();
    if (val === needle) rows.push(i + 2);
  }
  return rows;
}
function rowToObject_(sh, rowNumber) {
  const hdrs = getHeaders_(sh);
  const values = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getDisplayValues()[0] || [];
  const obj = {};
  for (let i=0;i<hdrs.length;i++) obj[hdrs[i]] = values[i];
  return obj;
}
function writeColumns_(sh, rowNumber, _headers, patch) {
  Object.keys(patch).forEach(key => {
    const idx = headerIndex_(sh, key);
    if (idx !== -1) sh.getRange(rowNumber, idx + 1).setValue(patch[key]);
  });
}
function getValue_(sh, rowNumber, header) {
  const idx = headerIndex_(sh, header);
  if (idx === -1) return '';
  return sh.getRange(rowNumber, idx+1).getDisplayValue();
}

/** === Date helpers (safe, local, no UTC shift) === */
function makeLocalDate_(y, m, d) {
  // y=full year (e.g., 1980), m=1-12, d=1-31
  if (!(y && m && d)) return null;
  const dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || (dt.getMonth()+1) !== m || dt.getDate() !== d) return null;
  return dt;
}

function parseYMD_(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  const m = String(s).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return makeLocalDate_(+m[1], +m[2], +m[3]);
}

function parseDateTimeLoose_(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  const d = new Date(String(s));
  return isNaN(+d) ? null : d;
}

function nowDate_() { return new Date(); }

const DATE_HEADERS_BY_SHEET = {
  People:    ['DateOfBirth','DateOfDeath'],
  Marriages: ['MarriageDate'],
  Children:  ['BornDate','DiedDate'],
};
const DATETIME_HEADERS_BY_SHEET = {
  People:    ['Timestamp'],
  Marriages: ['Timestamp'],
  Children:  ['Timestamp'],
};

function formatRowDates_(sheet, rowNumber) {
  const name = sheet.getName();
  const dateHeaders = DATE_HEADERS_BY_SHEET[name] || [];
  const dtHeaders   = DATETIME_HEADERS_BY_SHEET[name] || [];
  dateHeaders.forEach(h => {
    const col = headerIndex_(sheet, h);
    if (col !== -1) sheet.getRange(rowNumber, col+1).setNumberFormat('yyyy-mm-dd');
  });
  dtHeaders.forEach(h => {
    const col = headerIndex_(sheet, h);
    if (col !== -1) sheet.getRange(rowNumber, col+1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
}

/** =====================================================================================
 *  READ API (dropdown data)
 * ===================================================================================== */
function listPeople() {
  const sh = peopleSheet_();
  const idIdx = headerIndex_(sh, 'PersonID');
  const nameIdx = headerIndex_(sh, 'FullName');
  const nRows = Math.max(0, sh.getLastRow()-1);
  const data = nRows ? sh.getRange(2, 1, nRows, sh.getLastColumn()).getDisplayValues() : [];
  const people = data
    .filter(r => (r[idIdx-0] && r[nameIdx-0]))
    .map(r => ({ id: String(r[idIdx]), name: String(r[nameIdx]) }))
    .sort((a,b) => a.name.localeCompare(b.name));
  return people;
}
function listMarriages(personId) {
  const sh = marriageSheet_();
  const rows = findRowsWhere_(sh, 'PersonID', personId).map(r => rowToObject_(sh, r));
  const items = rows.map(m => ({ id: String(m['MarriageID']), label: String(m['SpouseName'] || '(Unnamed spouse)') }));
  items.unshift({ id: '', label: '— Out-of-wedlock / Not linked to a marriage —' });
  return items;
}

/** =====================================================================================
 *  CREATE API (now writes real Date objects for dates/timestamps)
 * ===================================================================================== */
function addPerson(payload) {
  ensureAllTables();
  const sh = peopleSheet_();
  const id = uuid();

  const dob = parseYMD_(payload.dob);
  const dod = parseYMD_(payload.dod);

  sh.appendRow([
    id, (payload.fullName || '').trim(),
    dob || '', (payload.pob || '').trim(),
    dod || '', (payload.pod || '').trim(),
    (payload.parents || '').trim(), (payload.notes || '').trim(),
    nowDate_()
  ]);

  const row = sh.getLastRow();
  formatRowDates_(sh, row);

  return { ok: true, personId: id };
}
function addMarriage(payload) {
  ensureAllTables();
  const sh = marriageSheet_();
  const id = uuid();

  const mDate = parseYMD_(payload.mDate);

  sh.appendRow([
    id, (payload.personId || '').trim(),
    (payload.spouseName || '').trim(),
    mDate || '', (payload.mPlace || '').trim(),
    (payload.status || '').trim(), nowDate_()
  ]);

  const row = sh.getLastRow();
  formatRowDates_(sh, row);

  return { ok: true, marriageId: id };
}
function addChild(payload) {
  ensureAllTables();
  const sh = childrenSheet_();
  const id = uuid();

  const bDate = parseYMD_(payload.bDate);
  const dDate = parseYMD_(payload.dDate);

  sh.appendRow([
    id, (payload.personId || '').trim(),
    (payload.marriageId || '').trim(), // empty = out-of-wedlock
    (payload.childName || '').trim(),
    bDate || '', (payload.bPlace || '').trim(),
    dDate || '', (payload.dPlace || '').trim(),
    (payload.notes || '').trim(), nowDate_()
  ]);

  const row = sh.getLastRow();
  formatRowDates_(sh, row);

  return { ok: true, childId: id };
}

/** =====================================================================================
 *  SEARCH & EDIT API
 * ===================================================================================== */
function searchPeople(query) {
  ensureAllTables();
  const q = String(query || '').toLowerCase().trim();
  const sh = peopleSheet_();
  const nRows = Math.max(0, sh.getLastRow()-1);
  const data = nRows ? sh.getRange(2, 1, nRows, sh.getLastColumn()).getDisplayValues() : [];
  const idx = {
    id: headerIndex_(sh, 'PersonID'),
    name: headerIndex_(sh, 'FullName'),
    dob: headerIndex_(sh, 'DateOfBirth'),
    pob: headerIndex_(sh, 'PlaceOfBirth'),
    dod: headerIndex_(sh, 'DateOfDeath'),
    pod: headerIndex_(sh, 'PlaceOfDeath'),
    parents: headerIndex_(sh, 'Parents'),
    notes: headerIndex_(sh, 'Notes'),
  };
  const results = data
    .filter(r => !q || String(r[idx.name]).toLowerCase().includes(q) || String(r[idx.id]).toLowerCase().includes(q))
    .slice(0, 100)
    .map(r => ({
      personId: String(r[idx.id]), fullName: String(r[idx.name] || ''),
      dob: String(r[idx.dob] || ''), pob: String(r[idx.pob] || ''),
      dod: String(r[idx.dod] || ''), pod: String(r[idx.pod] || ''),
      parents: String(r[idx.parents] || ''), notes: String(r[idx.notes] || '')
    }));
  return results;
}
function getPersonBundle(personId) {
  ensureAllTables();
  const psh = peopleSheet_();
  const prow = findRowById_(psh, 'PersonID', personId);
  const person = prow ? rowToObject_(psh, prow) : null;

  const msh = marriageSheet_();
  const mrows = findRowsWhere_(msh, 'PersonID', personId).map(r => rowToObject_(msh, r));

  const csh = childrenSheet_();
  const crows = findRowsWhere_(csh, 'PersonID', personId).map(r => rowToObject_(csh, r));

  return { person, marriages: mrows, children: crows };
}

/** =====================================================================================
 *  UPDATE + RELINK + DELETE (updates now write real Date objects)
 * ===================================================================================== */
function updatePerson(payload) {
  const sh = peopleSheet_();
  const row = findRowById_(sh, 'PersonID', payload.personId);
  if (!row) throw new Error('Person not found');
  const patch = {
    FullName: payload.fullName || getValue_(sh, row, 'FullName'),
    DateOfBirth: parseYMD_(payload.dob) || '',
    PlaceOfBirth: (payload.pob || '').trim(),
    DateOfDeath: parseYMD_(payload.dod) || '',
    PlaceOfDeath: (payload.pod || '').trim(),
    Parents: (payload.parents || '').trim(),
    Notes: (payload.notes || '').trim()
  };
  writeColumns_(sh, row, HEADERS_PEOPLE, patch);
  formatRowDates_(sh, row);
  return { ok: true };
}
function updateMarriage(payload) {
  const sh = marriageSheet_();
  const row = findRowById_(sh, 'MarriageID', payload.marriageId);
  if (!row) throw new Error('Marriage not found');
  const patch = {
    SpouseName: (payload.spouseName || '').trim(),
    MarriageDate: parseYMD_(payload.mDate) || '',
    MarriagePlace: (payload.mPlace || '').trim(),
    Status: (payload.status || '').trim()
  };
  writeColumns_(sh, row, HEADERS_MARRIAGE, patch);
  formatRowDates_(sh, row);
  return { ok: true };
}
function updateChild(payload) {
  const sh = childrenSheet_();
  const row = findRowById_(sh, 'ChildID', payload.childId);
  if (!row) throw new Error('Child not found');
  const patch = {
    MarriageID: (payload.marriageId || '').trim(),
    ChildName: (payload.childName || '').trim(),
    BornDate: parseYMD_(payload.bDate) || '',
    BornPlace: (payload.bPlace || '').trim(),
    DiedDate: parseYMD_(payload.dDate) || '',
    DiedPlace: (payload.dPlace || '').trim(),
    Notes: (payload.notes || '').trim()
  };
  writeColumns_(sh, row, HEADERS_CHILDREN, patch);
  formatRowDates_(sh, row);
  return { ok: true };
}
function relinkChild(childId, newMarriageId) {
  const sh = childrenSheet_();
  const row = findRowById_(sh, 'ChildID', childId);
  if (!row) throw new Error('Child not found');
  writeColumns_(sh, row, HEADERS_CHILDREN, { MarriageID: (newMarriageId || '').trim() });
  return { ok: true };
}
function deleteChild(childId) {
  const sh = childrenSheet_();
  const row = findRowById_(sh, 'ChildID', childId);
  if (!row) throw new Error('Child not found');
  sh.deleteRow(row);
  return { ok: true };
}

/** Serve shared styles (for HTML files) */
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/** =====================================================================================
 *  NAMED RANGE RENAMER (Preview / Apply)
 * ===================================================================================== */

function NRR_preview() {
  const ui = SpreadsheetApp.getUi();
  const oldName = prompt_(ui, 'Enter the OLD named range (exact):');
  if (!oldName) return;
  const newName = prompt_(ui, 'Enter the NEW named range (exact):');
  if (!newName) return;

  const flags = getFlagsFromUser_(ui);
  const report = runAll_(oldName, newName, { dryRun: true, ...flags });
  ui.alert('Preview complete', report, ui.ButtonSet.OK);
}

function NRR_apply() {
  const ui = SpreadsheetApp.getUi();
  const oldName = prompt_(ui, 'Enter the OLD named range (exact):');
  if (!oldName) return;
  const newName = prompt_(ui, 'Enter the NEW named range (exact):');
  if (!newName) return;

  const proceed = ui.alert(
    'Confirm apply',
    'This will update formulas, rules, filters, and more across ALL sheets.\\n' +
    'Consider File → Make a copy first. You can Undo afterward.\\n\\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (proceed !== ui.Button.YES) return;

  const flags = getFlagsFromUser_(ui);
  let report = runAll_(oldName, newName, { dryRun: false, ...flags });

  const rename = ui.alert(
    'Rename Named Range object too?',
    `If a Named Range named "${oldName}" exists, rename it to "${newName}"?`,
    ui.ButtonSet.YES_NO
  );
  if (rename === ui.Button.YES) {
    report += '\\n\\n' + renameNamedRangeObject_(oldName, newName);
  }

  ui.alert('Apply complete', report, ui.ButtonSet.OK);
}

/** Optional flag prompt (Named Functions). */
function getFlagsFromUser_(ui) {
  const resp = ui.alert(
    'Optional: Update Named Functions?',
    'If you created Google Sheets Named Functions, we can try updating their text too (requires Advanced Sheets Service). ' +
    'Choose "Yes" to include them; "No" to skip.',
    ui.ButtonSet.YES_NO
  );
  return { updateNamedFunctions: resp === ui.Button.YES };
}

/** Core pass across sheets. */
function runAll_(oldName, newName, opts) {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const boundary = '[A-Z0-9_]';
  const re = new RegExp('(^|[^' + boundary + '])' + escapeForRegex_(oldName) + '(?=$|[^' + boundary + '])', 'g');

  let totals = { formulaCells: 0, dataValidation: 0, condFormatting: 0, filters: 0, slicers: 0, namedFunctions: 0 };
  const perSheetLines = [];

  sheets.forEach(sheet => {
    let sheetSum = { formulaCells: 0, dataValidation: 0, condFormatting: 0, filters: 0, slicers: 0 };

    // 1) Cells (formulas)
    {
      const range = sheet.getDataRange();
      const formulas = range.getFormulas();
      let changed = 0;
      const newFormulas = formulas.map(row => row.slice());
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const f = formulas[r][c];
          if (!f) continue;
          const rep = f.replace(re, (_, left) => left + newName);
          if (rep !== f) {
            changed++;
            if (!opts.dryRun) range.getCell(r+1,c+1).setFormula(rep);
          }
        }
      }
      totals.formulaCells += changed;
      sheetSum.formulaCells += changed;
    }

    // 2) Data validation (custom formulas)
    {
      const range = sheet.getDataRange();
      const rules2D = range.getDataValidations();
      if (rules2D && rules2D.length) {
        let changed = 0;
        const newRules2D = rules2D.map(row => row ? row.slice() : row);
        for (let r = 0; r < rules2D.length; r++) {
          const row = rules2D[r];
          if (!row) continue;
          for (let c = 0; c < row.length; c++) {
            const rule = row[c];
            if (!rule) continue;
            const type = rule.getCriteriaType && rule.getCriteriaType();
            if (type && String(type).indexOf('CUSTOM_FORMULA') !== -1) {
              const vals = rule.getCriteriaValues();
              const oldFormula = vals && vals[0];
              if (typeof oldFormula === 'string' && oldFormula) {
                const rep = oldFormula.replace(re, (_, left) => left + newName);
                if (rep !== oldFormula) {
                  changed++;
                  if (!opts.dryRun) {
                    const b = rule.copy();
                    b.requireFormulaSatisfied(rep);
                    newRules2D[r][c] = b.build();
                  }
                }
              }
            }
          }
        }
        if (!opts.dryRun && changed > 0) range.setDataValidations(newRules2D);
        totals.dataValidation += changed;
        sheetSum.dataValidation += changed;
      }
    }

    // 3) Conditional formatting (custom formulas)
    {
      const rules = sheet.getConditionalFormatRules();
      let changed = 0;
      const out = [];
      rules.forEach(rule => {
        const bc = rule.getBooleanCondition && rule.getBooleanCondition();
        if (bc && bc.getCriteriaType && String(bc.getCriteriaType()).indexOf('CUSTOM_FORMULA') !== -1) {
          const vals = bc.getCriteriaValues();
          let replacedAny = false;
          const newVals = (vals || []).map(v => {
            if (typeof v === 'string') {
              const rep = v.replace(re, (_, left) => left + newName);
              if (rep !== v) replacedAny = true;
              return rep;
            }
            return v;
          });
          if (replacedAny) {
            changed++;
            if (!opts.dryRun) {
              const b = rule.copy();
              b.withCriteria(bc.getCriteriaType(), newVals);
              out.push(b.build());
              return;
            }
          }
        }
        out.push(rule);
      });
      if (!opts.dryRun && changed > 0) sheet.setConditionalFormatRules(out);
      totals.condFormatting += changed;
      sheetSum.condFormatting += changed;
    }

    // 4) Filters & filter views (custom formulas)
    {
      let changed = 0;

      const filter = sheet.getFilter && sheet.getFilter();
      if (filter) {
        const lastColumn = sheet.getLastColumn();
        for (let col = 1; col <= lastColumn; col++) {
          const fc = filter.getColumnFilterCriteria(col);
          if (!fc) continue;
          if (fc.getCriteriaType && String(fc.getCriteriaType()).indexOf('CUSTOM_FORMULA') !== -1) {
            const vals = fc.getCriteriaValues();
            const oldFormula = vals && vals[0];
            if (typeof oldFormula === 'string' && oldFormula) {
              const rep = oldFormula.replace(re, (_, left) => left + newName);
              if (rep !== oldFormula) {
                changed++;
                if (!opts.dryRun) {
                  const b = fc.copy();
                  b.whenFormulaSatisfied(rep);
                  filter.setColumnFilterCriteria(col, b.build());
                }
              }
            }
          }
        }
      }

      const views = sheet.getFilterViews && sheet.getFilterViews();
      if (views && views.length) {
        views.forEach(view => {
          const range = view.getRange();
          const cols = range.getNumColumns();
          for (let i = 1; i <= cols; i++) {
            const col = range.getColumn() + i - 1;
            const fc = view.getColumnFilterCriteria(col);
            if (!fc) continue;
            if (fc.getCriteriaType && String(fc.getCriteriaType()).indexOf('CUSTOM_FORMULA') !== -1) {
              const vals = fc.getCriteriaValues();
              const oldFormula = vals && vals[0];
              if (typeof oldFormula === 'string' && oldFormula) {
                const rep = oldFormula.replace(re, (_, left) => left + newName);
                if (rep !== oldFormula) {
                  changed++;
                  if (!opts.dryRun) {
                    const b = fc.copy();
                    b.whenFormulaSatisfied(rep);
                    view.setColumnFilterCriteria(col, b.build());
                  }
                }
              }
            }
          }
        });
      }

      totals.filters += changed;
      sheetSum.filters += changed;
    }

    // 5) Slicers (custom formula criteria)
    {
      let changed = 0;
      const slicers = sheet.getSlicers && sheet.getSlicers();
      if (slicers && slicers.length) {
        slicers.forEach(slicer => {
          const crit = slicer.getColumnFilterCriteria();
          if (!crit) return;
          if (crit.getCriteriaType && String(crit.getCriteriaType()).indexOf('CUSTOM_FORMULA') !== -1) {
            const vals = crit.getCriteriaValues();
            const oldFormula = vals && vals[0];
            if (typeof oldFormula === 'string' && oldFormula) {
              const rep = oldFormula.replace(re, (_, left) => left + newName);
              if (rep !== oldFormula) {
                changed++;
                if (!opts.dryRun) {
                  const b = crit.copy();
                  b.whenFormulaSatisfied(rep);
                  slicer.setColumnFilterCriteria(b.build());
                }
              }
            }
          }
        });
      }
      totals.slicers += changed;
      sheetSum.slicers += changed;
    }

    perSheetLines.push(
      `${sheet.getName()}: cells ${sheetSum.formulaCells}, dataVal ${sheetSum.dataValidation}, condFmt ${sheetSum.condFormatting}, filters ${sheetSum.filters}, slicers ${sheetSum.slicers}`
    );
  });

  // 6) Named Functions (optional; Advanced Sheets Service must be enabled)
  if (opts.updateNamedFunctions) {
    try {
      totals.namedFunctions += replaceInNamedFunctions_(oldName, newName, re, opts.dryRun);
    } catch (e) {
      perSheetLines.push(`Named Functions: ⚠️ Skipped (error: ${e && e.message ? e.message : e})`);
    }
  }

  const header = opts.dryRun ? 'PREVIEW (no changes applied)' : 'APPLY (changes applied)';
  return [
    header,
    `Old name: ${oldName}`,
    `New name: ${newName}`,
    '',
    'Per-sheet updates:',
    ...perSheetLines,
    '',
    'Totals:',
    `• Cell formulas changed: ${totals.formulaCells}`,
    `• Data validation (custom formulas) changed: ${totals.dataValidation}`,
    `• Conditional formatting (custom formulas) changed: ${totals.condFormatting}`,
    `• Filters & filter views (custom formulas) changed: ${totals.filters}`,
    `• Slicers (custom formulas) changed: ${totals.slicers}`,
    `• Named Functions updated: ${totals.namedFunctions}`,
  ].join('\\n');
}

/** Rename the Named Range object itself, if present. */
function renameNamedRangeObject_(oldName, newName) {
  const ss = SpreadsheetApp.getActive();
  const ranges = ss.getNamedRanges();
  const conflict = ranges.find(nr => nr.getName() === newName);
  if (conflict) return `⚠️ Skipped renaming Named Range object: "${newName}" already exists.`;
  const target = ranges.find(nr => nr.getName() === oldName);
  if (!target) return `ℹ️ No Named Range object named "${oldName}" found.`;
  target.setName(newName);
  return `✅ Renamed Named Range object "${oldName}" → "${newName}".`;
}

/** Named Functions updater (Advanced Sheets Service). */
function replaceInNamedFunctions_(oldName, newName, re, dryRun) {
  const ss = SpreadsheetApp.getActive();
  const ssId = ss.getId();
  const meta = Sheets.Spreadsheets.get(ssId, { fields: 'namedFunctions' });
  const nfs = (meta.namedFunctions || []);
  let changed = 0;
  const requests = [];

  nfs.forEach(nf => {
    const body = nf.functionBody || '';
    const rep = body.replace(re, (_, left) => left + newName);
    if (rep !== body) {
      changed++;
      if (!dryRun) {
        requests.push({
          updateNamedFunction: {
            namedFunction: {
              name: nf.name,
              displayName: nf.displayName,
              parameters: nf.parameters || [],
              functionBody: rep
            },
            fields: 'functionBody',
          }
        });
      }
    }
  });

  if (!dryRun && requests.length) {
    Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
  }
  return changed;
}

/** Small UI helpers for prompts + regex escape */
function prompt_(ui, message) {
  const res = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  return res.getResponseText().trim();
}
function escapeForRegex_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** =========================
 * Timeline backend (Apps Script)
 * ========================= */

function doGet(e) {
  var family = (e && e.parameter && e.parameter.family) || 'All';
  var t = HtmlService.createTemplateFromFile('index');
  t.family = family;
  return t.evaluate()
    .setTitle('Family Timeline')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Data API consumed by index.html */
function getPeopleData(_family) {
  const ssTZ = Session.getScriptTimeZone() || 'America/Chicago';
  const currentYear = Number(Utilities.formatDate(new Date(), ssTZ, 'yyyy'));

  // Robust resolvers (rename-proof)
  const psh = peopleSheet_();
  const msh = marriageSheet_();
  const csh = childrenSheet_();

  const P = (k) => headerIndex_(psh, k);
  const M = (k) => headerIndex_(msh, k);
  const C = (k) => headerIndex_(csh, k);

  const pIdx = {
    PersonID: P('PersonID'),
    FullName: P('FullName'),
    DateOfBirth: P('DateOfBirth'),
    PlaceOfBirth: P('PlaceOfBirth'),
    DateOfDeath: P('DateOfDeath'),
    PlaceOfDeath: P('PlaceOfDeath'),
    Parents: P('Parents'),
    Notes: P('Notes'),
    Generation: P('Generation')
  };
  const mIdx = {
    MarriageID: M('MarriageID'),
    PersonID:   M('PersonID'),
    SpouseName: M('SpouseName'),
    MarriageDate:  M('MarriageDate'),
    MarriagePlace: M('MarriagePlace'),
    Status:     M('Status')
  };
  const cIdx = {
    ChildID:   C('ChildID'),
    PersonID:  C('PersonID'),
    MarriageID:C('MarriageID'),
    ChildName: C('ChildName'),
    BornDate:  C('BornDate'),
    BornPlace: C('BornPlace'),
    DiedDate:  C('DiedDate'),
    DiedPlace: C('DiedPlace'),
    Notes:     C('Notes')
  };

  const pVals = psh.getLastRow() > 1 ? psh.getRange(2, 1, psh.getLastRow()-1, psh.getLastColumn()).getDisplayValues() : [];
  const mVals = msh.getLastRow() > 1 ? msh.getRange(2, 1, msh.getLastRow()-1, msh.getLastColumn()).getDisplayValues() : [];
  const cVals = csh.getLastRow() > 1 ? csh.getRange(2, 1, csh.getLastRow()-1, csh.getLastColumn()).getDisplayValues() : [];

  const marriagesByPID = {};
  mVals.forEach(r => {
    const pid = TL_clean(r[mIdx.PersonID]);
    if (pid) (marriagesByPID[pid] = marriagesByPID[pid] || []).push(r);
  });
  const childrenByPID = {};
  cVals.forEach(r => {
    const pid = TL_clean(r[cIdx.PersonID]);
    if (pid) (childrenByPID[pid] = childrenByPID[pid] || []).push(r);
  });

  const people = pVals.map(r => {
    const pid = TL_clean(r[pIdx.PersonID]);
    theFullName = TL_clean(r[pIdx.FullName]); // keep var name stable for readability
    const fullName = theFullName;

    const by = TL_yearFrom(r[pIdx.DateOfBirth]);
    const dy = TL_yearFrom(r[pIdx.DateOfDeath]);

    const ageAtDeath = (by != null && dy != null && dy >= by) ? (dy - by) : null;
    const ageIfAlive = (by != null && dy == null) ? (currentYear - by) : null;

    const marriages = marriagesByPID[pid] || [];
    const kids      = childrenByPID[pid]   || [];

    const lifespan = [by != null ? String(by) : '', dy != null ? String(dy) : '']
                      .join('–').replace(/^-+|-+$/g, '');

    const decade = (by != null) ? (Math.floor(by/10)*10 + 's') : '';

    // Build a structured list of marriages for the tooltip
    const spouses = marriages.map(mr => {
      const spouseName = TL_clean(mr[mIdx.SpouseName]);
      const mDate      = TL_clean(mr[mIdx.MarriageDate]);
      const mPlace     = TL_scrubPlace(mr[mIdx.MarriagePlace]);
      const status     = TL_clean(mr[mIdx.Status]);

      let married = '';
      if (mDate && mPlace) married = `${mDate} (${mPlace})`;
      else if (mDate)      married = mDate;
      else if (mPlace)     married = mPlace;

      const divorced = /divorc/i.test(status) ? status.replace(/^Divorced\\s*:?\s*/i, 'Divorced: ') : '';

      return { spouse: spouseName, married, divorced };
    }).filter(x => x.spouse || x.married || x.divorced);

    // Back-compat concatenated string
    const spouse = spouses
      .map(x => [x.spouse ? `Spouse: ${x.spouse}` : '', x.married ? `Married: ${x.married}` : '', x.divorced || ''].filter(Boolean).join(' — '))
      .filter(Boolean)
      .join(' • ');

    const children = kids.map(kr => {
      const nm = TL_clean(kr[cIdx.ChildName]);
      const y  = TL_yearFrom(kr[cIdx.BornDate]);
      return nm ? (y ? `${nm} (${y})` : nm) : null;
    }).filter(Boolean);

    return {
      id: pid,
      name: fullName,
      generation: TL_extractGen(r[pIdx.Generation], r[pIdx.Notes]),
      born: by,
      died: dy,
      age_at_death: ageAtDeath,
      current_age_if_alive: ageIfAlive,
      lifespan: lifespan,
      decade: decade,
      meta: TL_makeMeta(r[pIdx.PlaceOfBirth], r[pIdx.DateOfBirth], r[pIdx.PlaceOfDeath], r[pIdx.DateOfDeath], r[pIdx.Notes]),
      spouses: spouses,
      spouse: spouse,
      children: children
    };
  }).filter(p => p.name || p.born != null || p.died != null);

  people.sort((a,b) => (a.generation !== b.generation ? a.generation - b.generation : (a.born ?? 0) - (b.born ?? 0)));

  return people;
}

/** ========= Helpers (prefixed TL_) ========= */
function TL_clean(s) {
  return String(s || '')
    .replace(/\\u00A0|\\u200B|\\uFEFF/g, '')
    .replace(/\\s+/g, ' ')
    .trim();
}
function TL_scrubPlace(s) {
  s = TL_clean(s)
    .replace(/\\b(?:Unknown|Unknonw|Uknown)\\b/gi, '')
    .replace(/(^[,\\s]+|[,\\s]+$)/g, '');
  while (/, ,|,,|,\\s*,/.test(s)) s = s.replace(/, ,|,,|,\\s*,/g, ', ');
  return s.replace(/^\\s*,\\s*|\\s*,\\s*$/g, '').trim();
}
function TL_yearFrom(s) {
  s = TL_clean(s);
  if (!s) return null;
  const mIso = s.match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);
  if (mIso) {
    const y = Number(mIso[1]);
    return Number.isFinite(y) ? y : null;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear();
  const m = s.match(/\\b(1[6-9]\\d{2}|20\\d{2})\\b/);
  return m ? Number(m[1]) : null;
}
function TL_extractGen(genCell, notesCell) {
  const g = TL_clean(genCell);
  if (g) {
    const n = parseInt(g, 10);
    if (Number.isFinite(n)) return n;
  }
  const notes = TL_clean(notesCell);
  const m = notes.match(/Generation\\s*:\\s*(\\d+)/i);
  return m ? Number(m[1]) : 0;
}
function TL_makeMeta(placeBirth, dateBirth, placeDeath, dateDeath, notes) {
  const born = [TL_scrubPlace(placeBirth), TL_clean(dateBirth)].filter(Boolean).join(', ');
  const died = [TL_scrubPlace(placeDeath), TL_clean(dateDeath)].filter(Boolean).join(', ');
  const pieces = [];
  if (born) pieces.push(`Born: ${born}`);
  if (died) pieces.push(`Died: ${died}`);
  const main = pieces.join(' — ');
  const rest = TL_clean(notes);
  return rest ? (main ? `${main} • ${rest}` : rest) : main;
}

/** =========================
 * Delete Person (impact + cascade)
 * ========================= */

/** Return counts/ids of linked rows so UI can warn the user. */
function personHasDependents(personId) {
  ensureAllTables();
  const psh = peopleSheet_();
  const msh = marriageSheet_();
  const csh = childrenSheet_();

  const prow = findRowById_(psh, 'PersonID', personId);
  if (!prow) return { ok: false, code: 'NOT_FOUND', error: 'Person not found' };

  const personName = getValue_(psh, prow, 'FullName') || '';

  const mRows = findRowsWhere_(msh, 'PersonID', personId);
  const cRows = findRowsWhere_(csh, 'PersonID', personId);

  const mIds = mRows.map(r => getValue_(msh, r, 'MarriageID')).filter(Boolean);
  const cIds = cRows.map(r => getValue_(csh, r, 'ChildID')).filter(Boolean);

  return {
    ok: true,
    person: { id: personId, name: personName },
    marriages: { count: mRows.length, ids: mIds },
    children:  { count: cRows.length, ids: cIds }
  };
}

/**
 * Delete a person. If there are linked marriages/children, require cascade flags.
 * Deletes in safe order: children -> marriages -> person.
 */
function deletePersonWithCascade(personId, opts) {
  ensureAllTables();
  opts = opts || {};
  const psh = peopleSheet_();
  const msh = marriageSheet_();
  const csh = childrenSheet_();

  const prow = findRowById_(psh, 'PersonID', personId);
  if (!prow) return { ok: false, code: 'NOT_FOUND', error: 'Person not found' };

  const mRows = findRowsWhere_(msh, 'PersonID', personId);
  const cRows = findRowsWhere_(csh, 'PersonID', personId);

  if (mRows.length && !opts.cascadeMarriages) {
    return { ok: false, code: 'NEEDS_CASCADE_MARRIAGES', marriages: mRows.length, children: cRows.length, error: 'Person has marriage rows. Set cascadeMarriages:true to proceed.' };
  }
  if (cRows.length && !opts.cascadeChildren) {
    return { ok: false, code: 'NEEDS_CASCADE_CHILDREN', marriages: mRows.length, children: cRows.length, error: 'Person has child rows. Set cascadeChildren:true to proceed.' };
  }

  const deletedChildren = [];
  cRows.slice().sort((a,b)=>b-a).forEach(row => {
    const id = getValue_(csh, row, 'ChildID');
    if (id) deletedChildren.push(id);
    csh.deleteRow(row);
  });

  const deletedMarriages = [];
  mRows.slice().sort((a,b)=>b-a).forEach(row => {
    const id = getValue_(msh, row, 'MarriageID');
    if (id) deletedMarriages.push(id);
    msh.deleteRow(row);
  });

  const deletedPersonId = getValue_(psh, prow, 'PersonID');
  psh.deleteRow(prow);

  return { ok: true, deleted: { person: deletedPersonId, marriages: deletedMarriages, children: deletedChildren } };
}

/** Wrappers used by the sidebar */
function GX_getDeleteImpact(personId) {
  const res = personHasDependents(personId);
  if (!res.ok) return res;
  return {
    ok: true,
    person: res.person,
    marriages: res.marriages.count,
    children: res.children.count
  };
}

function GX_deletePersonCascade(personId, cascadeMarriages, cascadeChildren) {
  return deletePersonWithCascade(personId, {
    cascadeMarriages: !!cascadeMarriages,
    cascadeChildren:  !!cascadeChildren
  });
}

/** -------------------------------------------------------------------------------------
 * Optional one-off: convert existing text dates to real Date cells (no menu changes)
 * Run this from the Apps Script editor if you need to upgrade current data.
 * ------------------------------------------------------------------------------------- */
function GX_convertExistingDateStrings_exactHeaders() {
  const ss = SpreadsheetApp.getActive();
  const specs = [
    { sheetName: 'People',
      dateHeaders: ['DateOfBirth','DateOfDeath'],
      datetimeHeaders: ['Timestamp']
    },
    { sheetName: 'Marriages',
      dateHeaders: ['MarriageDate'],
      datetimeHeaders: ['Timestamp']
    },
    { sheetName: 'Children',
      dateHeaders: ['BornDate','DiedDate'],
      datetimeHeaders: ['Timestamp']
    },
  ];

  let report = [];
  specs.forEach(spec => {
    const sh = ss.getSheetByName(spec.sheetName);
    if (!sh) { report.push(`❌ Sheet not found: ${spec.sheetName}`); return; }

    const rng = sh.getDataRange();
    const vals = rng.getValues();
    const headers = (vals[0] || []).map(v => String(v||'').trim());
    const toIdx = h => headers.findIndex(x => x.toLowerCase() === h.toLowerCase());

    const dCols  = (spec.dateHeaders||[]).map(toIdx).filter(i => i>=0);
    const dtCols = (spec.datetimeHeaders||[]).map(toIdx).filter(i => i>=0);

    let converted = 0, skipped = 0;

    for (let r = 1; r < vals.length; r++) {
      dCols.forEach(ci => {
        const v = vals[r][ci];
        if (v === '' || v == null || v instanceof Date) return;
        const d = parseYMD_(v);
        if (d) { vals[r][ci] = d; converted++; } else { skipped++; }
      });
      dtCols.forEach(ci => {
        const v = vals[r][ci];
        if (v === '' || v == null || v instanceof Date) return;
        const d = parseDateTimeLoose_(v) || parseYMD_(v);
        if (d) { vals[r][ci] = d; converted++; } else { skipped++; }
      });
    }

    if (converted) {
      rng.setValues(vals);
      dCols.forEach(ci => sh.getRange(2, ci+1, vals.length-1, 1).setNumberFormat('yyyy-mm-dd'));
      dtCols.forEach(ci => sh.getRange(2, ci+1, vals.length-1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'));
    }

    report.push(`✅ ${spec.sheetName}: converted ${converted}, skipped ${skipped}`);
  });

  SpreadsheetApp.getUi().alert(report.join('\\n'));
}
/** =========================
 * Diagnostics: list & highlight date problems + fix People timestamps
 * ========================= */

/** Create a "Date Conversion Report" sheet listing any remaining string dates and blank People timestamps. */
function GX_showDateConversionIssues() {
  ensureAllTables();

  const ss = SpreadsheetApp.getActive();
  const reportName = 'Date Conversion Report';
  let rep = ss.getSheetByName(reportName);
  if (!rep) rep = ss.insertSheet(reportName);
  rep.clear();

  // Header
  const cols = ['Sheet','Row','ID','Field','Current Value','Go'];
  rep.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold');

  let out = [];

  // --- People ---
  (function(){
    const sh = peopleSheet_();
    const R = sh.getLastRow();
    if (R <= 1) return;
    const ids = pullColumn_(sh, 'PersonID');
    const dob = pullColumn_(sh, 'DateOfBirth');
    const dod = pullColumn_(sh, 'DateOfDeath');
    const ts  = pullColumn_(sh, 'Timestamp');

    for (let i=0;i<ids.length;i++){
      const row = i + 2;
      // any non-empty strings left in date columns?
      if (isNonEmptyString_(dob[i])) out.push(rowRecord_(sh, row, ids[i], 'DateOfBirth', dob[i]));
      if (isNonEmptyString_(dod[i])) out.push(rowRecord_(sh, row, ids[i], 'DateOfDeath', dod[i]));
      // Timestamp blanks? (flag only blanks; if you also want strings, add: isNonEmptyString_(ts[i]))
      if (isBlank_(ts[i])) out.push(rowRecord_(sh, row, ids[i], 'Timestamp', '(blank)'));
    }
  })();

  // --- Marriages (FYI: your converter should have handled these; included for completeness) ---
  (function(){
    const sh = marriageSheet_();
    const R = sh.getLastRow();
    if (R <= 1) return;
    const ids = pullColumn_(sh, 'MarriageID');
    const md  = pullColumn_(sh, 'MarriageDate');
    const ts  = pullColumn_(sh, 'Timestamp');
    for (let i=0;i<ids.length;i++){
      const row = i + 2;
      if (isNonEmptyString_(md[i])) out.push(rowRecord_(sh, row, ids[i], 'MarriageDate', md[i]));
      // (Not flagging blank timestamps here, but easy to add if you want)
    }
  })();

  // --- Children ---
  (function(){
    const sh = childrenSheet_();
    const R = sh.getLastRow();
    if (R <= 1) return;
    const ids = pullColumn_(sh, 'ChildID');
    const bd  = pullColumn_(sh, 'BornDate');
    const dd  = pullColumn_(sh, 'DiedDate');
    const ts  = pullColumn_(sh, 'Timestamp');

    for (let i=0;i<ids.length;i++){
      const row = i + 2;
      if (isNonEmptyString_(bd[i])) out.push(rowRecord_(sh, row, ids[i], 'BornDate', bd[i]));
      if (isNonEmptyString_(dd[i])) out.push(rowRecord_(sh, row, ids[i], 'DiedDate', dd[i]));
      // (You can also flag Children timestamp blanks by uncommenting below)
      // if (isBlank_(ts[i])) out.push(rowRecord_(sh, row, ids[i], 'Timestamp', '(blank)'));
    }
  })();

  if (!out.length) {
    rep.getRange(2,1,1,cols.length).setValues([['✓ No issues found','','','','','']]);
    SpreadsheetApp.getUi().alert('No remaining string dates or blank People timestamps found.');
    return;
  }

  // Write report rows
  const values = out.map(r => [r.sheet, r.row, r.id, r.field, r.value, makeA1Link_(r.gid, r.a1)]);
  rep.getRange(2,1,values.length,cols.length).setValues(values);
  rep.autoResizeColumns(1, cols.length);
  ss.setActiveSheet(rep);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(`Report created: ${out.length} issue(s) listed in "${reportName}".`);
}

/** Lightly highlight problem cells in-place (yellow background). */
function GX_highlightDateIssues() {
  ensureAllTables();
  const Y = '#fff3cd'; // pale yellow

  // People
  (function(){
    const sh = peopleSheet_();
    const R = sh.getLastRow(); if (R <= 1) return;
    const dobI = headerIndex_(sh,'DateOfBirth')+1;
    const dodI = headerIndex_(sh,'DateOfDeath')+1;
    const tsI  = headerIndex_(sh,'Timestamp')+1;

    const dvDob = sh.getRange(2, dobI, R-1, 1).getDisplayValues();
    const dvDod = sh.getRange(2, dodI, R-1, 1).getDisplayValues();
    const dvTs  = sh.getRange(2, tsI,  R-1, 1).getDisplayValues();

    const bgDob = sh.getRange(2, dobI, R-1, 1).getBackgrounds();
    const bgDod = sh.getRange(2, dodI, R-1, 1).getBackgrounds();
    const bgTs  = sh.getRange(2, tsI,  R-1, 1).getBackgrounds();

    for (let i=0;i<R-1;i++){
      if (isNonEmptyString_(dvDob[i][0])) bgDob[i][0] = Y;
      if (isNonEmptyString_(dvDod[i][0])) bgDod[i][0] = Y;
      if (isBlank_(dvTs[i][0]))           bgTs[i][0]  = Y;
    }
    sh.getRange(2, dobI, R-1, 1).setBackgrounds(bgDob);
    sh.getRange(2, dodI, R-1, 1).setBackgrounds(bgDod);
    sh.getRange(2, tsI,  R-1, 1).setBackgrounds(bgTs);
  })();

  // Marriages
  (function(){
    const sh = marriageSheet_();
    const R = sh.getLastRow(); if (R <= 1) return;
    const mdI = headerIndex_(sh,'MarriageDate')+1;
    const dv = sh.getRange(2, mdI, R-1, 1).getDisplayValues();
    const bg = sh.getRange(2, mdI, R-1, 1).getBackgrounds();
    for (let i=0;i<R-1;i++) if (isNonEmptyString_(dv[i][0])) bg[i][0] = Y;
    sh.getRange(2, mdI, R-1, 1).setBackgrounds(bg);
  })();

  // Children
  (function(){
    const sh = childrenSheet_();
    const R = sh.getLastRow(); if (R <= 1) return;
    const bdI = headerIndex_(sh,'BornDate')+1;
    const ddI = headerIndex_(sh,'DiedDate')+1;

    const dvBd = sh.getRange(2, bdI, R-1, 1).getDisplayValues();
    const dvDd = sh.getRange(2, ddI, R-1, 1).getDisplayValues();
    const bgBd = sh.getRange(2, bdI, R-1, 1).getBackgrounds();
    const bgDd = sh.getRange(2, ddI, R-1, 1).getBackgrounds();

    for (let i=0;i<R-1;i++){
      if (isNonEmptyString_(dvBd[i][0])) bgBd[i][0] = Y;
      if (isNonEmptyString_(dvDd[i][0])) bgDd[i][0] = Y;
    }
    sh.getRange(2, bdI, R-1, 1).setBackgrounds(bgBd);
    sh.getRange(2, ddI, R-1, 1).setBackgrounds(bgDd);
  })();

  SpreadsheetApp.getUi().alert('Problem cells highlighted (pale yellow).');
}

/** Backfill blank Timestamp cells in People with the current time, and format the column. */
function GX_fixBlankPeopleTimestamps() {
  ensureAllTables();
  const sh = peopleSheet_();
  const R = sh.getLastRow();
  if (R <= 1) { SpreadsheetApp.getUi().alert('People sheet is empty.'); return; }

  const tsCol = headerIndex_(sh, 'Timestamp') + 1;
  if (tsCol <= 0) throw new Error('Timestamp column not found in People.');

  const rng = sh.getRange(2, tsCol, R-1, 1);
  const vals = rng.getValues();        // preserves Dates if any already
  const disp = rng.getDisplayValues(); // for detecting blanks reliably

  let filled = 0;
  const now = new Date();
  for (let i=0;i<vals.length;i++){
    if (!disp[i][0]) { // blank as displayed
      vals[i][0] = now;
      filled++;
    }
  }
  if (filled > 0) {
    rng.setValues(vals);
  }

  // Make sure the Timestamp column uses a datetime format
  sh.getRange(2, tsCol, Math.max(1, R-1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  SpreadsheetApp.getUi().alert(`People Timestamp backfill complete. Filled ${filled} blank cell(s).`);
}

/** ---------- tiny utils used above ---------- */
function pullColumn_(sh, header){
  const idx = headerIndex_(sh, header) + 1;
  if (idx <= 0) return [];
  const R = sh.getLastRow(); if (R <= 1) return [];
  return sh.getRange(2, idx, R-1, 1).getDisplayValues().map(r => r[0]);
}
function isNonEmptyString_(v){
  return typeof v === 'string' && String(v).trim() !== '';
}
function isBlank_(v){
  return v === '' || v === null || typeof v === 'undefined';
}
function rowRecord_(sh, row, id, field, value){
  const gid = sh.getSheetId();
  const a1  = sh.getRange(row, headerIndex_(sh, field)+1).getA1Notation();
  return { sheet: sh.getName(), row: row, id: String(id||''), field, value: String(value||''), gid, a1 };
}
function makeA1Link_(gid, a1){
  // Clickable link that jumps to the exact cell in the source sheet
  return '=HYPERLINK("#gid=' + gid + '&range=' + a1 + '","Open")';
}
/** ============================================================
 * One-click: force-convert ISO-like text (YYYY-MM-DD[/ time]) to Date
 * across People, Marriages, and Children.
 *
 * Usage: Run GX_forceConvertIsoDates()
 * ============================================================ */
function GX_forceConvertIsoDates() {
  ensureAllTables();
  const tz = Session.getScriptTimeZone() || 'America/Chicago';

  // Sheet → columns to convert (by header)
  const PLAN = [
    { sh: peopleSheet_(),   cols: ['DateOfBirth','DateOfDeath','Timestamp'], dateFmt: 'yyyy-mm-dd', tsFmt: 'yyyy-mm-dd hh:mm:ss' },
    { sh: marriageSheet_(), cols: ['MarriageDate','Timestamp'],               dateFmt: 'yyyy-mm-dd', tsFmt: 'yyyy-mm-dd hh:mm:ss' },
    { sh: childrenSheet_(), cols: ['BornDate','DiedDate','Timestamp'],        dateFmt: 'yyyy-mm-dd', tsFmt: 'yyyy-mm-dd hh:mm:ss' },
  ];

  let totalChanged = 0;
  const perSheet = [];

  PLAN.forEach(({ sh, cols, dateFmt, tsFmt }) => {
    const name = sh.getName();
    const R = sh.getLastRow();
    if (R <= 1) { perSheet.push(`${name}: 0 changes (empty)`); return; }

    // Build index map and formats
    const idxMap = {};
    cols.forEach(h => { idxMap[h] = headerIndex_(sh, h) + 1; });

    // Pull display values for matching; pull raw for writing
    const rng = sh.getRange(2, 1, R-1, sh.getLastColumn());
    const disp = rng.getDisplayValues(); // strings as shown
    const vals = rng.getValues();        // preserves Dates

    let changed = 0;

    cols.forEach(h => {
      const c = idxMap[h];
      if (!c || c < 2) return; // header missing or wrong

      // Decide format to apply after write
      const isTimestamp = /timestamp/i.test(h);
      const fmt = isTimestamp ? tsFmt : dateFmt;

      for (let r = 0; r < R-1; r++) {
        const shown = disp[r][c-1];       // what user sees
        const cur   = vals[r][c-1];       // raw value
        if (!shown) continue;             // blank → skip
        if (cur instanceof Date) continue; // already Date

        // Match ISO date only or ISO date+time
        // e.g. 2020-07-04 or 2020-07-04 13:45:59 or 2020-07-04T13:45:59
        const mDT = String(shown).trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}):(\d{2}))?$/);
        if (!mDT) continue;

        const y = Number(mDT[1]), M = Number(mDT[2]), d = Number(mDT[3]);
        let H = 0, min = 0, s = 0;
        if (mDT[4] != null) { H = Number(mDT[4]); min = Number(mDT[5]); s = Number(mDT[6]); }

        // Construct a local Date (year, month-1, day, ...)
        // Using Date.UTC + Utilities.formatDate for TZ safety would be fine too,
        // but local constructor is straightforward for sheet storage.
        const asDate = new Date(y, M - 1, d, H, min, s);
        if (!isNaN(asDate.getTime())) {
          vals[r][c-1] = asDate;
          changed++;
        }
      }

      // Write the column back (only this column to reduce work)
      if (changed > 0) {
        sh.getRange(2, c, R-1, 1).setValues(vals.map(row => [row[c-1]]));
        // Apply a consistent format
        sh.getRange(2, c, Math.max(1, R-1), 1).setNumberFormat(fmt);
      }
    });

    totalChanged += changed;
    perSheet.push(`${name}: ${changed} cell(s) converted`);
  });

  SpreadsheetApp.getUi().alert(
    ['Force convert complete.',
     ...perSheet,
     '',
     `TOTAL converted: ${totalChanged}`].join('\n')
  );
}

/* (Optional) narrow version for just People if you want a quick rerun
function GX_forceConvertIsoDates_PeopleOnly() {
  const sh = peopleSheet_();
  _forceConvertIsoDates_onSheet(sh, ['DateOfBirth','DateOfDeath','Timestamp']);
}
*/

// If you prefer a reusable inner routine:
function _forceConvertIsoDates_onSheet(sh, headers, dateFmt='yyyy-mm-dd', tsFmt='yyyy-mm-dd hh:mm:ss') {
  const R = sh.getLastRow();
  if (R <= 1) return 0;
  const idx = {}; headers.forEach(h => idx[h] = headerIndex_(sh, h) + 1);

  const rng  = sh.getRange(2, 1, R-1, sh.getLastColumn());
  const disp = rng.getDisplayValues();
  const vals = rng.getValues();

  let changed = 0;

  headers.forEach(h => {
    const c = idx[h]; if (!c || c < 2) return;
    const isTimestamp = /timestamp/i.test(h);
    const fmt = isTimestamp ? tsFmt : dateFmt;

    for (let r = 0; r < R-1; r++) {
      const shown = disp[r][c-1];
      const cur   = vals[r][c-1];
      if (!shown) continue;
      if (cur instanceof Date) continue;

      const mDT = String(shown).trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}):(\d{2}))?$/);
      if (!mDT) continue;

      const y = Number(mDT[1]), M = Number(mDT[2]), d = Number(mDT[3]);
      let H = 0, min = 0, s = 0;
      if (mDT[4] != null) { H = Number(mDT[4]); min = Number(mDT[5]); s = Number(mDT[6]); }

      const asDate = new Date(y, M - 1, d, H, min, s);
      if (!isNaN(asDate.getTime())) {
        vals[r][c-1] = asDate;
        changed++;
      }
    }

    if (changed > 0) {
      sh.getRange(2, c, R-1, 1).setValues(vals.map(row => [row[c-1]]));
      sh.getRange(2, c, Math.max(1, R-1), 1).setNumberFormat(fmt);
    }
  });

  return changed;
}
/** =========================
 * REAL type-aware diagnostics (uses getValues(), not display strings)
 * ========================= */

/** Report any cells that are NOT true Dates in date/timestamp columns, and blank People Timestamps. */
function GX_showDateConversionIssues_REAL() {
  ensureAllTables();

  const ss = SpreadsheetApp.getActive();
  const reportName = 'Date Conversion Report';
  let rep = ss.getSheetByName(reportName);
  if (!rep) rep = ss.insertSheet(reportName);
  rep.clear();

  const headers = ['Sheet','Row','ID','Field','Current Value (shown)','Go'];
  rep.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');

  const out = [];

  scanSheetDates_(peopleSheet_(),  [
    {field:'DateOfBirth',   idField:'PersonID', flagBlank:false},
    {field:'DateOfDeath',   idField:'PersonID', flagBlank:false},
    {field:'Timestamp',     idField:'PersonID', flagBlank:true}  // also flag blanks here
  ], out);

  scanSheetDates_(marriageSheet_(),[
    {field:'MarriageDate',  idField:'MarriageID', flagBlank:false},
    {field:'Timestamp',     idField:'MarriageID', flagBlank:false}
  ], out);

  scanSheetDates_(childrenSheet_(),[
    {field:'BornDate',      idField:'ChildID', flagBlank:false},
    {field:'DiedDate',      idField:'ChildID', flagBlank:false},
    {field:'Timestamp',     idField:'ChildID', flagBlank:false}
  ], out);

  if (!out.length) {
    rep.getRange(2,1,1,headers.length).setValues([['✓ No type issues found','','','','','']]);
    SpreadsheetApp.getUi().alert('All date/timestamp cells are true Dates. No issues.');
    ss.setActiveSheet(rep);
    return;
  }

  // write rows
  const values = out.map(r => [r.sheet, r.row, r.id, r.field, r.display, makeA1Link_(r.gid, r.a1)]);
  rep.getRange(2,1,values.length,headers.length).setValues(values);
  rep.autoResizeColumns(1, headers.length);
  ss.setActiveSheet(rep);
  SpreadsheetApp.getUi().alert(`Report created: ${out.length} actual issue(s).`);
}

/** Highlight only true problems (non-Date values in date/timestamp columns; and blank People Timestamps). */
function GX_highlightDateIssues_REAL() {
  ensureAllTables();
  const Y = '#fff3cd'; // pale yellow

  highlightNonDates_(peopleSheet_(),  ['DateOfBirth','DateOfDeath'], Y);
  highlightNonDates_(marriageSheet_(),['MarriageDate'], Y);
  highlightNonDates_(childrenSheet_(),['BornDate','DiedDate'], Y);

  // People Timestamp: highlight blanks only
  (function(){
    const sh = peopleSheet_();
    const R = sh.getLastRow(); if (R <= 1) return;
    const c = headerIndex_(sh,'Timestamp') + 1; if (c <= 0) return;

    const rng = sh.getRange(2, c, R-1, 1);
    const vals = rng.getValues();        // types
    const disp = rng.getDisplayValues(); // to detect blanks
    const bgs  = rng.getBackgrounds();

    for (let i=0;i<R-1;i++){
      const v = vals[i][0];
      const isBlank = !disp[i][0];
      if (isBlank) bgs[i][0] = Y;
      else if (!(v instanceof Date)) bgs[i][0] = Y;
    }
    rng.setBackgrounds(bgs);
  })();

  SpreadsheetApp.getUi().alert('Type-based highlights applied.');
}

/** Clear all yellow highlights from date/timestamp columns, if needed. */
function GX_clearDateIssueHighlights() {
  [peopleSheet_(), marriageSheet_(), childrenSheet_()].forEach(sh => {
    const R = sh.getLastRow(); if (R <= 1) return;
    const cols = ['DateOfBirth','DateOfDeath','MarriageDate','BornDate','DiedDate','Timestamp']
      .map(h => headerIndex_(sh,h)+1).filter(c => c > 0);
    cols.forEach(c => sh.getRange(2,c,R-1,1).setBackground(null));
  });
}

/** ---------- helpers (real-type scanning) ---------- */

function scanSheetDates_(sh, plan, out) {
  const name = sh.getName();
  const R = sh.getLastRow(); if (R <= 1) return;

  // Build column indexes once
  const idx = {};
  const allHeaders = new Set(plan.map(p => p.field).concat(plan.map(p => p.idField)));
  allHeaders.forEach(h => { idx[h] = headerIndex_(sh, h) + 1; });

  const rngAll = sh.getRange(2, 1, R-1, sh.getLastColumn());
  const vals   = rngAll.getValues();         // types
  const shown  = rngAll.getDisplayValues();  // for user-friendly "Current Value"

  plan.forEach(p => {
    const cField = idx[p.field], cId = idx[p.idField];
    if (!cField || !cId) return;

    for (let r=0; r<R-1; r++){
      const v = vals[r][cField-1];
      const dText = shown[r][cField-1];      // what user sees
      const id = shown[r][cId-1];

      const isBlank = dText === '' || dText == null;
      const isDate = (v instanceof Date) && !isNaN(v.getTime());

      // Flag blanks only if instructed (People Timestamp), otherwise flag non-Date non-blank
      if (p.flagBlank && isBlank) {
        out.push(makeIssueRecord_(sh, r+2, id, p.field, '(blank)'));
      } else if (!isBlank && !isDate) {
        out.push(makeIssueRecord_(sh, r+2, id, p.field, dText));
      }
    }
  });
}

function highlightNonDates_(sh, fields, color) {
  const R = sh.getLastRow(); if (R <= 1) return;
  const rngAll = sh.getRange(2, 1, R-1, sh.getLastColumn());
  const vals   = rngAll.getValues();
  const shown  = rngAll.getDisplayValues();

  fields.forEach(h => {
    const c = headerIndex_(sh,h) + 1; if (c <= 0) return;
    const rng = sh.getRange(2, c, R-1, 1);
    const bgs = rng.getBackgrounds();

    for (let i=0;i<R-1;i++){
      const v = vals[i][c-1];
      const isBlank = shown[i][c-1] === '' || shown[i][c-1] == null;
      const isDate = (v instanceof Date) && !isNaN(v.getTime());
      if (!isBlank && !isDate) bgs[i][0] = color;
    }
    rng.setBackgrounds(bgs);
  });
}

function makeIssueRecord_(sh, row, id, field, displayText){
  const gid = sh.getSheetId();
  const a1  = sh.getRange(row, headerIndex_(sh, field)+1).getA1Notation();
  return { sheet: sh.getName(), row, id: String(id||''), field, display: String(displayText||''), gid, a1 };
}

/** You already have makeA1Link_ from earlier helpers; reusing it. If not present, uncomment:
function makeA1Link_(gid, a1){
  return '=HYPERLINK("#gid=' + gid + '&range=' + a1 + '","Open")';
}
*/
/**
 * One-click maintenance:
 * 1) Apply standard formats (yyyy-mm-dd; yyyy-mm-dd hh:mm:ss for Timestamp)
 * 2) Backfill People→Timestamp blanks with now
 * 3) Scan for REAL type issues (non-Date cells in date/timestamp columns)
 * 4) If the REAL report function exists, regenerate the report
 */
function GX_maintenanceDates() {
  ensureAllTables();

  const PLAN = [
    { sh: peopleSheet_(),   dates: ['DateOfBirth','DateOfDeath'], ts: 'Timestamp' },
    { sh: marriageSheet_(), dates: ['MarriageDate'],               ts: 'Timestamp' },
    { sh: childrenSheet_(), dates: ['BornDate','DiedDate'],        ts: 'Timestamp' },
  ];

  let filledTimestamps = 0;
  let nonDateIssues = 0;
  const perSheetNotes = [];

  PLAN.forEach(({ sh, dates, ts }) => {
    const name = sh.getName();
    const R = sh.getLastRow();
    if (R <= 1) { perSheetNotes.push(`${name}: empty`); return; }

    // 1) Apply formats to each date column
    dates.forEach(h => {
      const c = headerIndex_(sh, h) + 1; if (c <= 0) return;
      sh.getRange(2, c, R-1, 1).setNumberFormat('yyyy-mm-dd');
    });

    // Timestamp column formatting
    let tsCol = -1;
    if (ts) {
      tsCol = headerIndex_(sh, ts) + 1;
      if (tsCol > 0) sh.getRange(2, tsCol, R-1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }

    // 2) Backfill People → Timestamp blanks only
    if (name === 'People' && tsCol > 0) {
      const rng = sh.getRange(2, tsCol, R-1, 1);
      const vals = rng.getValues();
      const shown = rng.getDisplayValues();
      const now = new Date();
      let filled = 0;
      for (let i=0;i<vals.length;i++){
        if (!shown[i][0]) { vals[i][0] = now; filled++; }
      }
      if (filled) rng.setValues(vals);
      filledTimestamps += filled;
    }

    // 3) Scan for REAL type issues (non-Date, non-blank) in dates + timestamp
    const allCols = [...dates, ts].filter(Boolean).map(h => headerIndex_(sh, h) + 1).filter(c => c > 0);
    if (allCols.length) {
      const rngAll = sh.getRange(2, 1, R-1, sh.getLastColumn());
      const vals   = rngAll.getValues();         // real types
      const shown  = rngAll.getDisplayValues();  // to check blank vs non-blank

      allCols.forEach(c => {
        for (let r=0; r<R-1; r++){
          const display = shown[r][c-1];
          if (!display) continue;                 // ignore blanks here (People blanks were filled)
          const v = vals[r][c-1];
          if (!(v instanceof Date) || isNaN(v.getTime())) nonDateIssues++;
        }
      });
    }

    perSheetNotes.push(`${name}: formats applied${name==='People' ? `; ${filledTimestamps} timestamp(s) filled so far` : ''}`);
  });

  // 4) Optional: regenerate the REAL report if present
  try {
    if (typeof GX_showDateConversionIssues_REAL === 'function') {
      GX_showDateConversionIssues_REAL();
    } else {
      // If the REAL report function isn't present, create a tiny summary sheet
      const ss = SpreadsheetApp.getActive();
      const name = 'Date Conversion Report';
      let rep = ss.getSheetByName(name);
      if (!rep) rep = ss.insertSheet(name);
      rep.clear();
      rep.getRange(1,1,1,2).setValues([['Summary','Value']]).setFontWeight('bold');
      rep.getRange(2,1,1,2).setValues([['Non-Date issues found (scan)', nonDateIssues]]);
      ss.setActiveSheet(rep);
    }
  } catch (e) {
    // Non-fatal; continue to alert
  }

  SpreadsheetApp.getUi().alert(
    [
      'Maintenance complete.',
      `• People timestamps filled: ${filledTimestamps}`,
      `• Non-Date issues detected (scan): ${nonDateIssues}`,
      '',
      'Sheets:',
      ...perSheetNotes
    ].join('\n')
  );
}
// Prompt-driven entry point
function findMissingAnySurname() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Find Missing by Surname', 'Enter a surname (e.g., Kelley):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const surname = String(res.getResponseText() || '').trim();
  if (!surname) {
    ui.alert('Please enter a non-empty surname.');
    return;
  }
  findMissingBySurname(surname);
}

// Core function with a parameter
function findMissingBySurname(surname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const childrenSheet = ss.getSheetByName('Children');
  const peopleSheet   = ss.getSheetByName('People');
  if (!childrenSheet) throw new Error("Missing 'Children' sheet");
  if (!peopleSheet)   throw new Error("Missing 'People' sheet");
  const targetSurname = String(surname || '').trim().toLowerCase();
  if (!targetSurname) throw new Error('Surname is required');

  // ----- helpers -----
  const getColIndex = (sheet, headerName) => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const i = headers.indexOf(headerName);
    if (i === -1) throw new Error(`No '${headerName}' column found in '${sheet.getName()}' sheet.`);
    return i + 1; // 1-based
  };
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const lastWord = s => {
    const t = String(s || '').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    const parts = t.split(' ');
    return parts[parts.length - 1].toLowerCase();
  };
  const parseIsoDateStrict = val => {
    if (typeof val !== 'string') return null;
    const m = val.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = +m[1], mo = +m[2], d = +m[3];
    const dt = new Date(y, mo - 1, d); // local midnight, avoids TZ drift
    return (dt.getFullYear() === y && dt.getMonth() + 1 === mo && dt.getDate() === d) ? dt : null;
  };
  const coerceToDateOrBlank = v => {
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    const parsed = parseIsoDateStrict(v);
    return parsed ? parsed : '';
  };

  // ----- columns -----
  // Children
  const colChildName = getColIndex(childrenSheet, 'ChildName');
  const colBornDate  = getColIndex(childrenSheet, 'BornDate');
  const colBornPlace = getColIndex(childrenSheet, 'BornPlace');
  const colDiedDate  = getColIndex(childrenSheet, 'DiedDate');
  const colDiedPlace = getColIndex(childrenSheet, 'DiedPlace');
  // People
  const colFullName  = getColIndex(peopleSheet, 'FullName');  // <-- this was missing

  // ----- data -----
  const childNumRows  = Math.max(0, childrenSheet.getLastRow() - 1);
  const peopleNumRows = Math.max(0, peopleSheet.getLastRow() - 1);

  const childrenRows = childNumRows
    ? childrenSheet.getRange(2, 1, childNumRows, childrenSheet.getLastColumn()).getValues()
    : [];

  const peopleValues = peopleNumRows
    ? peopleSheet.getRange(2, colFullName, peopleNumRows, 1).getValues().flat()
    : [];

  const peopleSet = new Set(peopleValues.map(v => norm(v)));

  const idx = {
    ChildName:  colChildName - 1,
    BornDate:   colBornDate  - 1,
    BornPlace:  colBornPlace - 1,
    DiedDate:   colDiedDate  - 1,
    DiedPlace:  colDiedPlace - 1,
  };

  // ----- find missing -----
  const missingMap = new Map(); // normFullName -> [FullName, DOB(Date or ''), POB, DOD(Date or ''), POD]
  for (const row of childrenRows) {
    const childName = row[idx.ChildName];
    const cleanName = String(childName || '').replace(/\s+/g, ' ').trim();
    if (!cleanName) continue;

    if (lastWord(cleanName) !== targetSurname) continue;     // surname filter
    if (peopleSet.has(norm(cleanName))) continue;            // already exists in People

    // Dates: keep real Dates; parse ISO 'YYYY-MM-DD' if present; else blank
    const dob = coerceToDateOrBlank(row[idx.BornDate]);
    const dod = coerceToDateOrBlank(row[idx.DiedDate]);

    const pob = String(row[idx.BornPlace] || '').trim();
    const pod = String(row[idx.DiedPlace] || '').trim();

    const key = norm(cleanName);
    if (!missingMap.has(key)) {
      missingMap.set(key, [cleanName, dob, pob, dod, pod]);
    }
  }

  // ----- output -----
  const outName = `Missing_${capitalizeSafe(targetSurname)}`;
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName); else out.clear();

  const headers = ['FullName', 'DateOfBirth', 'PlaceOfBirth', 'DateOfDeath', 'PlaceOfDeath'];
  out.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = Array.from(missingMap.values());
  if (rows.length) {
    out.getRange(2, 1, rows.length, headers.length).setValues(rows);
    // format date columns (doesn't coerce text, only affects real Dates)
    out.getRange(2, 2, rows.length, 1).setNumberFormat('yyyy-mm-dd'); // DateOfBirth
    out.getRange(2, 4, rows.length, 1).setNumberFormat('yyyy-mm-dd'); // DateOfDeath
  } else {
    out.getRange(2, 1).setValue(`No missing ${capitalizeSafe(targetSurname)} found!`);
  }
}

function capitalizeSafe(s) {
  s = String(s || '').trim();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
}


// Helper to pretty-capitalize a surname for the output sheet name
function capitalizeSafe(s) {
  s = String(s || '').trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function recalcGenerationsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Recalculate Generations', 'Enter a surname (e.g., Kelley):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const surname = String(res.getResponseText() || '').trim();
  if (!surname) { ui.alert('Please enter a non-empty surname.'); return; }
  recalcGenerationsForSurname(surname);
}

function recalcGenerationsForSurname(surname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const peopleSheet = ss.getSheetByName('People');
  const childrenSheet = ss.getSheetByName('Children');
  if (!peopleSheet) throw new Error("Missing 'People' sheet");
  if (!childrenSheet) throw new Error("Missing 'Children' sheet");

  const targetSurname = String(surname).trim().toLowerCase();

  // --- Helpers ---
  const getColIndex = (sheet, headerName) => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const i = headers.indexOf(headerName);
    if (i === -1) throw new Error(`No '${headerName}' column found in '${sheet.getName()}' sheet.`);
    return i + 1; // 1-based
  };
  const getFirstExistingColIndex = (sheet, headerNames) => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (const h of headerNames) {
      const i = headers.indexOf(h);
      if (i !== -1) return i + 1;
    }
    throw new Error(`None of the columns [${headerNames.join(', ')}] were found in '${sheet.getName()}' sheet.`);
  };
  const lastWord = s => {
    const t = String(s || '').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    const parts = t.split(' ');
    return parts[parts.length - 1].toLowerCase();
  };
  const isSurname = name => lastWord(name) === targetSurname;
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();

  // --- Column indexes ---
  // People
  const colPID        = getColIndex(peopleSheet, 'PersonID');
  const colFullName   = getColIndex(peopleSheet, 'FullName');
  const colGeneration = getColIndex(peopleSheet, 'Generation');

  // Children
  const colChildName = getColIndex(childrenSheet, 'ChildName');
  // Try common parent ID column names (your Add Child stores the selected person)
  const colParentPID = getFirstExistingColIndex(childrenSheet, ['ParentPersonID', 'ParentID']);

  // --- Read data ---
  const peopleRowsN = Math.max(0, peopleSheet.getLastRow() - 1);
  const childrenRowsN = Math.max(0, childrenSheet.getLastRow() - 1);

  const peopleRange = peopleRowsN
    ? peopleSheet.getRange(2, 1, peopleRowsN, peopleSheet.getLastColumn()).getValues()
    : [];
  const childrenRange = childrenRowsN
    ? childrenSheet.getRange(2, 1, childrenRowsN, childrenSheet.getLastColumn()).getValues()
    : [];

  // --- Build fast lookups ---
  // peopleRowByPID: PID -> {rowIndex, pid, fullname, gen}
  const peopleRowByPID = new Map();
  // pidByFullName: norm(fullname) -> PID
  const pidByFullName = new Map();

  for (let r = 0; r < peopleRange.length; r++) {
    const row = peopleRange[r];
    const pid = String(row[colPID - 1] || '').trim();
    if (!pid) continue;
    const fullname = String(row[colFullName - 1] || '').trim();
    const genVal = row[colGeneration - 1];
    const gen = (typeof genVal === 'number' && isFinite(genVal)) ? genVal : null;

    peopleRowByPID.set(pid, { rowIndex: r, pid, fullname, gen });
    if (fullname) pidByFullName.set(norm(fullname), pid);
  }

  // childrenByParentPID: parentPID -> [childPID]
  // We resolve child PID by matching ChildName to People.FullName (exact match, case-insensitive, trimmed).
  const childrenByParentPID = new Map();
  for (let r = 0; r < childrenRange.length; r++) {
    const row = childrenRange[r];
    const parentPID = String(row[colParentPID - 1] || '').trim();
    const childName = String(row[colChildName - 1] || '').trim();
    if (!parentPID || !childName) continue;

    const childPID = pidByFullName.get(norm(childName));
    if (!childPID) continue; // child not yet in People; skip

    if (!childrenByParentPID.has(parentPID)) childrenByParentPID.set(parentPID, []);
    childrenByParentPID.get(parentPID).push(childPID);
  }

  // --- BFS from all Generation==1 seeds of the target surname ---
  const queue = [];
  // Track updates we’ll write back at the end
  const pendingGenByPID = new Map();

  for (const { pid, fullname, gen } of peopleRowByPID.values()) {
    if (gen === 1 && isSurname(fullname)) {
      queue.push({ pid, gen: 1 });
    }
  }

  // If there are no seeds, we do nothing (avoid setting arbitrary generation)
  if (queue.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No Generation=1 seeds found in People for surname "${surname}".\n` +
      `Set one or more founders (Generation = 1) first, then rerun.`
    );
    return;
  }

  // Use existing generations as lower bounds
  const currentGenByPID = new Map();
  for (const { pid, gen } of peopleRowByPID.values()) {
    if (typeof gen === 'number' && isFinite(gen)) {
      currentGenByPID.set(pid, gen);
    }
  }

  while (queue.length) {
    const { pid: parentPID, gen: parentGen } = queue.shift();
    const kids = childrenByParentPID.get(parentPID) || [];
    for (const childPID of kids) {
      const childInfo = peopleRowByPID.get(childPID);
      if (!childInfo) continue;

      // Only assign/adjust generation if child’s surname matches target
      if (!isSurname(childInfo.fullname)) continue;

      const proposed = parentGen + 1;
      const existing = currentGenByPID.get(childPID);

      // If no generation yet, or we found a smaller generation number, update
      if (!(typeof existing === 'number' && isFinite(existing)) || proposed < existing) {
        currentGenByPID.set(childPID, proposed);
        pendingGenByPID.set(childPID, proposed);
        queue.push({ pid: childPID, gen: proposed });
      }
    }
  }

  // --- Write back Generation only where it changed (batch write) ---
  if (pendingGenByPID.size === 0) {
    // Nothing to update; done.
    return;
  }

  // Prepare a sparse write: collect [row, value]
  const writes = [];
  for (const [pid, newGen] of pendingGenByPID.entries()) {
    const info = peopleRowByPID.get(pid);
    if (!info) continue;
    writes.push({ rowIndex: info.rowIndex, value: newGen });
  }

  // Sort by row for efficient setValues chunks
  writes.sort((a, b) => a.rowIndex - b.rowIndex);

  // We’ll write in contiguous blocks
  let i = 0;
  while (i < writes.length) {
    const startRow = writes[i].rowIndex;
    let end = i;
    // extend contiguous region
    while (end + 1 < writes.length && writes[end + 1].rowIndex === writes[end].rowIndex + 1) {
      end++;
    }
    const numRows = (end - i + 1);
    const block = new Array(numRows).fill(0).map((_, k) => [writes[i + k].value]);

    // Rows in sheet are 1-based, with headers on row 1; our rowIndex is zero-based from data start
    peopleSheet.getRange(2 + startRow, colGeneration, numRows, 1).setValues(block);

    i = end + 1;
  }

  // Optional: format the Generation column as integer
  peopleSheet.getRange(2, colGeneration, Math.max(peopleRowsN, 1), 1).setNumberFormat('0');
}
function recalcGenerationsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Recalculate Generations', 'Enter a surname (e.g., Kelley):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const surname = String(res.getResponseText() || '').trim();
  if (!surname) { ui.alert('Please enter a non-empty surname.'); return; }
  recalcGenerationsForSurname(surname);
}

function recalcGenerationsForSurname(surname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const peopleSheet    = ss.getSheetByName('People');
  const childrenSheet  = ss.getSheetByName('Children');
  const marriagesSheet = ss.getSheetByName('Marriages');
  if (!peopleSheet || !childrenSheet) throw new Error("Missing 'People' or 'Children' sheet");

  const targetSurname = String(surname).trim().toLowerCase();

  // Helpers
  const headersOf = sh => sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const getColIndex = (sheet, headerName) => {
    const headers = headersOf(sheet);
    const i = headers.indexOf(headerName);
    if (i === -1) throw new Error(`No '${headerName}' column found in '${sheet.getName()}' sheet.`);
    return i + 1; // 1-based
  };
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const lastWord = s => {
    const t = String(s || '').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    const parts = t.split(' ');
    return parts[parts.length - 1].toLowerCase();
  };
  const isSurname = name => lastWord(name) === targetSurname;

  // Column indexes (exact headers you gave)
  const pPID        = getColIndex(peopleSheet,   'PersonID');
  const pFullName   = getColIndex(peopleSheet,   'FullName');
  const pGeneration = getColIndex(peopleSheet,   'Generation');

  const cChildID    = getColIndex(childrenSheet, 'ChildID');    // not used directly, but confirms layout
  const cParentPID  = getColIndex(childrenSheet, 'PersonID');   // <-- parent link
  const cMarriageID = getColIndex(childrenSheet, 'MarriageID'); // <-- marriage link (optional)
  const cChildName  = getColIndex(childrenSheet, 'ChildName');

  // Marriages is optional, but useful for second parent
  let mMarriageID = 0, mPersonID = 0, mSpouseName = 0;
  if (marriagesSheet) {
    mMarriageID = getColIndex(marriagesSheet, 'MarriageID');
    mPersonID   = getColIndex(marriagesSheet, 'PersonID');   // primary spouse PID
    mSpouseName = getColIndex(marriagesSheet, 'SpouseName'); // will resolve to PID via People.FullName
  }

  // Read People
  const peopleN = Math.max(0, peopleSheet.getLastRow() - 1);
  const peopleData = peopleN ? peopleSheet.getRange(2, 1, peopleN, peopleSheet.getLastColumn()).getValues() : [];

  // Build lookups
  const peopleRowByPID = new Map();    // PID -> {rowIndex, fullname, gen}
  const pidByFullName  = new Map();    // norm(fullname) -> PID
  for (let r = 0; r < peopleData.length; r++) {
    const row = peopleData[r];
    const pid = String(row[pPID - 1] || '').trim();
    if (!pid) continue;
    const fullname = String(row[pFullName - 1] || '').trim();
    const genVal = row[pGeneration - 1];
    const gen = (typeof genVal === 'number' && isFinite(genVal)) ? genVal : null;
    peopleRowByPID.set(pid, { rowIndex: r, fullname, gen });
    if (fullname) pidByFullName.set(norm(fullname), pid);
  }

  // Read Children
  const childN = Math.max(0, childrenSheet.getLastRow() - 1);
  const childData = childN ? childrenSheet.getRange(2, 1, childN, childrenSheet.getLastColumn()).getValues() : [];

  // Read Marriages (if present)
  const marriageByID = new Map(); // MarriageID -> {primaryPID, spousePID?}
  if (marriagesSheet) {
    const mN = Math.max(0, marriagesSheet.getLastRow() - 1);
    const marrData = mN ? marriagesSheet.getRange(2, 1, mN, marriagesSheet.getLastColumn()).getValues() : [];
    for (const row of marrData) {
      const mid = String(row[mMarriageID - 1] || '').trim();
      if (!mid) continue;
      const primaryPID = String(row[mPersonID - 1] || '').trim() || null;
      const spouseName = String(row[mSpouseName - 1] || '').trim();
      const spousePID  = spouseName ? (pidByFullName.get(norm(spouseName)) || null) : null;
      marriageByID.set(mid, { primaryPID, spousePID });
    }
  }

  // Build parentPID -> [childPID] map using:
  // 1) Children.PersonID (direct parent)
  // 2) Children.MarriageID -> (Marriages.PersonID + resolved SpouseName)
  const childrenByParentPID = new Map();
  const pushChild = (parentPID, childPID) => {
    if (!parentPID || !childPID) return;
    if (!childrenByParentPID.has(parentPID)) childrenByParentPID.set(parentPID, []);
    childrenByParentPID.get(parentPID).push(childPID);
  };

  for (const row of childData) {
    const parentPID = String(row[cParentPID - 1] || '').trim();     // direct parent
    const mid       = String(row[cMarriageID - 1] || '').trim();    // marriage link
    const childName = String(row[cChildName - 1] || '').trim();
    if (!childName) continue;

    const childPID = pidByFullName.get(norm(childName));
    if (!childPID) continue; // only propagate to children already present in People

    if (parentPID) pushChild(parentPID, childPID);
    if (mid && marriageByID.has(mid)) {
      const { primaryPID, spousePID } = marriageByID.get(mid);
      if (primaryPID) pushChild(primaryPID, childPID);
      if (spousePID)  pushChild(spousePID,  childPID);
    }
  }

  // BFS from all Generation==1 seeds with matching surname
  const queue = [];
  const currentGenByPID = new Map(); // starting point: whatever already in People
  for (const [pid, info] of peopleRowByPID.entries()) {
    if (typeof info.gen === 'number' && isFinite(info.gen)) currentGenByPID.set(pid, info.gen);
    if (info.gen === 1 && isSurname(info.fullname)) queue.push({ pid, gen: 1 });
  }

  if (queue.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No Generation=1 seeds found for surname "${surname}".\n` +
      `Set one or more founders (Generation = 1) on People, then rerun.`
    );
    return;
  }

  const pendingGenByPID = new Map(); // updates to write

  while (queue.length) {
    const { pid: parentPID, gen: parentGen } = queue.shift();
    const kids = childrenByParentPID.get(parentPID) || [];
    for (const childPID of kids) {
      const childInfo = peopleRowByPID.get(childPID);
      if (!childInfo) continue;

      // Only assign generation if child's surname matches target
      if (!isSurname(childInfo.fullname)) continue;

      const proposed = parentGen + 1;
      const existing = currentGenByPID.get(childPID);

      if (!(typeof existing === 'number' && isFinite(existing)) || proposed < existing) {
        currentGenByPID.set(childPID, proposed);
        pendingGenByPID.set(childPID, proposed);
        queue.push({ pid: childPID, gen: proposed });
      }
    }
  }

  if (pendingGenByPID.size === 0) return; // nothing to update

  // Batch write back to People.Generation
  const writes = [];
  for (const [pid, newGen] of pendingGenByPID.entries()) {
    const info = peopleRowByPID.get(pid);
    if (!info) continue;
    writes.push({ rowIndex: info.rowIndex, value: newGen });
  }
  writes.sort((a, b) => a.rowIndex - b.rowIndex);

  let i = 0;
  while (i < writes.length) {
    const start = writes[i].rowIndex;
    let end = i;
    while (end + 1 < writes.length && writes[end + 1].rowIndex === writes[end].rowIndex + 1) end++;
    const len = end - i + 1;
    const block = new Array(len).fill(0).map((_, k) => [writes[i + k].value]);
    peopleSheet.getRange(2 + start, pGeneration, len, 1).setValues(block);
    i = end + 1;
  }

  // Format Generation as integer
  peopleSheet.getRange(2, pGeneration, Math.max(peopleN, 1), 1).setNumberFormat('0');
}
/**
 * Import all rows from a Missing_* sheet into People and auto-assign PersonID.
 * Example: GX_importMissingToPeople('Missing_Kelley')
 */
function GX_importMissingToPeople(missingSheetName) {
  const ss = SpreadsheetApp.getActive();
  const people = ss.getSheetByName('People');
  const missing = ss.getSheetByName(missingSheetName);
  if (!people) throw new Error("Missing 'People' sheet");
  if (!missing) throw new Error(`Missing '${missingSheetName}' sheet`);

  const pHeaders = people.getRange(1,1,1,people.getLastColumn()).getValues()[0];
  const mHeaders = missing.getRange(1,1,1,missing.getLastColumn()).getValues()[0];

  const pMap = (name) => {
    const i = pHeaders.indexOf(name);
    if (i === -1) throw new Error(`People missing '${name}' column`);
    return i + 1;
  };
  const mMap = (name) => {
    const i = mHeaders.indexOf(name);
    if (i === -1) throw new Error(`${missingSheetName} missing '${name}' column`);
    return i + 1;
  };

  const colP_PersonID  = pMap('PersonID');
  const colP_FullName  = pMap('FullName');
  const colP_DOB       = pMap('DateOfBirth');
  const colP_POB       = pMap('PlaceOfBirth');
  const colP_DOD       = pMap('DateOfDeath');
  const colP_POD       = pMap('PlaceOfDeath');
  const colP_Parents   = pMap('Parents');
  const colP_Notes     = pMap('Notes');
  const colP_Generation= pMap('Generation');
  const colP_Timestamp = pMap('Timestamp');

  const colM_FullName = mMap('FullName');
  const colM_DOB      = mMap('DateOfBirth');
  const colM_POB      = mMap('PlaceOfBirth');
  const colM_DOD      = mMap('DateOfDeath');
  const colM_POD      = mMap('PlaceOfDeath');

  const lastRow = missing.getLastRow();
  if (lastRow < 2) return;

  const mVals = missing.getRange(2, 1, lastRow - 1, missing.getLastColumn()).getValues();

  // Skip duplicates by FullName already in People
  const pLast = people.getLastRow();
  const pFullCol = pLast > 1 ? people.getRange(2, colP_FullName, pLast - 1, 1).getValues().flat() : [];
  const pFullSet = new Set(pFullCol.map(v => String(v || '').trim().toLowerCase()));

  const nextId = () => Utilities.getUuid();

  const outRows = [];
  const importedFullNames = []; // track names we actually import

  for (const row of mVals) {
    const fullName = String(row[colM_FullName - 1] || '').trim();
    if (!fullName) continue;
    if (pFullSet.has(fullName.toLowerCase())) continue; // already in People

    const dob = row[colM_DOB - 1]; // Date or ''
    const pob = row[colM_POB - 1];
    const dod = row[colM_DOD - 1]; // Date or ''
    const pod = row[colM_POD - 1];

    const newRow = new Array(pHeaders.length).fill('');
    newRow[colP_PersonID - 1]   = nextId();
    newRow[colP_FullName - 1]   = fullName;
    newRow[colP_DOB - 1]        = (dob instanceof Date && !isNaN(dob)) ? dob : '';
    newRow[colP_POB - 1]        = String(pob || '').trim();
    newRow[colP_DOD - 1]        = (dod instanceof Date && !isNaN(dod)) ? dod : '';
    newRow[colP_POD - 1]        = String(pod || '').trim();
    newRow[colP_Parents - 1]    = '';
    newRow[colP_Notes - 1]      = `Imported from ${missingSheetName}`;
    newRow[colP_Generation - 1] = '';
    newRow[colP_Timestamp - 1]  = new Date();

    outRows.push(newRow);
    importedFullNames.push(fullName);
  }

  if (outRows.length === 0) return;

  // Append to People
  const startRow = people.getLastRow() + 1;
  people.getRange(startRow, 1, outRows.length, pHeaders.length).setValues(outRows);

  // Friendly formats for new rows
  people.getRange(startRow, colP_DOB, outRows.length, 1).setNumberFormat('yyyy-mm-dd');
  people.getRange(startRow, colP_DOD, outRows.length, 1).setNumberFormat('yyyy-mm-dd');
  people.getRange(startRow, colP_Generation, outRows.length, 1).setNumberFormat('0');
  people.getRange(startRow, colP_Timestamp, outRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');

  // === NEW: auto-recalc Generation for the imported surnames ===
  // Requires helpers: _gx_lastWordLower, _gx_capitalize, _gx_throttledRecalcGenerationsForSurname
  const surnameSet = new Set();
  for (const name of importedFullNames) {
    const s = _gx_lastWordLower(name);
    if (s) surnameSet.add(_gx_capitalize(s));
  }
  for (const s of surnameSet) {
    _gx_throttledRecalcGenerationsForSurname(s);
  }
}

// === Helpers for surname handling & throttled generation recalculation ===
// === Helpers for surname handling & throttled generation recalculation ===

/** Return the last word of a name in lowercase (used for surname detection). */
function _gx_lastWordLower(s) {
  const t = String(s || '').replace(/\s+/g, ' ').trim();
  if (!t) return '';
  const parts = t.split(' ');
  return parts[parts.length - 1].toLowerCase();
}

/** Capitalize the first letter of a string (Kelley, Johnson, etc.). */
function _gx_capitalize(word) {
  word = String(word || '').trim();
  return word ? word[0].toUpperCase() + word.slice(1) : '';
}

/**
 * Debounced + locked call to recalcGenerationsForSurname(surname).
 * Ensures only one recalculation at a time and skips repeats within 10s.
 */
function _gx_throttledRecalcGenerationsForSurname(surname) {
  const SURNAME = String(surname || '').trim();
  if (!SURNAME) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // couldn't get lock quickly

  try {
    const key = 'lastGenRecalc_' + SURNAME.toLowerCase(); // avoid backticks
    const props = PropertiesService.getScriptProperties();
    const last = Number(props.getProperty(key) || '0');
    const now  = Date.now();

    if (now - last < 10000) return; // debounce 10s

    recalcGenerationsForSurname(SURNAME);
    props.setProperty(key, String(now));
  } finally {
    lock.releaseLock();
  }
}
function GX_importMissingPrompt() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const missingSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => /^Missing_/i.test(n));

  if (missingSheets.length === 0) {
    ui.alert('No sheets found whose name starts with "Missing_". Run your finder first.');
    return;
  }

  if (missingSheets.length === 1) {
    const only = missingSheets[0];
    const confirm = ui.alert('Import Missing → People', `Import from "${only}"?`, ui.ButtonSet.OK_CANCEL);
    if (confirm !== ui.Button.OK) return;
    GX_importMissingToPeople(only);
    return;
  }

  // Multiple options: prompt user to type one of the names
  const res = ui.prompt(
    'Import Missing → People',
    'Type one of these sheet names exactly:\n\n' + missingSheets.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const chosen = String(res.getResponseText() || '').trim();
  if (!missingSheets.includes(chosen)) {
    ui.alert(`"${chosen}" wasn’t found. Please choose one of:\n\n` + missingSheets.join('\n'));
    return;
  }
  GX_importMissingToPeople(chosen);
}
// Reusable no-op for menu headers
function GX_menuHeader() {
  // Do nothing (optional: show alert if you want feedback)
  return;
}



