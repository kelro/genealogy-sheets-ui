/** =========================================================
 * NormalizeDates.gs  —  One-time ISO date normalization
 * Converts legacy dates to ISO (YYYY-MM-DD) for sidebar pickers.
 * Safe with other scripts: all identifiers are NORM_* or unique.
 * =========================================================
 *
 * MODE:
 *  - 'STRICT' (default): Year-only stays blank.
 *  - 'APPROX'          : Year-only 'YYYY' -> 'YYYY-01-01' (logged).
 */
const NORMALIZE_MODE = 'APPROX'; // 'STRICT' | 'APPROX'

/** Target sheets & columns (unique names to avoid conflicts) */
const NORM_SHEET_PEOPLE   = 'People';
const NORM_SHEET_MARRIAGE = 'Marriages';
const NORM_SHEET_CHILDREN = 'Children';

const NORM_COLS_PEOPLE   = ['DateOfBirth','DateOfDeath'];
const NORM_COLS_MARRIAGE = ['MarriageDate'];
const NORM_COLS_CHILDREN = ['BornDate','DiedDate'];

/** ===== Entry point ===== */
function normalizeAllDates() {
  const ss = SpreadsheetApp.getActive();
  const log = NORM_prepareLog_(ss);

  const results = [];
  results.push(NORM_normalizeSheetDates_(ss, NORM_SHEET_PEOPLE,   NORM_COLS_PEOPLE,   log));
  results.push(NORM_normalizeSheetDates_(ss, NORM_SHEET_MARRIAGE, NORM_COLS_MARRIAGE, log));
  results.push(NORM_normalizeSheetDates_(ss, NORM_SHEET_CHILDREN, NORM_COLS_CHILDREN, log));

  const ui = SpreadsheetApp.getUi();
  const summary = results.map(r =>
    `${r.sheet}: ${r.changed} changed / ${r.rows} rows${r.approx ? ` (approx ${r.approx} year-only)` : ''}`
  ).join('\n');
  ui.alert('Normalization complete', summary, ui.ButtonSet.OK);
}

/** ===== Core: normalize specified date columns in a sheet ===== */
function NORM_normalizeSheetDates_(ss, sheetName, dateHeaders, log) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { sheet: sheetName, rows: 0, changed: 0, approx: 0 };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { sheet: sheetName, rows: 0, changed: 0, approx: 0 };

  // Header mapping
  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => String(h).trim());
  const hIndex = Object.fromEntries(headers.map((h,i) => [h, i])); // header -> 0-based index

  const missing = dateHeaders.filter(h => !(h in hIndex));
  if (missing.length) {
    NORM_logLine_(log, sheetName, '', `SKIP: missing columns ${missing.join(', ')}`);
    return { sheet: sheetName, rows: lastRow-1, changed: 0, approx: 0 };
  }

  const nRows = lastRow - 1;
  const maxRows = sh.getMaxRows();

  // PASS 0: Pre-force entire target columns to plain text (rows 2..maxRows)
  dateHeaders.forEach(h => {
    const col = hIndex[h] + 1; // 1-based
    sh.getRange(2, col, Math.max(0, maxRows - 1), 1).setNumberFormat('@');
  });

  // Read raw + display
  const dataRaw  = sh.getRange(2,1,nRows,lastCol).getValues();          // typed
  const dataDisp = sh.getRange(2,1,nRows,lastCol).getDisplayValues();   // strings (for logging)

  let changed = 0, approxCount = 0;
  const colIdx = dateHeaders.map(h => hIndex[h]);

  // PASS 1: Convert and write back as strings
  for (let r = 0; r < nRows; r++) {
    let rowChanged = false;
    for (const idx of colIdx) {
      const rawVal  = dataRaw[r][idx];
      const dispVal = (dataDisp[r][idx] || '').trim();

      let conv = { value: '', approx: false };
      if (rawVal instanceof Date) {
        conv.value = NORM_dateObjToISO_(rawVal);
      } else if (typeof rawVal === 'string') {
        conv = NORM_toISO_(rawVal);
      } else {
        conv.value = '';
      }
      if (conv.approx) approxCount++;

      const currentString = (typeof rawVal === 'string') ? rawVal.trim() : dispVal;
      if (conv.value !== currentString) {
        dataRaw[r][idx] = conv.value; // ensure string
        rowChanged = true;
        if (currentString) {
          NORM_logLine_(log, sheetName, r+2,
            `Converted "${currentString}" → "${conv.value || '(blank)'}"${conv.approx ? ' [approx]' : ''}`);
        }
      }
    }
    if (rowChanged) changed++;
  }

  if (changed) {
    sh.getRange(2,1,nRows,lastCol).setValues(dataRaw);
  }

  // PASS 1.5: Force plain text again to quash any auto-detect
  dateHeaders.forEach(h => {
    const col = hIndex[h] + 1;
    sh.getRange(2, col, Math.max(0, maxRows - 1), 1).setNumberFormat('@');
  });
  SpreadsheetApp.flush();

  // PASS 2: Scrub any stragglers that still display like dates (e.g., "1/1/1939")
  // Re-read typed + display; if typed is Date OR display looks like m/d/yyyy, overwrite with ISO string as text.
  const chkRaw  = sh.getRange(2,1,nRows,lastCol).getValues();
  const chkDisp = sh.getRange(2,1,nRows,lastCol).getDisplayValues();

  let scrubbed = 0;
  for (let r = 0; r < nRows; r++) {
    for (const idx of colIdx) {
      const typed = chkRaw[r][idx];
      const disp  = (chkDisp[r][idx] || '').trim();

      // If it's still a Date object, force to ISO string
      if (typed instanceof Date) {
        chkRaw[r][idx] = NORM_dateObjToISO_(typed);
        scrubbed++;
        continue;
      }

      // If display looks like m/d/yyyy (or d/m/yyyy), coerce to ISO if we can parse it
      // We'll try to parse with Utilities first; if it fails, leave it.
      if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(disp)) {
        try {
          const parts = disp.split('/');
          const mm = parts[0].padStart(2,'0');
          const dd = parts[1].padStart(2,'0');
          const yy = parts[2];
          chkRaw[r][idx] = `${yy}-${mm}-${dd}`;
          scrubbed++;
        } catch (e) {
          // ignore and leave as-is
        }
      }
    }
  }

  if (scrubbed) {
    sh.getRange(2,1,nRows,lastCol).setValues(chkRaw);
    // Final plain-text enforce
    dateHeaders.forEach(h => {
      const col = hIndex[h] + 1;
      sh.getRange(2, col, Math.max(0, maxRows - 1), 1).setNumberFormat('@');
    });
  }

  return { sheet: sheetName, rows: nRows, changed: changed + scrubbed, approx: approxCount };
}

/** ===== Robust string → ISO converter =====
 * Handles:
 *  - Already ISO
 *  - dd mmm yyyy   (14 Mar 1863)
 *  - dd, mmm, yyyy (14, Mar, 1863)
 *  - Month dd yyyy (March 14 1863)
 *  - extra spaces, NBSPs, optional commas
 *  - Year-only (per NORMALIZE_MODE)
 */
function NORM_toISO_(s) {
  const out = { value: '', approx: false };
  if (s == null) return out;

  // Normalize whitespace & commas
  let t = String(s)
    .replace(/\u00A0/g, ' ')   // NBSP → space
    .replace(/\s+/g, ' ')      // collapse spaces
    .replace(/\s*,\s*/g, ',')  // trim comma spacing
    .trim();

  if (!t || /^unknown$/i.test(t)) return out;

  // Already ISO?
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) { out.value = t; return out; }

  const monthNum = (m) => {
    if (!m) return '';
    const key = m.toLowerCase().slice(0,3); // accept full names ("September") and "Sept"
    const map = {
      jan:'01', feb:'02', mar:'03', apr:'04', may:'05', jun:'06',
      jul:'07', aug:'08', sep:'09', oct:'10', nov:'11', dec:'12'
    };
    return map[key] || '';
  };

  // 1) dd Mmm yyyy      ("14 Mar 1863" or "14 Mar, 1863")
  let m = t.match(/^(\d{1,2})\s+([A-Za-z]+),?\s+(\d{4})$/);
  if (m) {
    const d  = m[1].padStart(2,'0');
    const mo = monthNum(m[2]);
    const y  = m[3];
    if (mo) { out.value = `${y}-${mo}-${d}`; return out; }
  }

  // 2) dd, Mmm, yyyy    ("14, Mar, 1863")
  m = t.match(/^(\d{1,2}),\s*([A-Za-z]+),\s*(\d{4})$/);
  if (m) {
    const d  = m[1].padStart(2,'0');
    const mo = monthNum(m[2]);
    const y  = m[3];
    if (mo) { out.value = `${y}-${mo}-${d}`; return out; }
  }

  // 3) Month dd yyyy    ("March 14 1863" or "March 14, 1863")
  m = t.match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const mo = monthNum(m[1]);
    const d  = m[2].padStart(2,'0');
    const y  = m[3];
    if (mo) { out.value = `${y}-${mo}-${d}`; return out; }
  }

  // 4) Year-only
  m = t.match(/^(\d{4})$/);
  if (m) {
    if (NORMALIZE_MODE === 'APPROX') {
      out.value = `${m[1]}-01-01`;
      out.approx = true;
    }
    return out;
  }

  // Unrecognized -> blank
  return out;
}

/** Date object -> ISO YYYY-MM-DD (UTC) */
function NORM_dateObjToISO_(d) {
  const y = Utilities.formatDate(d, 'UTC', 'yyyy');
  const m = Utilities.formatDate(d, 'UTC', 'MM');
  const day = Utilities.formatDate(d, 'UTC', 'dd');
  return `${y}-${m}-${day}`;
}

/** ===== Logging ===== */
function NORM_prepareLog_(ss) {
  const old = ss.getSheetByName('Normalization_Log');
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet('Normalization_Log');
  sh.getRange(1,1,1,4).setValues([['Sheet','Row','Message','Mode']]);
  sh.getRange('A:D').setWrap(true);
  return sh;
}
function NORM_logLine_(logSh, sheet, row, message) {
  const mode = (NORMALIZE_MODE === 'APPROX') ? 'APPROX' : 'STRICT';
  logSh.appendRow([sheet, row || '', String(message), mode]);
}
