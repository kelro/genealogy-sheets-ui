/** ==============================================
 * PRODUCTION IMPORTER (no menu)
 * - Run once from Apps Script editor:
 *     Run ▶︎ runProductionImportFromPrompt
 * - Prompts for CSV (Drive) URL/ID
 * - Auto BACKUP of current spreadsheet
 * - Clears & rewrites People / Marriages / Children
 * - Logs progress to Import_Log
 * ============================================== */

// Use unique constant names so they don't conflict with the dry-run importer
const PROD_HEADERS_PEOPLE   = ['PersonID','FullName','DateOfBirth','PlaceOfBirth','DateOfDeath','PlaceOfDeath','Parents','Notes','Timestamp'];
const PROD_HEADERS_MARRIAGE = ['MarriageID','PersonID','SpouseName','MarriageDate','MarriagePlace','Status','Timestamp'];
const PROD_HEADERS_CHILDREN = ['ChildID','PersonID','MarriageID','ChildName','BornDate','BornPlace','DiedDate','DiedPlace','Notes','Timestamp'];

/** ======== ENTRYPOINT (Prompt) ======== */
function runProductionImportFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Legacy CSV on Google Drive',
    'Paste the Google Drive file URL or ID for your legacy CSV (tip: use "Manage versions" to keep the same ID for updates).',
    ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const fileId = extractDriveFileId_(p.getResponseText());
  if (!fileId) { ui.alert('Could not extract a Google Drive file ID from your input.'); return; }

  const ok = ui.alert(
    'Final Confirmation',
    'This will CLEAR and REWRITE People, Marriages, and Children in THIS spreadsheet.\n\nA full backup copy will be created first.',
    ui.ButtonSet.OK_CANCEL
  );
  if (ok !== ui.Button.OK) return;

  runProductionImport_(fileId);
}

/** ======== CORE (no UI prompts) ======== */
function runProductionImport_(fileId) {
  const ss = SpreadsheetApp.getActive();
  const logSh = ensureLogSheet_(ss);
  log_(logSh, '=== PRODUCTION IMPORT START ===');

  // Backup
  const backupUrl = makeBackupCopy_();
  log_(logSh, 'Backup copy created: ' + backupUrl);

  try {
    // Read CSV + meta (nice for auditing)
    const { rows, meta, preview } = readCsvWithMeta_(fileId);
    log_(logSh, `Reading CSV: ${meta.name} (ID ${meta.id})`);
    log_(logSh, `Last modified: ${meta.updated} • Size: ${meta.size} bytes`);
    log_(logSh, `Preview line 1: ${preview[0] || ''}`);
    log_(logSh, `Preview line 2: ${preview[1] || ''}`);

    if (!rows || !rows.length) throw new Error('CSV appears empty.');
    const header = rows[0].map(h => String(h || '').trim());
    const data   = rows.slice(1);

    const idx = colIndexMap_(header, ['generation','name','meta','spouse','children']);
    if (idx.name === -1 || idx.meta === -1 || idx.spouse === -1) {
      throw new Error('CSV must include at least: name, meta, spouse. (children/generation optional)');
    }

    // Prepare target sheets (clear + headers)
    const peopleSh   = truncateAndEnsure_(ss, 'People',    HEADERS_PEOPLE);
    const marriageSh = truncateAndEnsure_(ss, 'Marriages', HEADERS_MARRIAGE);
    const childrenSh = truncateAndEnsure_(ss, 'Children',  HEADERS_CHILDREN);

    const now = nowString_();
    const warn = [];
    const peopleRows=[], marriageRows=[], childRows=[];

    for (let i=0;i<data.length;i++){
      const r = data[i], rowNum = i+2;
      const generation = pick_(r, idx.generation);
      const name       = pick_(r, idx.name);
      const metaCell   = pick_(r, idx.meta);
      const spouseCell = pick_(r, idx.spouse);
      const childrenCell = pick_(r, idx.children);

      if (!name){ warn.push(`Row ${rowNum}: missing name; skipped.`); continue; }

      const { dob, pob, dod, pod } = parseMeta_(metaCell);
      const personId = uuid_();
      const notes = appendGenerationToNotes_(generation, '');

      peopleRows.push([personId, name, dob, pob, dod, pod, '', notes, now]);

      parseSpouses_(spouseCell).forEach(m => {
        marriageRows.push([uuid_(), personId, m.spouseName||'', m.mDate||'', m.mPlace||'', m.status||'', now]);
      });

      parseChildren_(childrenCell).forEach(c => {
        childRows.push([uuid_(), personId, '', c.childName||'', c.bDate||'', c.bPlace||'', c.dDate||'', c.dPlace||'', '', now]);
      });

      if ((i+1)%50===0) log_(logSh, `Processed ${i+1} rows…`);
    }

    // Bulk write
    log_(logSh, `Writing People (${peopleRows.length})…`);     fastAppend_(peopleSh, peopleRows);
    log_(logSh, `Writing Marriages (${marriageRows.length})…`); fastAppend_(marriageSh, marriageRows);
    log_(logSh, `Writing Children (${childRows.length})…`);    fastAppend_(childrenSh, childRows);

    if (warn.length) writeWarnings_(logSh, warn);
    log_(logSh, '=== IMPORT COMPLETE ===');
    SpreadsheetApp.getUi().alert(
      'Production Import Complete',
      `People: ${peopleRows.length}\nMarriages: ${marriageRows.length}\nChildren: ${childRows.length}\n\nBackup: ${backupUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (e) {
    log_(logSh, 'ERROR: ' + e.message);
    SpreadsheetApp.getUi().alert('Import failed: ' + e.message + '\nBackup is safe: ' + backupUrl);
  }
}

/** ======== Helpers: Drive + Logging + Sheets ======== */
function extractDriveFileId_(input) {
  const s = String(input || '').trim();
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : '';
}
function readCsvWithMeta_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const meta = { id:file.getId(), name:file.getName(), updated:file.getLastUpdated(), size:file.getSize() };
  let text;
  try { text = file.getBlob().getDataAsString('UTF-8'); }
  catch (e) { text = file.getBlob().getDataAsString(); }
  const rows = Utilities.parseCsv(text);
  return { rows, meta, preview: text.split(/\r?\n/).slice(0, 2) };
}
function ensureLogSheet_(ss) {
  const sh = ss.getSheetByName('Import_Log') || ss.insertSheet('Import_Log');
  sh.clear();
  sh.getRange(1,1,1,2).setValues([['Timestamp','Message']]);
  return sh;
}
function log_(logSh, msg) { logSh.appendRow([nowString_(), msg]); }
function writeWarnings_(logSh, warnArr) {
  logSh.appendRow(['', '--- Warnings ---']);
  const rows = warnArr.map(w => [nowString_(), w]);
  if (rows.length) logSh.getRange(logSh.getLastRow()+1, 1, rows.length, 2).setValues(rows);
}

function truncateAndEnsure_(ss, name, headers){
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function fastAppend_(sh, rows, chunkSize = 500) {
  if (!rows.length) return;
  const startRow = sh.getLastRow() + 1;
  let r = 0, at = startRow;
  while (r < rows.length) {
    const end = Math.min(r + chunkSize, rows.length);
    const block = rows.slice(r, end);
    sh.getRange(at, 1, block.length, block[0].length).setValues(block);
    r = end; at += block.length;
  }
}
function makeBackupCopy_() {
  const src = SpreadsheetApp.getActive();
  const name = src.getName() + ' (BACKUP ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm') + ')';
  const dstFile = DriveApp.getFileById(src.getId()).makeCopy(name);
  return 'https://docs.google.com/spreadsheets/d/' + dstFile.getId();
}

/** ======== Small utilities ======== */
function pick_(arr, idx){ return (idx >= 0 && idx < arr.length) ? String(arr[idx]).trim() : ''; }
function colIndexMap_(hdrs, names) {
  const lower = hdrs.map(h => h.toLowerCase()); const o = {};
  names.forEach(n => { o[n] = lower.indexOf(n); }); return o;
}
function uuid_() {
  const a='xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.split('');
  for (let i=0;i<a.length;i++){ const c=a[i];
    if (c==='x'||c==='y'){ const r=Math.random()*16|0; const v=(c==='x')?r:(r&0x3|0x8); a[i]=v.toString(16); }
  }
  return a.join('');
}
function nowString_(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone()||'America/Chicago', "yyyy-MM-dd'T'HH:mm:ss");
}
function appendGenerationToNotes_(generation, notes){
  const g = String(generation||'').trim();
  if (!g) return notes || '';
  const prefix = notes ? (notes + ' ') : '';
  return prefix + '(Generation: ' + g + ')';
}

/** ======== Parsers (same as dry run) ======== */
// META: "Born: <date>, city, county, state, nation • Died: <date>, city, county, state, nation"
function parseMeta_(metaRaw) {
  const out = { dob: '', pob: '', dod: '', pod: '' };
  if (!metaRaw) return out;
  const parts = String(metaRaw).split(/•/).map(s => s.trim());
  function parseSide(side) {
    const s = side.replace(/^(Born|Died)\s*:\s*/i, '').trim();
    const tokens = s.split(',').map(t => t.trim());
    const [date, city, county, state, nation] = [tokens[0]||'', tokens[1]||'', tokens[2]||'', tokens[3]||'', tokens[4]||''];
    const dateOut = /^unknown$/i.test(date) ? '' : date;
    const place = [city, county, state, nation].filter(v => v && !/^unknown$/i.test(v)).join(', ');
    return { date: dateOut, place };
  }
  const bornSeg = parts.find(p => /^Born\s*:/i.test(p));
  const diedSeg = parts.find(p => /^Died\s*:/i.test(p));
  if (bornSeg){ const b = parseSide(bornSeg); out.dob=b.date; out.pob=b.place; }
  if (diedSeg){ const d = parseSide(diedSeg); out.dod=d.date; out.pod=d.place; }
  return out;
}
// SPOUSE → [{spouseName, spouseBirth, spouseDeath, mDate, mPlace, status}]
function parseSpouses_(spouseRaw) {
  if (!spouseRaw) return [];
  const text = String(spouseRaw).trim()
    .replace(/\s+Spouse:/g, ' • Spouse:')
    .replace(/\s*\.\s*Spouse:/g, ' • Spouse:')
    .replace(/\.\s*$/,'');
  const blocks = text.split(/•/).map(s => s.trim()).filter(Boolean);
  const out = [];
  blocks.forEach(block => {
    if (!/^Spouse\s*:/i.test(block)) return;
    const segs = block.split(/\s—\s/).map(s => s.trim()).filter(Boolean);
    const head = segs.shift() || '';
    const mHead = head.match(/^Spouse\s*:\s*(.+?)\s*\((.*?)\)\s*$/i) || head.match(/^Spouse\s*:\s*(.+?)\s*$/i);
    let spouseName='', spouseBirth='', spouseDeath='';
    if (mHead) {
      spouseName = (mHead[1]||'').trim();
      const years = (mHead[2]||'').trim();
      if (years) {
        const yr = years.replace(/\u2013/g,'-').replace(/[^\d\-]/g,'');
        const parts = yr.split('-'); spouseBirth=(parts[0]||'')||''; spouseDeath=(parts[1]||'')||'';
      }
    }
    const marriedSegs = []; let divorceInfo = '';
    segs.forEach(s => {
      if (/^Married\s*:/i.test(s)) marriedSegs.push(s);
      else if (/^Divorced\s*:/i.test(s)) divorceInfo = s.replace(/^Divorced\s*:\s*/i,'').trim();
    });
    let mDate='', mPlace='';
    if (marriedSegs.length) {
      const parsed = marriedSegs.map(ms => {
        const body = ms.replace(/^Married\s*:\s*/i,'').trim().replace(/\.$/,'');
        const t = body.split(',').map(x => x.trim());
        const date = /^unknown$/i.test(t[0]||'') ? '' : (t[0]||'');
        const parts = [t[1]||'', t[2]||'', t[3]||'', t[4]||''].filter(v => v && !/^unknown$/i.test(v));
        return { date, place: parts.join(', '), score: parts.length };
      });
      parsed.sort((a,b) => (b.score - a.score) || (b.date ? 1 : -1));
      const best = parsed[0] || { date:'', place:'' }; mDate=best.date; mPlace=best.place;
    }
    let status = ''; if (divorceInfo) status = 'Divorced ' + divorceInfo;
    out.push({ spouseName, spouseBirth, spouseDeath, mDate, mPlace, status });
  });
  return out;
}
// CHILDREN → [{childName, bDate, dDate, bPlace, dPlace}]
function parseChildren_(childrenRaw) {
  if (!childrenRaw) return [];
  const text = String(childrenRaw).trim(); if (!text) return [];
  const entries = text.split(/;|\n/).map(s => s.trim()).filter(Boolean);
  return entries.map(entry => {
    const m = entry.match(/^(.+?)\s*\(([^)]+)\)\s*(.*)?$/);
    let childName='', datePart='', trailing='';
    if (m) { childName=m[1].trim(); datePart=m[2].trim(); trailing=(m[3]||'').trim().replace(/\.$/,''); }
    else   { return { childName: entry, bDate:'', dDate:'', bPlace:'', dPlace:'' }; }
    const norm = datePart.replace(/\u2013/g,'-').replace(/\s+/g,'');
    let bDate='', dDate='';
    if (/^[-–]$/.test(datePart)) { /* both unknown */ }
    else if (/^\d{4}$/.test(norm) || /^[0-3]?\d[A-Za-z]{3}\s+\d{4}$/.test(datePart)) { bDate=datePart; }
    else if (/^\d{4}-\d{4}$/.test(norm)) { const [b,d]=norm.split('-'); bDate=b; dDate=d; }
    else if (/^\d{4}-$/.test(norm)) { bDate=norm.slice(0,4); }
    else if (/^-\d{4}$/.test(norm)) { dDate=norm.slice(1); }
    else { const parts=norm.split('-'); if (parts.length===2){ bDate=parts[0]||''; dDate=parts[1]||''; } else { bDate=datePart; } }
    let bPlace='', dPlace=''; if (trailing && !/^unknown$/i.test(trailing)) bPlace=trailing;
    return { childName, bDate, dDate, bPlace, dPlace };
  });
}
