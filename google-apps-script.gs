/**
 * Shallow Well Survey — Google Sheets receiver + dashboard data feed
 * (multi-user, no duplicates, baseline/after rounds, review from the PC dashboard)
 *
 * HOW TO INSTALL
 * 1. Open the Google Sheet → Extensions → Apps Script.
 * 2. Delete the old code, paste ALL of this file.
 * 3. Change DASHBOARD_KEY below to your own secret word (needed to open the PC dashboard).
 * 4. Save → Deploy → New deployment → type "Web app"
 *      Execute as: Me        Who has access: Anyone
 *    Deploy and allow the permissions (Sheets + Drive for photos).
 * 5. Copy the Web App URL (ends in /exec):
 *      - phones: Settings → Google Apps Script Web App URL
 *      - PC dashboard: Connect Sheet → same URL + your DASHBOARD_KEY
 *    When you edit this code later: Deploy → Manage deployments → Edit → Version: New version
 *    (the URL stays the same).
 *
 * WHY THERE ARE NO DUPLICATES
 * - Every record gets a permanent Record ID on the phone (device ID + time + random).
 * - This script looks up the Record ID before writing: same ID → update that row,
 *   new ID → add a row. Re-sending is harmless.
 * - A lock makes simultaneous uploads from many phones wait their turn.
 * - The phone marks a record "uploaded" only after this script confirms it.
 *
 * BASELINE / AFTER
 * - "Survey Round" is Baseline (first visit) or After (return visit to the same well).
 * - "Possible Duplicate Of" only flags records of the SAME round that look like the same
 *   well (within 10 m, or same interviewee in the same barangay). A Baseline and an After
 *   visit at the same spot are a before/after pair, not duplicates.
 *
 * PRIVACY
 * - Phones can only ADD/UPDATE their own records. Reading all records (dashboard) and
 *   saving reviews requires DASHBOARD_KEY.
 */

const DASHBOARD_KEY = 'CHANGE-THIS-KEY';
const SHEET_NAME = 'Survey Records';
const PHOTO_FOLDER_NAME = 'Well Survey Photos';
const DUPLICATE_DISTANCE_M = 10;

const APP_HEADERS = [
  'Record ID', 'Surveyor', 'Device ID', 'Saved On Phone', 'Last Uploaded',
  'District', 'Barangay', 'Zone', 'Latitude X', 'Latitude Y', 'Interviewee', 'Interview Date/Time', 'Survey Round',
  'Used by Family (Owned)', 'Used by Family & Residents', 'Population', 'Well Depth (Tabay)',
  'Normal Level', 'Current Level', 'Other Water Sources', 'Water Use',
  'Volume per Well (L)', 'Daily Need (L)', 'Current Supply (L)', 'Available Supply (L)',
  'Recharge (Tabay)', 'Recharge (L)', 'Status', 'Affected', 'Remarks', 'Photo',
  'Possible Duplicate Of'
];
const REVIEW_HEADERS = ['Checked By', 'Rating (1-5)', 'Checker Remarks', 'Checked On'];
const ALL_HEADERS = APP_HEADERS.concat(REVIEW_HEADERS);

/* ---------------- GET: connection test / dashboard data ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const ctx = getSheet_();
  const count = Math.max(ctx.sheet.getLastRow() - 1, 0);
  if (p.action !== 'records') return json_({ ok: true, sheet: SHEET_NAME, records: count });
  if (String(p.key || '') !== DASHBOARD_KEY) return json_({ ok: false, error: 'Wrong dashboard key.' });
  const values = count ? ctx.sheet.getRange(2, 1, count, ctx.width).getValues() : [];
  const records = values.filter(r => r[ctx.col['Record ID'] - 1]).map(r => {
    const o = {};
    ctx.headers.forEach((h, i) => { const v = r[i]; o[h] = v instanceof Date ? v.toISOString() : v; });
    return o;
  });
  return json_({ ok: true, sheet: SHEET_NAME, records: records, headers: ctx.headers, generatedAt: new Date().toISOString() });
}

/* ---------------- POST: upload records / save review ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ ok: false, error: 'Server busy, try again.' });
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'review') return saveReview_(body);
    return saveRecords_(body);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveRecords_(body) {
  const incoming = Array.isArray(body.records) ? body.records : [];
  const ctx = getSheet_();
  const sheet = ctx.sheet, col = ctx.col, width = ctx.width;
  const now = new Date();
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
  const rowById = {};
  rows.forEach((row, i) => { if (row[col['Record ID'] - 1]) rowById[String(row[col['Record ID'] - 1])] = i; });

  const saved = [], appendRows = [], touched = [];
  incoming.forEach(rec => {
    const id = String(rec.id || legacyId_(rec));
    const idx = rowById[id];
    const row = idx !== undefined ? rows[idx].slice() : new Array(width).fill('');
    let photo = row[col['Photo'] - 1];
    if (!photo && rec.photoData) photo = savePhoto_(id, rec.photoData);
    else if (!photo) photo = rec.photoName || '';
    fillAppValues_(row, col, rec, id, body, now, photo);
    if (idx !== undefined) {
      rows[idx] = row;
      sheet.getRange(idx + 2, 1, 1, width).setValues([row]);
      touched.push(idx);
    } else {
      rowById[id] = rows.length;
      rows.push(row); appendRows.push(row); touched.push(rows.length - 1);
    }
    saved.push(id);
  });
  if (appendRows.length) sheet.getRange(lastRow + 1, 1, appendRows.length, width).setValues(appendRows);
  flagDuplicates_(sheet, col, rows, touched);
  return json_({ ok: true, saved: saved, total: rows.length });
}

function saveReview_(body) {
  if (String(body.key || '') !== DASHBOARD_KEY) return json_({ ok: false, error: 'Wrong dashboard key.' });
  const ctx = getSheet_();
  const sheet = ctx.sheet, col = ctx.col;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_({ ok: false, error: 'No records.' });
  const ids = sheet.getRange(2, col['Record ID'], lastRow - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(body.id));
  if (i < 0) return json_({ ok: false, error: 'Record not found: ' + body.id });
  const r = i + 2, now = new Date();
  const rating = body.rating ? String(Math.max(1, Math.min(5, Math.round(Number(body.rating))))) : '';
  sheet.getRange(r, col['Checked By']).setValue(body.checkedBy || '');
  sheet.getRange(r, col['Rating (1-5)']).setValue(rating);
  sheet.getRange(r, col['Checker Remarks']).setValue(body.remarks || '');
  sheet.getRange(r, col['Checked On']).setValue(now);
  return json_({ ok: true, id: String(body.id), checkedOn: now.toISOString() });
}

function fillAppValues_(v, col, r, id, body, now, photo) {
  const set = (h, x) => { v[col[h] - 1] = (x === undefined || x === null) ? '' : x; };
  set('Record ID', id);
  set('Surveyor', r.surveyor || body.surveyor || '');
  set('Device ID', r.deviceId || body.deviceId || '');
  set('Saved On Phone', r.savedAt || '');
  set('Last Uploaded', now);
  set('District', r.district); set('Barangay', r.barangay); set('Zone', r.zone);
  set('Latitude X', r.latX); set('Latitude Y', r.latY);
  set('Interviewee', r.interviewee); set('Interview Date/Time', r.interviewDateTime);
  set('Survey Round', r.surveyRound || 'Baseline');
  set('Used by Family (Owned)', r.usedByFamily); set('Used by Family & Residents', r.usedByFamilyResidents);
  set('Population', r.population); set('Well Depth (Tabay)', r.wellDepth);
  set('Normal Level', r.normalLevel); set('Current Level', r.currentLevel);
  set('Other Water Sources', r.otherWaterSource); set('Water Use', r.waterUse);
  set('Volume per Well (L)', r.volumePerWell); set('Daily Need (L)', r.dailyNeed);
  set('Current Supply (L)', r.currentSupply); set('Available Supply (L)', r.availableSupply);
  set('Recharge (Tabay)', r.wellRechargeTabay); set('Recharge (L)', r.wellRechargeLiters);
  set('Status', r.status); set('Affected', r.affected); set('Remarks', r.remarks);
  set('Photo', photo);
}

// Records from older app versions have no ID: build a stable one from their content.
function legacyId_(r) {
  const key = [r.district, r.barangay, r.zone, r.interviewee, r.interviewDateTime, r.latX, r.latY].join('|');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key);
  return 'legacy-' + bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 16);
}

function savePhoto_(id, dataUrl) {
  try {
    const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
    if (!m) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '.jpg');
    const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
    return folder.createFile(blob).getUrl();
  } catch (err) {
    return 'Photo upload failed: ' + err;
  }
}

function flagDuplicates_(sheet, col, rows, touched) {
  if (!touched.length) return;
  const c = h => col[h] - 1;
  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const round = r => /after|follow/i.test(String(r[c('Survey Round')])) ? 'After' : 'Baseline';
  const lat = r => parseFloat(r[c('Latitude X')]), lon = r => parseFloat(r[c('Latitude Y')]);
  const dist = (a, b) => {
    const R = 6371000, rad = x => x * Math.PI / 180;
    const dLat = rad(lat(b) - lat(a)), dLon = rad(lon(b) - lon(a));
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(rad(lat(a))) * Math.cos(rad(lat(b))) * Math.sin(dLon / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  };
  const ID = c('Record ID'), DUP = c('Possible Duplicate Of');
  const changed = new Set();
  touched.forEach(i => {
    const a = rows[i];
    const matches = [];
    rows.forEach((b, j) => {
      if (j === i || !b[ID] || round(a) !== round(b)) return;
      const near = isFinite(lat(a)) && isFinite(lon(a)) && isFinite(lat(b)) && isFinite(lon(b)) && dist(a, b) <= DUPLICATE_DISTANCE_M;
      const sameName = norm(a[c('Interviewee')]) && norm(a[c('Interviewee')]) === norm(b[c('Interviewee')]) &&
        norm(a[c('Barangay')]) === norm(b[c('Barangay')]);
      if (near || sameName) {
        matches.push(b[ID]);
        const other = String(b[DUP] || '');
        if (other.split(', ').indexOf(String(a[ID])) === -1) {
          b[DUP] = other ? other + ', ' + a[ID] : String(a[ID]);
          changed.add(j);
        }
      }
    });
    a[DUP] = matches.join(', ');
    changed.add(i);
  });
  changed.forEach(i => sheet.getRange(i + 2, DUP + 1).setValue(rows[i][DUP]));
}

// Creates the sheet if needed and adds any missing columns at the end,
// so an existing sheet from an older version keeps working.
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME), created = false;
  if (!sheet) {
    created = true;
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, ALL_HEADERS.length).setValues([ALL_HEADERS]).setFontWeight('bold').setBackground('#d1fae5');
    sheet.setFrozenRows(1);
  }
  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  const missing = ALL_HEADERS.filter(h => headers.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#d1fae5');
    headers = headers.concat(missing);
  }
  const col = {};
  headers.forEach((h, i) => { if (h && !col[h]) col[h] = i + 1; });
  if (missing.length || created) {
    sheet.getRange(1, col['Possible Duplicate Of']).setBackground('#fef3c7');
    REVIEW_HEADERS.forEach(h => sheet.getRange(1, col[h]).setBackground('#dbeafe'));
  }
  return { sheet: sheet, col: col, headers: headers, width: headers.length };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
