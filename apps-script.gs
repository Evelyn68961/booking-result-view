/**
 * Google Apps Script backend for the booking-result-view sync feature.
 *
 * SETUP
 * 1. Create or open a Google Sheet that you and your managers will share.
 * 2. Extensions → Apps Script. Replace the default code with this file. Save.
 * 3. Deploy → New deployment → type: Web app.
 *      Execute as: Me
 *      Who has access: Anyone (or "Anyone with Google account" if you prefer)
 * 4. Copy the deployment URL (ends with .../exec).
 * 5. In index.html: unlock the manager tab, paste the URL into the
 *    "Google Sheet 同步" panel, click 上傳到 Sheet.
 * 6. Share the Google Sheet itself (the spreadsheet URL, not the script URL)
 *    with your managers so they can view / edit cells directly.
 *
 * NOTES
 * - The "Bookings" sheet is created on first push. Don't rename it.
 * - Don't edit the `id` column — it's how the app maps a sheet row back to
 *   a baked record (b:N) or a manager record (m:_id). Renaming an id will
 *   create a duplicate on the next pull.
 * - The `deleted` column accepts TRUE / FALSE. TRUE hides that record in
 *   the app on next pull (for baked rows it becomes a tombstone patch;
 *   for manager rows it removes them from local storage).
 * - The "QuotaConfig" sheet is created on first push and stores the per-day
 *   limit settings. Each row has columns kind/from/to/quota/note/ts.
 *   `kind=default, from=weekday|weekend, quota=N` sets the weekday/weekend defaults;
 *   `kind=override, from=YYYY-MM-DD, to=YYYY-MM-DD, quota=N, note=...` adds a
 *   range override. The whole sheet is rewritten on every push, so direct
 *   edits get overwritten by the next sync — change the limits in the app's
 *   "上限例外" panel.
 * - Browsers block fetch() from file:// to https URLs, so index.html must
 *   be served over http(s) — GitHub Pages / Netlify / `python -m http.server`
 *   all work.
 */

const SHEET_NAME = 'Bookings';
const HEADERS = ['id', 'name', 'status', 'start', 'end', 'daysCat', 'submittedAt', 'source', 'deleted'];
const QUOTA_SHEET_NAME = 'QuotaConfig';
const QUOTA_HEADERS = ['kind', 'from', 'to', 'quota', 'note', 'ts'];

function doGet(e) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const records = [];
  if (values.length >= 2) {
    const hdr = values[0].map(String);
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row.every(v => v === '' || v === null)) continue;
      const o = {};
      hdr.forEach((h, j) => {
        let v = row[j];
        if (v instanceof Date) v = Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
        o[h] = v;
      });
      records.push(o);
    }
  }
  return jsonOut({ ok: true, records: records, quotaConfig: readQuotaConfig() });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (body.action === 'overwrite') {
    const sheet = getSheet();
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    const records = body.records || [];
    if (records.length) {
      const rows = records.map(r => HEADERS.map(h => {
        const v = r[h];
        return v === undefined || v === null ? '' : v;
      }));
      sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
      sheet.getRange(2, 1, rows.length, HEADERS.length).setNumberFormat('@'); // text format, prevents auto date conversion
    }
    if (body.quotaConfig) writeQuotaConfig(body.quotaConfig);
    return jsonOut({ ok: true, count: records.length, ts: new Date().toISOString() });
  }
  return jsonOut({ ok: false, error: 'unknown action: ' + body.action });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getQuotaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUOTA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUOTA_SHEET_NAME);
    sheet.getRange(1, 1, 1, QUOTA_HEADERS.length).setValues([QUOTA_HEADERS]);
    sheet.getRange(1, 1, 1, QUOTA_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readQuotaConfig() {
  const sheet = getQuotaSheet();
  const values = sheet.getDataRange().getValues();
  const out = { weekday: 2, weekend: 4, overrides: [] };
  if (values.length < 2) return out;
  const hdr = values[0].map(String);
  const col = name => hdr.indexOf(name);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(v => v === '' || v === null)) continue;
    const kind = String(row[col('kind')] || '').toLowerCase();
    const fromRaw = row[col('from')];
    const toRaw   = row[col('to')];
    const quota   = Number(row[col('quota')]);
    const note    = String(row[col('note')] || '');
    const ts      = Number(row[col('ts')]) || 0;
    const fmt = v => v instanceof Date ? Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd') : String(v || '');
    if (kind === 'default') {
      const which = fmt(fromRaw).toLowerCase();
      if (which === 'weekday' && Number.isFinite(quota)) out.weekday = Math.max(0, quota);
      if (which === 'weekend' && Number.isFinite(quota)) out.weekend = Math.max(0, quota);
    } else if (kind === 'override') {
      const from = fmt(fromRaw), to = fmt(toRaw);
      if (from && to && Number.isFinite(quota)) {
        out.overrides.push({ id: 'sheet-' + i, from, to, quota: Math.max(0, quota), note, ts });
      }
    }
  }
  return out;
}

function writeQuotaConfig(qc) {
  const sheet = getQuotaSheet();
  sheet.clear();
  sheet.getRange(1, 1, 1, QUOTA_HEADERS.length).setValues([QUOTA_HEADERS]);
  sheet.getRange(1, 1, 1, QUOTA_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rows = [];
  rows.push(['default', 'weekday', '', Number(qc.weekday) || 0, '', '']);
  rows.push(['default', 'weekend', '', Number(qc.weekend) || 0, '', '']);
  for (const o of (qc.overrides || [])) {
    rows.push(['override', String(o.from || ''), String(o.to || ''), Number(o.quota) || 0, String(o.note || ''), Number(o.ts) || 0]);
  }
  sheet.getRange(2, 1, rows.length, QUOTA_HEADERS.length).setValues(rows);
  sheet.getRange(2, 1, rows.length, QUOTA_HEADERS.length).setNumberFormat('@');
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
