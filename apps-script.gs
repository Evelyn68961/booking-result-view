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
 * - Browsers block fetch() from file:// to https URLs, so index.html must
 *   be served over http(s) — GitHub Pages / Netlify / `python -m http.server`
 *   all work.
 */

const SHEET_NAME = 'Bookings';
const HEADERS = ['id', 'name', 'status', 'start', 'end', 'daysCat', 'submittedAt', 'source', 'deleted'];

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
  return jsonOut({ ok: true, records: records });
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

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
