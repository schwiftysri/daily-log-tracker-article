/**
 * ================= CONFIG =================
 */
const SHEET_NAME = 'Logs';

// Change this to your own long random string. Your Shortcut must send the
// same value. This stops anyone who finds your web app URL from writing
// into your sheet.
const SECRET = 'PUT-A-LONG-RANDOM-PASSWORD-HERE';

// CONFIRMED BUG: Apps Script writes a raw `new Date()` into a Sheets cell
// using its UTC fields, not local/IST fields, even though
// Session.getScriptTimeZone() correctly reports Asia/Kolkata. Reproduced
// directly in the editor (bypassing the web app entirely), so it isn't a
// deployment issue. India has no daylight saving, so a fixed +5:30 offset
// is safe to hardcode. If you ever log from outside IST, this needs to
// become timezone-aware instead of a fixed constant.
const IST_OFFSET_MS = 5.5 * 60 * 60 * 1000;
function nowInIST() {
  return new Date(Date.now() + IST_OFFSET_MS);
}

/**
 * Every POST from either Shortcut lands here. Body shape:
 * {
 *   secret:      "...",
 *   type:        "task" | "retro",
 *
 *   // type = task
 *   task:        string,
 *   urgent:      "Y" | "N",
 *   important:   "Y" | "N",
 *   timeTaken:   number (minutes),
 *   mood:        "Positive" | "Neutral" | "Negative" | ""   (optional on a task)
 *   energy:      "Positive" | "Neutral" | "Negative" | ""   (optional on a task)
 *
 *   // type = retro
 *   mood:        "Positive" | "Neutral" | "Negative"        (required on a retro)
 *   energy:      "Positive" | "Neutral" | "Negative"        (required on a retro)
 *   note:        string (optional, either type)
 * }
 * Timestamp is set here, server-side, at the moment the row is appended —
 * you never send a time from the Shortcut.
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) {
      return jsonOut({ ok: false, error: 'bad secret' });
    }
    if (body.type !== 'task' && body.type !== 'retro') {
      return jsonOut({ ok: false, error: 'type must be "task" or "retro"' });
    }
    if (body.type === 'retro' && (!body.mood || !body.energy)) {
      return jsonOut({ ok: false, error: 'retro needs mood and energy' });
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return jsonOut({ ok: false, error: 'sheet not found: ' + SHEET_NAME });

    const timeTaken = (body.timeTaken !== undefined && body.timeTaken !== '')
      ? Number(body.timeTaken)
      : '';

    sheet.appendRow([
      nowInIST(),
      body.type,
      body.task      || '',
      body.urgent    || '',
      body.important || '',
      timeTaken,
      body.mood      || '',
      body.energy    || '',
      body.note      || ''
    ]);

    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
