/**
 * RSVP receiver for the wedding website.
 * Paste this whole file into Extensions > Apps Script of your RSVP Google Sheet,
 * then deploy it as a Web app (see README.md).
 *
 * Security notes
 * - The web app can only ADD rows. It never reads, edits or deletes anything,
 *   and it returns nothing from the sheet, so the endpoint cannot leak guest data.
 * - Every text value is cleaned so a guest cannot inject a spreadsheet formula
 *   (e.g. "=IMPORTXML(...)" or "=HYPERLINK(...)") into your sheet.
 * - Only expected fields are stored, with length limits, fixed choices for
 *   "Attending" and a capped number for "Party size".
 * - A hidden spam-trap field and a rate limit block bots and floods.
 */
// Your "Wedding RSVPs" Google Sheet. This lets the script work whether it was
// opened from the sheet (Extensions > Apps Script) or created at script.google.com.
const SHEET_ID = '1o33bruZ8vaZGmCXZKFhdWUlvH6jqpeZ6OKE7X0sfqew';
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Received', 'Name', 'Attending', 'Party size', 'Guest names',
                 'Dietary requirements', 'Email', 'Message', 'Invite (from personal link)'];
const MAX_GUESTS = 12;
const MAX_PER_10_MIN = 40;   // all submissions combined; real guests will never hit this

/** Make any guest-supplied text safe to store as plain text. */
function clean(value, maxLen) {
  let s = String(value == null ? '' : value);
  s = s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');   // control characters
  s = s.replace(/\r\n?/g, '\n').trim().slice(0, maxLen);
  // Anything that a spreadsheet could read as a formula is forced to be text.
  // The leading apostrophe is hidden by Sheets and is not part of the value.
  if (/^[=+\-@\t\r\n＝＋－＠]/.test(s)) s = "'" + s;
  return s;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const p = (e && e.parameter) || {};

  if (p.website) return respond({ ok: true });                 // spam trap filled in: ignore quietly

  const name = clean(p.name, 120);
  const attending = p.attending === 'Yes' ? 'Yes' : p.attending === 'No' ? 'No' : '';
  if (!name || !attending) return respond({ ok: false, error: 'missing fields' });

  let guests = parseInt(p.guests, 10);
  if (attending === 'No' || !(guests >= 1)) guests = attending === 'Yes' ? 1 : 0;
  guests = Math.min(guests, MAX_GUESTS);

  // Simple flood protection
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get('rsvp-count') || 0);
  if (count >= MAX_PER_10_MIN) return respond({ ok: false, error: 'busy' });
  cache.put('rsvp-count', String(count + 1), 600);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      name,
      attending,
      guests,
      clean(p.guestNames, 400),
      clean(p.dietary, 600),
      clean(p.email, 200),
      clean(p.message, 1500),
      clean(p.invite, 120),
    ]);
    return respond({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Opening the Web App URL in a browser shows this, which is handy for checking it is live.
function doGet() {
  return ContentService.createTextOutput('RSVP endpoint is running.');
}

/**
 * Optional: run this once from the Apps Script editor (select "selfTest", click Run)
 * to check the cleaning works. It only logs results; it writes nothing to the sheet.
 */
function selfTest() {
  ['=IMPORTXML("http://evil","//a")', '+1+1', '-2', '@SUM(A1)', '  =1+1', '\t=1', 'Maria Santos']
    .forEach(function (v) { Logger.log(JSON.stringify(v) + ' -> ' + JSON.stringify(clean(v, 200))); });
}
