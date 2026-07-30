// Google Apps Script — deploy as Web App
// 1. Open the "ICC Lesson Log Worksheet for App" Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it into the app's Settings
//
// Tabs are created automatically on first use — no manual sheet setup:
// - "Pros" tab: the pro list (seeded with the current roster). To add or
//   remove a pro, just edit the names in column A — the app's dropdown
//   updates on next load.
// - Monthly tabs ("July 2026", "August 2026", ...): one tab per calendar
//   month, auto-created when the first lesson of that month is logged. Each
//   row is a lesson with a "Charged" checkbox the shop manager ticks after
//   entering it in Jonas. The weekly reminder scans every month at once, so
//   nothing gets lost in an old tab.
//
// WEEKLY UNCHARGED REMINDER (optional — one-time setup):
//   After pasting this code, run the "setupWeeklyReminder" function once
//   (pick it from the toolbar dropdown and click Run, then authorize).
//   Every Monday morning it emails the people listed on the "Reminders"
//   tab about any lessons whose "Charged" box is still unchecked.
//   - "Reminders" tab columns: Name | Email | Send What
//   - Send What = "All lessons" (whole outstanding list, for the manager)
//     or "Only their own" (just that person's lessons, for each pro).
//   Add or change emails on that tab anytime — no code edits needed.

const HEADERS = ['Date', 'Pro', 'Client Name', 'Member/Guest', 'Duration', 'People', 'Notes', 'Charged'];
const PROS_SHEET_NAME = 'Pros';
const REMINDERS_SHEET_NAME = 'Reminders';
const CHARGED_COL = 8;
const GUEST_MEMBER_COL = 4;

// Lessons uncharged longer than this many days get a ⚠️ flag in the email.
const AGING_DAYS = 14;

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

const DEFAULT_PROS = [
  'J.C. Freeman',
  'Joey Francis',
  'A.B. Hill',
  'Will Davidson',
  'Matt Kendrick',
  'Lisa Webb',
  'Stephanie Heckler'
];

function getProsSheet_(ss) {
  var prosSheet = ss.getSheetByName(PROS_SHEET_NAME);
  if (!prosSheet) {
    prosSheet = ss.insertSheet(PROS_SHEET_NAME);
    prosSheet.appendRow(['Pro Names (edit anytime — the app updates on next load)']);
    prosSheet.getRange(1, 1).setFontWeight('bold');
    prosSheet.setFrozenRows(1);
    for (var i = 0; i < DEFAULT_PROS.length; i++) {
      prosSheet.appendRow([DEFAULT_PROS[i]]);
    }
    prosSheet.setColumnWidth(1, 400);
  }
  return prosSheet;
}

// Turns a lesson date "m/d/yy" into a monthly tab name like "July 2026".
// Falls back to the current month if the date can't be parsed.
function monthTabName_(dateStr) {
  var parts = String(dateStr).split('/');
  if (parts.length === 3) {
    var m = parseInt(parts[0], 10);
    var y = parts[2].length === 2 ? '20' + parts[2] : parts[2];
    if (m >= 1 && m <= 12) return MONTH_NAMES[m - 1] + ' ' + y;
  }
  var now = new Date();
  return MONTH_NAMES[now.getMonth()] + ' ' + now.getFullYear();
}

// Gets (or creates + formats) the monthly lesson tab.
function getOrCreateLessonSheet_(ss, tabName) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#1a1a2e');
    sheet.getRange(1, 1, 1, HEADERS.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// True if a tab looks like a lesson tab (matching header row). Lets the weekly
// reminder scan every monthly tab without caring what they're named.
function isLessonSheet_(sheet) {
  var name = sheet.getName();
  if (name === PROS_SHEET_NAME || name === REMINDERS_SHEET_NAME) return false;
  if (sheet.getLastColumn() < HEADERS.length || sheet.getLastRow() < 1) return false;
  var header = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  return header[0] === HEADERS[0] && header[CHARGED_COL - 1] === HEADERS[CHARGED_COL - 1];
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'getPros') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prosSheet = getProsSheet_(ss);

    var pros = [];
    var lastRow = prosSheet.getLastRow();
    if (lastRow > 1) {
      var values = prosSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < values.length; i++) {
        var name = values[i][0].toString().trim();
        if (name) pros.push(name);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', pros: pros }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput('ICC Lesson Log API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateLessonSheet_(ss, monthTabName_(data.date));

    sheet.appendRow([
      data.date,
      data.pro || 'Unknown',
      data.client,
      data.guestMember,
      data.duration,
      data.people || 1,
      data.notes || '',
      false
    ]);

    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, CHARGED_COL).insertCheckboxes();

    var gmCell = sheet.getRange(newRow, GUEST_MEMBER_COL);
    if (data.guestMember === 'GUEST') {
      gmCell.setBackground('#e74c3c');
      gmCell.setFontColor('#ffffff');
      gmCell.setFontWeight('bold');
    } else if (data.guestMember === 'MEMBER') {
      gmCell.setBackground('#2ecc71');
      gmCell.setFontColor('#ffffff');
      gmCell.setFontWeight('bold');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Weekly uncharged-lessons reminder
// ---------------------------------------------------------------------------

// Run this ONCE (from the Apps Script editor toolbar) to schedule the weekly
// email. Safe to run again — it clears any old copy of the schedule first.
// Also creates the "Reminders" tab now so you can fill in emails before the
// first send.
function setupWeeklyReminder() {
  getRemindersSheet_(SpreadsheetApp.getActiveSpreadsheet());

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendUnchargedDigest') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendUnchargedDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
}

// Turn a Date or "m/d/yy" string cell into a Date (or null if unparseable).
function coerceDate_(val) {
  if (val instanceof Date && !isNaN(val)) return val;
  if (typeof val === 'string') {
    var parts = val.split('/');
    if (parts.length === 3) {
      var m = parseInt(parts[0], 10);
      var d = parseInt(parts[1], 10);
      var y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      var dt = new Date(y, m - 1, d);
      if (!isNaN(dt)) return dt;
    }
  }
  return null;
}

// Returns the "Reminders" tab, creating and seeding it on first use so J.C.
// only has to fill in email addresses.
function getRemindersSheet_(ss) {
  var sheet = ss.getSheetByName(REMINDERS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(REMINDERS_SHEET_NAME);
  sheet.appendRow(['Name', 'Email', 'Send What ("All lessons" or "Only their own")']);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  sheet.appendRow(['Shop Manager', '', 'All lessons']);
  for (var i = 0; i < DEFAULT_PROS.length; i++) {
    var name = DEFAULT_PROS[i];
    var isJC = name === 'J.C. Freeman';
    sheet.appendRow([name, isJC ? 'invernessjuniortennisacademy@gmail.com' : '', isJC ? 'All lessons' : 'Only their own']);
  }

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 320);
  return sheet;
}

// Reads the Reminders tab into [{name, email, ownOnly}]. Rows without a valid
// email are skipped. "Send What" containing "own" => that person's lessons only.
function getReminderRecipients_(ss) {
  var sheet = getRemindersSheet_(ss);
  var recipients = [];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < values.length; i++) {
      var name = (values[i][0] || '').toString().trim();
      var email = (values[i][1] || '').toString().trim();
      var scope = (values[i][2] || '').toString().trim().toLowerCase();
      if (email.indexOf('@') === -1) continue;
      recipients.push({ name: name, email: email, ownOnly: scope.indexOf('own') !== -1 });
    }
  }
  return recipients;
}

// Builds one recipient's email from their relevant uncharged lessons.
function buildDigestEmail_(list, opts) {
  var agingCount = 0;
  var rowsHtml = '';
  for (var j = 0; j < list.length; j++) {
    var o = list[j];
    var aging = o.ageDays >= AGING_DAYS;
    if (aging) agingCount++;
    var flag = aging ? '⚠️ ' : '';
    var rowBg = aging ? '#fdecea' : (j % 2 === 0 ? '#ffffff' : '#f5f5f5');
    var gmColor = (o.guestMember === 'GUEST') ? '#e74c3c' : '#2ecc71';
    rowsHtml +=
      '<tr style="background:' + rowBg + '">' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + flag + o.date + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + o.ageDays + 'd</td>' +
      (opts.ownOnly ? '' : '<td style="padding:6px 10px;border:1px solid #ddd">' + o.pro + '</td>') +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + o.client + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd;color:' + gmColor + ';font-weight:bold">' + o.guestMember + '</td>' +
      '<td style="padding:6px 10px;border:1px solid #ddd">' + o.duration + (o.people > 1 ? ' · ' + o.people + ' people' : '') + '</td>' +
      '</tr>';
  }

  var intro = opts.ownOnly
    ? 'Hi ' + (opts.name || 'there') + ' — a heads-up that these lessons you logged haven\'t been marked <b>Charged</b> yet. ' +
      'If any look overdue or wrong, give the pro shop a nudge.'
    : 'These lessons are logged but the <b>Charged</b> box is still unchecked. ' +
      'Charge each in Jonas, then tick its box in the sheet and it drops off this list.';

  var subject = '🎾 ' + list.length + (opts.ownOnly ? ' of your lessons' : ' lesson' + (list.length === 1 ? '' : 's')) +
    ' still to charge' + (agingCount > 0 ? ' (' + agingCount + ' aging)' : '');

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' +
    '<p>' + intro + '</p>' +
    (agingCount > 0
      ? '<p style="color:#c0392b"><b>⚠️ ' + agingCount + ' ' + (agingCount === 1 ? 'lesson has' : 'lessons have') +
        ' been waiting ' + AGING_DAYS + '+ days.</b></p>'
      : '') +
    '<table style="border-collapse:collapse;font-size:13px">' +
    '<tr style="background:#1a1a2e;color:#fff">' +
    '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Date</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Age</th>' +
    (opts.ownOnly ? '' : '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Pro</th>') +
    '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Client</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">M/G</th>' +
    '<th style="padding:6px 10px;border:1px solid #ddd;text-align:left">Lesson</th>' +
    '</tr>' + rowsHtml + '</table>' +
    '<p style="margin-top:14px"><a href="' + opts.sheetUrl + '">Open the Lesson Log</a></p>' +
    '</div>';

  return { subject: subject, htmlBody: htmlBody };
}

// Scans every monthly lesson tab for unchecked "Charged" rows and emails each
// person on the Reminders tab. Called by the weekly trigger; also runnable by
// hand to test.
function sendUnchargedDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var outstanding = [];
  var now = new Date();

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!isLessonSheet_(sheet)) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var charged = row[CHARGED_COL - 1] === true;
      var client = (row[2] || '').toString().trim();
      if (charged || !client) continue;

      var dt = coerceDate_(row[0]);
      var ageDays = dt ? Math.floor((now - dt) / 86400000) : 0;
      outstanding.push({
        date: dt ? Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), 'M/d/yy') : (row[0] || ''),
        sortKey: dt ? dt.getTime() : Number.MAX_SAFE_INTEGER,
        ageDays: ageDays,
        pro: (row[1] || '').toString(),
        client: client,
        guestMember: row[3] || '',
        duration: row[4] || '',
        people: row[5] || 1
      });
    }
  }

  outstanding.sort(function (a, b) { return a.sortKey - b.sortKey; });

  var sheetUrl = ss.getUrl();
  var recipients = getReminderRecipients_(ss);

  for (var r = 0; r < recipients.length; r++) {
    var person = recipients[r];
    var list = outstanding;
    if (person.ownOnly) {
      var target = person.name.toLowerCase();
      list = outstanding.filter(function (o) { return o.pro.trim().toLowerCase() === target; });
      // Don't nag a pro who's all caught up.
      if (list.length === 0) continue;
    }

    var email;
    if (list.length === 0) {
      // "All lessons" recipient with nothing outstanding — send the all-clear.
      email = {
        subject: '✅ Lesson charging: all caught up',
        htmlBody:
          '<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">' +
          '<p>Nice work — every logged lesson is marked <b>Charged</b>. Nothing outstanding this week.</p>' +
          '<p><a href="' + sheetUrl + '">Open the Lesson Log</a></p>' +
          '</div>'
      };
    } else {
      email = buildDigestEmail_(list, { ownOnly: person.ownOnly, name: person.name, sheetUrl: sheetUrl });
    }

    MailApp.sendEmail({ to: person.email, subject: email.subject, htmlBody: email.htmlBody });
  }
}
