// Google Apps Script — deploy as Web App
// 1. Open your existing "Member/Guest Pros Charging Sheet" Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code (replace anything already there)
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL — that's all you need.
//    The app already has the default URL hardcoded, but if you redeploy
//    you can update it in Settings.
//
// This script writes to your EXISTING per-pro monthly tabs
// (e.g., "J.C. September 2026") using your existing column format:
//   Date | Name | Guest or Member | Length of Lesson/Amount | Notes | Charged
//
// It adds a "Charged" checkbox column (F) for billing tracking.
// If a pro's tab for the current month doesn't exist yet, it creates one.
//
// The "Pros" tab lists the pro roster — edit column A to add/remove pros.
// The app's dropdown updates on next load.
//
// WEEKLY UNCHARGED REMINDER (optional — one-time setup):
//   After pasting this code, run "setupWeeklyReminder" once from the
//   toolbar dropdown (click Run, then authorize). Every Monday at 7am
//   it emails people listed on the "Reminders" tab about uncharged lessons.

const HEADERS = ['Date:', 'Name:', 'Guest or Member:', 'Length of Lesson/Amount:', 'Notes:', 'Charged'];
const PROS_SHEET_NAME = 'Pros';
const REMINDERS_SHEET_NAME = 'Reminders';
const CHARGED_COL = 6;
const GUEST_MEMBER_COL = 3;

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

// Builds a tab name like "J.C. September 2026" from the pro name and lesson date.
// Existing tabs use the pro's first name only ("Will September 2026"), so we
// look for that first, then the full name, and create with the short form.
function proMonthTabName_(ss, proName, dateStr) {
  var suffix;
  var parts = String(dateStr).split('/');
  var m = parts.length === 3 ? parseInt(parts[0], 10) : 0;
  if (m >= 1 && m <= 12) {
    var y = parts[2].length === 2 ? '20' + parts[2] : parts[2];
    suffix = ' ' + MONTH_NAMES[m - 1] + ' ' + y;
  } else {
    var now = new Date();
    suffix = ' ' + MONTH_NAMES[now.getMonth()] + ' ' + now.getFullYear();
  }
  var shortName = String(proName).trim().split(/\s+/)[0];
  if (ss.getSheetByName(shortName + suffix)) return shortName + suffix;
  if (ss.getSheetByName(proName + suffix)) return proName + suffix;
  return shortName + suffix;
}

// Extracts the pro name from a tab name like "J.C. September 2026" → "J.C."
function proNameFromTab_(tabName) {
  for (var i = 0; i < MONTH_NAMES.length; i++) {
    var idx = tabName.indexOf(' ' + MONTH_NAMES[i] + ' ');
    if (idx !== -1) return tabName.substring(0, idx);
  }
  return null;
}

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

// True if a tab looks like a per-pro lesson tab (has the right headers or
// matches the existing format with Date in A1).
function isLessonSheet_(sheet) {
  var name = sheet.getName();
  if (name === PROS_SHEET_NAME || name === REMINDERS_SHEET_NAME) return false;
  if (sheet.getLastRow() < 2) return false;

  var firstCell = sheet.getRange(1, 1).getValue().toString().trim();
  if (firstCell === 'Date:' || firstCell === 'Date') return true;

  return false;
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
    const proName = data.pro || 'Unknown';
    const tabName = proMonthTabName_(ss, proName, data.date);
    var sheet = getOrCreateLessonSheet_(ss, tabName);

    // Combine duration + people into one field to match existing format
    var lessonAmount = data.duration || '1 hour';
    if (data.people && data.people > 1) {
      lessonAmount += ' ' + data.people + ' people';
    }

    sheet.appendRow([
      data.date,
      data.client,
      data.guestMember,
      lessonAmount,
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
      '<td style="padding:6px 10px;border:1px solid #ddd">' + o.lesson + '</td>' +
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

// Scans every lesson tab for unchecked "Charged" rows and emails each
// recipient. Extracts the pro name from the tab name.
function sendUnchargedDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var outstanding = [];
  var now = new Date();

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (!isLessonSheet_(sheet)) continue;

    var tabName = sheet.getName();
    var proFromTab = proNameFromTab_(tabName) || tabName;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    var numCols = Math.min(sheet.getLastColumn(), HEADERS.length);
    var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var client = (row[1] || '').toString().trim();
      if (!client) continue;

      // Check "Charged" column if it exists on this tab
      var charged = numCols >= CHARGED_COL ? row[CHARGED_COL - 1] === true : false;
      if (charged) continue;

      var dt = coerceDate_(row[0]);
      var ageDays = dt ? Math.floor((now - dt) / 86400000) : 0;
      outstanding.push({
        date: dt ? Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), 'M/d/yy') : (row[0] || ''),
        sortKey: dt ? dt.getTime() : Number.MAX_SAFE_INTEGER,
        ageDays: ageDays,
        pro: proFromTab,
        client: client,
        guestMember: row[2] || '',
        lesson: row[3] || ''
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
      var target = person.name.trim().split(/\s+/)[0].toLowerCase();
      list = outstanding.filter(function (o) { return o.pro.trim().split(/\s+/)[0].toLowerCase() === target; });
      if (list.length === 0) continue;
    }

    var email;
    if (list.length === 0) {
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
