// Google Apps Script — deploy as Web App
// 1. Create a new Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it into the app's Settings

const HEADERS = ['Date', 'Client Name', 'Guest/Member', 'Duration', 'People', 'Rate', 'Notes'];

const MONTH_COLORS = {
  even: '#ffffff',
  odd: '#e8f0fe'
};

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const proName = data.pro || 'Unknown';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(proName);

    if (!sheet) {
      sheet = ss.insertSheet(proName);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#1a1a2e');
      sheet.getRange(1, 1, 1, HEADERS.length).setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    var dateParts = data.date.split('/');
    var lessonMonth = parseInt(dateParts[0]);
    var lessonYear = '20' + dateParts[2];
    var monthLabel = MONTH_NAMES[lessonMonth - 1] + ' ' + lessonYear;

    var needsDivider = true;
    var lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      for (var r = lastRow; r >= 2; r--) {
        var cellVal = sheet.getRange(r, 1).getValue();
        if (typeof cellVal === 'string' && cellVal.indexOf('---') === 0) {
          if (cellVal.indexOf(monthLabel) !== -1) {
            needsDivider = false;
          }
          break;
        }
      }
    }

    if (needsDivider) {
      sheet.appendRow(['--- ' + monthLabel + ' ---', '', '', '', '', '', '']);
      var dividerRow = sheet.getLastRow();
      sheet.getRange(dividerRow, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.getRange(dividerRow, 1, 1, HEADERS.length).setFontSize(11);
      sheet.getRange(dividerRow, 1, 1, HEADERS.length).setBackground('#d9d9d9');
      sheet.getRange(dividerRow, 1, 1, HEADERS.length).merge();
    }

    var isOddMonth = lessonMonth % 2 === 1;
    var rowColor = isOddMonth ? MONTH_COLORS.odd : MONTH_COLORS.even;

    sheet.appendRow([
      data.date,
      data.client,
      data.guestMember,
      data.duration,
      data.people || 1,
      data.rate || 'FULL',
      data.notes || ''
    ]);

    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 1, 1, HEADERS.length).setBackground(rowColor);

    var gmCell = sheet.getRange(newRow, 3);
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

function doGet() {
  return ContentService
    .createTextOutput('ICC Lesson Log API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
