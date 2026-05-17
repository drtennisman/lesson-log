// Google Apps Script — deploy as Web App
// 1. Create a new Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it into the app's Settings

const SHEET_NAME = 'Lessons';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
      || createSheet();

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.date,
      data.pro,
      data.client,
      data.guestMember,
      data.duration,
      data.people || 1,
      data.rate || 'FULL',
      data.notes || ''
    ]);

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

function createSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_NAME);
  const headers = ['Date', 'Pro', 'Client Name', 'Guest/Member', 'Duration', 'People', 'Rate', 'Notes'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
  return sheet;
}

function doGet() {
  return ContentService
    .createTextOutput('ICC Lesson Log API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
