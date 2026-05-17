// Google Apps Script — deploy as Web App
// 1. Create a Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it into the PWA's SCRIPT_URL variable

const SHEET_NAME = 'Lessons';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Pro Name', 'Client', 'Duration (min)', 'Notes']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date().toLocaleString(),
      data.proName || '',
      data.client,
      data.duration,
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

function doGet() {
  return ContentService
    .createTextOutput('Lesson Input API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
