// MAC Športni Bar — Events Reader
// Deploy as Web App: Execute as Me, Who has access: Anyone
// Re-deploy (new version) after any changes to this script

const EVENTS_SHEET_ID = '1kvSxAfcgqPu-FJYZn4MEp6QeeLNLodcQNFPsI_j46Cg';

function doGet(e) {
  try {
    const sheet  = SpreadsheetApp.openById(EVENTS_SHEET_ID).getActiveSheet();
    const rows   = sheet.getDataRange().getValues();

    const events = rows.slice(1)
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => ({
        name:  (row[0] || '').toString().trim(),
        date:  formatDate(row[1]),
        time:  formatTime(row[2]),
        desc:  (row[3] || '').toString().trim(),
        entry: (row[4] || '').toString().trim()
      }));

    // Sets a global variable + calls ready function — works from file:// with no CORS/redirect issues
    const js = 'window.macEvents=' + JSON.stringify(events) + ';'
             + 'if(typeof window.macEventsReady==="function")window.macEventsReady();';

    return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);

  } catch (err) {
    const js = 'window.macEvents=[];'
             + 'if(typeof window.macEventsReady==="function")window.macEventsReady();';
    return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function formatDate(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return val ? val.toString().trim() : '';
}

function formatTime(val) {
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  return val ? val.toString().trim() : '';
}
