/**
 * Bank Pulse — Google Sheet entry endpoint
 *
 * SETUP:
 * 1. Open your Google Sheet (the one with the KPIs/CASA/Segments/etc. tabs).
 * 2. Extensions → Apps Script.
 * 3. Delete any starter code and paste this file in.
 * 4. Click Deploy → New deployment → type: "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone with the link
 * 5. Click Deploy, authorize when prompted, and copy the Web App URL.
 * 6. Paste that URL into the "Entry endpoint" field on the DW Blueprint page.
 *
 * This script only appends rows to whatever sheet tab + columns you tell it
 * to, via a JSON POST body: { "tab": "KPIs", "row": ["Segmental Deposits", 5740259, 1.56] }
 */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var tabName = body.tab;
    var row = body.row;

    if (!tabName || !row || !Array.isArray(row)) {
      return jsonResponse({ ok: false, error: 'Expected { tab, row: [...] } in the request body.' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      return jsonResponse({ ok: false, error: 'No tab named "' + tabName + '" in this spreadsheet.' });
    }

    sheet.appendRow(row);

    return jsonResponse({ ok: true, tab: tabName, rowWritten: row, newRowCount: sheet.getLastRow() });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, message: 'Bank Pulse entry endpoint is live. POST { tab, row } to add data.' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
