// ============================================================
//  LE NEW YORKAIS — FRANCHISE BACKEND (Google Apps Script)
// ------------------------------------------------------------
//  How to deploy:
//  1. Go to https://script.google.com  →  New project
//  2. Paste this entire file into the editor
//  3. Change ADMIN_KEY below to a secret only you know
//  4. Deploy → New deployment → type: Web app
//        - Execute as: Me
//        - Who has access: Anyone
//  5. Copy the /exec URL that Google gives you
//  6. Paste that URL into index.html   → form action=
//     and into dashboard.html          → const API =
// ============================================================

const ADMIN_KEY   = 'newyorkais-2026';                 // ← CHANGE ME
const ADMIN_EMAIL = 'souh445@gmail.com';
const SHEET_NAME  = 'Candidatures';

// ---------- HANDLE FORM SUBMISSIONS ----------
function doPost(e) {
  try {
    const data = e.parameter || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Le New Yorkais — Candidatures');
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // If header row is missing, initialise it from the incoming keys
    if (sheet.getLastRow() === 0) {
      const headers = ['Date', ...Object.keys(data).filter(k => !k.startsWith('_'))];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold').setBackground('#111').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => h === 'Date' ? new Date() : (data[h] || ''));
    sheet.appendRow(row);

    // Notify admin by email
    const html = '<h2 style="font-family:Georgia,serif;color:#e11d24;margin:0 0 12px">' +
                 'Nouvelle candidature — Le New Yorkais</h2>' +
                 '<table style="font-family:Arial;font-size:14px;border-collapse:collapse;width:100%;max-width:640px">' +
                 headers.filter(h => h !== 'Date').map(h =>
                   '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee;font-weight:bold;width:38%;color:#555">' +
                   h + '</td><td style="padding:8px 12px;border-bottom:1px solid #eee">' +
                   (data[h] || '—') + '</td></tr>'
                 ).join('') +
                 '</table>';
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'Nouvelle candidature Franchise — ' + (data['Nom et prénom'] || 'Anonyme'),
      htmlBody: html
    });

    return ContentService.createTextOutput(
      JSON.stringify({ok:true, redirect: (data._next || '')})
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ok:false, error: err.message})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- SERVE DATA TO THE DASHBOARD ----------
function doGet(e) {
  const p = e.parameter || {};
  const callback = p.callback || '';           // JSONP support (bypasses CORS)
  const key = p.key || '';
  let payload;
  if (key !== ADMIN_KEY) {
    payload = {ok:false, error:'unauthorized'};
  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss && ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      payload = {ok:true, rows:[], headers:[]};
    } else {
      const values = sheet.getDataRange().getValues();
      const headers = values.shift();
      const rows = values.map(r => {
        const o = {};
        headers.forEach((h,i) => o[h] = r[i]);
        return o;
      }).reverse();
      payload = {ok:true, rows:rows, headers:headers};
    }
  }
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
