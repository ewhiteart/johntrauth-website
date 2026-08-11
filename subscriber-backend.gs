// ============================================================
// John Trauth — Subscriber Backend (Google Apps Script)
// Deploy as a Web App: Execute as Me, Anyone can access
//
// No setup() needed — spreadsheet ID is hardcoded below.
// ============================================================

var SPREADSHEET_ID  = '1szKoMHrBrEwgCmppsp4WZGOziKMd-eya7ni2uhjpFzc';
var SHEET_NAME      = 'Subscribers';
var NOTIFY_EMAILS   = ['johntrauth@gmail.com', 'ewhiteart@gmail.com'];

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

// All form submissions use GET (POST is blocked by Google for external callers)
function doGet(e) {
  var p = e.parameter;

  // Health check — no email param
  if (!p || !p.email) {
    return ContentService.createTextOutput('OK');
  }

  try {
    var email   = (p.email   || '').trim();
    var name    = (p.name    || '').trim();
    var message = (p.message || '').trim();
    var source  = (p.source  || 'subscribe').trim();

    getSheet().appendRow([new Date(), email, name, message, source]);

    if (source === 'contact' && email) {
      MailApp.sendEmail({
        to: NOTIFY_EMAILS[0],
        cc: NOTIFY_EMAILS.slice(1).join(','),
        replyTo: email,
        subject: 'New message via johntrauth.com — ' + (name || email),
        body: ['From:    ' + name, 'Email:   ' + email, '', message].join('\n')
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
