// ============================================================
// John Trauth — Subscriber Backend (Google Apps Script)
// Deploy as a Web App: Execute as Me, Anyone can access
//
// SETUP: Run the setup() function once from the editor before
// deploying. This authorizes all required permissions and
// creates the spreadsheet.
// ============================================================

var NOTIFY_EMAILS = ['johntrauth@gmail.com', 'ewhiteart@gmail.com'];

// Run this ONCE from the editor (Run → setup) before deploying
function setup() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.create('John Trauth Subscribers');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  var sheet = ss.insertSheet('Subscribers');
  sheet.appendRow(['Date', 'Email', 'Name', 'Message', 'Source']);
  sheet.getRange('1:1').setFontWeight('bold');
  Logger.log('✅ Setup complete. Spreadsheet: ' + ss.getUrl());
}

function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(ssId);
  return ss.getSheetByName('Subscribers');
}

function doPost(e) {
  try {
    var p = e.parameter;
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

function doGet() {
  return ContentService.createTextOutput('OK');
}
