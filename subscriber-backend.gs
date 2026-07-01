// ============================================================
// John Trauth — Subscriber Backend (Google Apps Script)
// Deploy as a Web App: Execute as Me, Anyone can access
// ============================================================

var NOTIFY_EMAILS = ['johntrauth@gmail.com', 'ewhiteart@gmail.com'];

function doPost(e) {
  try {
    var p = e.parameter;
    var email   = (p.email   || '').trim();
    var name    = (p.name    || '').trim();
    var message = (p.message || '').trim();
    var source  = (p.source  || 'subscribe').trim();

    // Write to Google Sheet
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Subscribers');
    if (!sheet) {
      sheet = ss.insertSheet('Subscribers');
      sheet.appendRow(['Date', 'Email', 'Name', 'Message', 'Source']);
      sheet.getRange('1:1').setFontWeight('bold');
    }
    sheet.appendRow([new Date(), email, name, message, source]);

    // Send email for contact-form submissions
    if (source === 'contact' && email) {
      MailApp.sendEmail({
        to: NOTIFY_EMAILS[0],
        cc: NOTIFY_EMAILS.slice(1).join(','),
        replyTo: email,
        subject: 'New message via johntrauth.com — ' + (name || email),
        body: [
          'From:    ' + name,
          'Email:   ' + email,
          '',
          message
        ].join('\n')
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

// Health-check — lets you test the URL in a browser
function doGet() {
  return ContentService.createTextOutput('OK');
}
