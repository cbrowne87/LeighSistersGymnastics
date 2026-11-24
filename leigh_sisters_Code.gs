const NOTIFICATION_EMAIL = 'cbrowne87@gmail.com';
const SHEET_NAME = 'Pledges';

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'Full Name',
      'Email',
      'Phone / WhatsApp',
      'Pledge Amount',
      'Message'
    ]);
  }

  // Authorize the services used by doPost.
  MailApp.getRemainingDailyQuota();
}

function doGet() {
  return HtmlService
    .createHtmlOutput('Leigh Sisters pledge form is connected.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = e && e.parameter ? e.parameter : {};

    const name = String(data.name || '').trim();
    const email = String(data.email || '').trim();
    const phone = String(data.phone || '').trim();
    const amount = String(data.amount || '').trim();
    const message = String(data.message || '').trim();

    if (!name || !email || !amount) {
      return response_('error');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp',
        'Full Name',
        'Email',
        'Phone / WhatsApp',
        'Pledge Amount',
        'Message'
      ]);
    }

    sheet.appendRow([
      new Date(),
      name,
      email,
      phone,
      amount,
      message
    ]);

    const subject = 'New Leigh Sisters Pledge';
    const body =
      'A new pledge has been submitted.\n\n' +
      'Full Name: ' + name + '\n' +
      'Email: ' + email + '\n' +
      'Phone / WhatsApp: ' + phone + '\n' +
      'Pledge Amount: ' + amount + '\n' +
      'Message: ' + message + '\n\n' +
      'The pledge has also been recorded in the Google Sheet.';

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      body: body,
      replyTo: email,
      name: 'Leigh Sisters Pledge Form'
    });

    return response_('success');

  } catch (error) {
    console.error(error);
    return response_('error');
  }
}

function response_(status) {
  const safeStatus = status === 'success' ? 'success' : 'error';

  const html =
    '<!doctype html><html><body>' +
    '<script>' +
    'window.parent.postMessage(' +
    JSON.stringify({
      type: 'pledgeSubmission',
      status: safeStatus
    }) +
    ', "*");' +
    '</script>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
