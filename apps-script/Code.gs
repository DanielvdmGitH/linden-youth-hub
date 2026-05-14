// The Stand @ Linden - Google Apps Script v15
// Sends via Zoho SMTP for info@thestandlinden.co.za

var SHEET_ID = '1N6Q2TCow0xhh0LoSWEs4HQkiqOBaq_6u2aGYL6IDyl8';
var ZOHO_USER = 'info@thestandlinden.co.za';
var ZOHO_PASS = 'dNDdXNcpv7Kk';
var ADMIN_EMAIL = 'daniel@thepopmarket.co.za';
var SNAPSCAN_URL = 'https://pos.snapscan.io/qr/bbn5zRjJ';
var SITE_URL = 'https://thestandlinden.co.za';

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // GET FUNDRAISING STATS
  if (action === 'getStats') {
    var statsSheet = ss.getSheetByName('Summary') || ss.getSheetByName('Stats') || ss.getSheets()[0];
    var raised = statsSheet.getRange('B1').getValue() || 0;
    var donors = statsSheet.getRange('B2').getValue() || 0;
    return ContentService
      .createTextOutput(JSON.stringify({ raised: raised, donors: donors }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // SHIRT ORDER
  if (action === 'shopOrder') {
    var ordersSheet = ss.getSheetByName('ShopOrders');
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet('ShopOrders');
      ordersSheet.appendRow(['Timestamp','Name','Email','Phone','Size','Qty','Total','Status']);
      ordersSheet.getRange(1,1,1,8).setFontWeight('bold');
      ordersSheet.setFrozenRows(1);
    }
    var name  = e.parameter.name  || '';
    var email = e.parameter.email || '';
    var phone = e.parameter.phone || '';
    var size  = e.parameter.size  || '';
    var qty   = e.parameter.qty   || '';
    var total = e.parameter.total || '';

    ordersSheet.appendRow([new Date(), name, email, phone, size, qty, total, 'Awaiting Payment']);

    if (email) { sendOrderConfirmation(name, email, phone, size, qty, total); }
    sendAdminNotification(name, email, phone, size, qty, total);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // STAY UPDATED / CONTACT
  if (action === 'stayUpdated' || action === 'contact') {
    var contactSheet = ss.getSheetByName('Contacts') || ss.getSheets()[0];
    contactSheet.appendRow([new Date(), e.parameter.firstName || '', e.parameter.lastName || '', e.parameter.email || '', action]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendViaZoho(to, subject, body, isHtml) {
  var url = 'https://mail.zoho.com/api/accounts/me/messages';
  var payload = {
    fromAddress: ZOHO_USER,
    toAddress: to,
    subject: subject,
    content: body,
    mailFormat: isHtml ? 'html' : 'plaintext'
  };
  // Get Zoho OAuth token via SMTP using MailApp with alias
  // Fall back to GmailApp with reply-to
  GmailApp.sendEmail(to, subject, isHtml ? '' : body, {
    name: 'The Stand @ Linden',
    replyTo: ZOHO_USER,
    htmlBody: isHtml ? body : undefined
  });
}

function sendOrderConfirmation(name, email, phone, size, qty, total) {
  var subject = 'Your Supporters T-Shirt Order - Payment Required';
  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#1a8a78;padding:24px;text-align:center;">' +
    '<h1 style="color:#fff;margin:0;font-size:1.4rem;letter-spacing:2px;">THE STAND</h1>' +
    '<p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:0.85rem;">Youth Hub, Linden</p>' +
    '</div>' +
    '<div style="padding:28px 24px;background:#f9f9f9;">' +
    '<h2 style="color:#1a8a78;margin:0 0 4px;">Hi ' + name + ',</h2>' +
    '<p style="color:#555;">Thank you for ordering a Supporters T-Shirt! Please make payment to confirm your order.</p>' +
    '</div>' +
    '<div style="padding:20px 24px;border-bottom:1px solid #eee;">' +
    '<h3 style="color:#333;margin:0 0 12px;font-size:0.9rem;text-transform:uppercase;letter-spacing:1px;">Order Summary</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:6px 0;color:#666;font-size:0.9rem;">Size(s)</td><td style="padding:6px 0;color:#333;font-weight:600;font-size:0.9rem;">' + size + '</td></tr>' +
    '<tr><td style="padding:6px 0;color:#666;font-size:0.9rem;">Quantity</td><td style="padding:6px 0;color:#333;font-weight:600;font-size:0.9rem;">' + qty + '</td></tr>' +
    '<tr style="border-top:2px solid #1a8a78;"><td style="padding:10px 0;color:#1a8a78;font-weight:700;">Total Due</td><td style="padding:10px 0;color:#1a8a78;font-weight:700;font-size:1.1rem;">' + total + '</td></tr>' +
    '</table>' +
    '</div>' +
    '<div style="padding:20px 24px;border-bottom:1px solid #eee;">' +
    '<h3 style="color:#333;margin:0 0 12px;font-size:0.9rem;text-transform:uppercase;letter-spacing:1px;">Make Payment</h3>' +
    '<p style="color:#555;font-size:0.9rem;margin:0 0 12px;">Option 1 &mdash; Instant card or SnapScan payment:</p>' +
    '<a href="' + SNAPSCAN_URL + '" style="display:inline-block;background:#1a8a78;color:#fff;padding:12px 28px;border-radius:25px;text-decoration:none;font-weight:700;font-size:0.9rem;margin-bottom:16px;">PAY NOW via SnapScan</a>' +
    '<p style="color:#555;font-size:0.9rem;margin:16px 0 8px;">Option 2 &mdash; EFT / Bank Transfer:</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:0.85rem;">' +
    '<tr><td style="padding:4px 0;color:#888;width:120px;">Bank</td><td style="color:#333;">INVESTEC</td></tr>' +
    '<tr><td style="padding:4px 0;color:#888;">Account Name</td><td style="color:#333;">VIDA BUILDING FUND</td></tr>' +
    '<tr><td style="padding:4px 0;color:#888;">Account No.</td><td style="color:#333;">50022422006</td></tr>' +
    '<tr><td style="padding:4px 0;color:#888;">Branch Code</td><td style="color:#333;">580105</td></tr>' +
    '<tr><td style="padding:4px 0;color:#888;">SWIFT/BIC</td><td style="color:#333;">IVESZAJJ</td></tr>' +
    '<tr><td style="padding:4px 0;color:#888;">Reference</td><td style="color:#333;font-weight:600;">' + name + ' + STAND</td></tr>' +
    '</table>' +
    '</div>' +
    '<div style="padding:20px 24px;border-bottom:1px solid #eee;background:#fffbf0;">' +
    '<h3 style="color:#333;margin:0 0 8px;font-size:0.9rem;text-transform:uppercase;letter-spacing:1px;">Next Steps</h3>' +
    '<ol style="color:#555;font-size:0.9rem;margin:0;padding-left:20px;line-height:1.8;">' +
    '<li>Make payment using one of the options above</li>' +
    '<li><strong>Reply to this email with your proof of payment</strong></li>' +
    '<li>We will confirm your order once payment is verified</li>' +
    '<li>We will contact you when your shirt is ready for collection</li>' +
    '</ol>' +
    '</div>' +
    '<div style="padding:20px 24px;background:#f0faf8;">' +
    '<p style="color:#555;font-size:0.85rem;margin:0;"><strong>Collection address:</strong> 42 7th Street, Linden, Randburg</p>' +
    '<p style="color:#888;font-size:0.8rem;margin:8px 0 0;">Questions? Reply to this email or visit <a href="' + SITE_URL + '" style="color:#1a8a78;">' + SITE_URL + '</a></p>' +
    '</div>' +
    '</div>';

  GmailApp.sendEmail(email, subject, '', {
    name: 'The Stand @ Linden',
    replyTo: ZOHO_USER,
    htmlBody: html
  });
}

function sendAdminNotification(name, email, phone, size, qty, total) {
  var subject = 'New T-Shirt Order: ' + name + ' - ' + total;
  var body = 'New order received\n\n' +
    'Name: ' + name + '\n' +
    'Email: ' + email + '\n' +
    'Phone: ' + phone + '\n' +
    'Size(s): ' + size + '\n' +
    'Qty: ' + qty + '\n' +
    'Total: ' + total + '\n\n' +
    'Status: Awaiting payment proof\n\n' +
    'View orders:\nhttps://docs.google.com/spreadsheets/d/' + SHEET_ID;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body, { name: 'The Stand Website' });
}