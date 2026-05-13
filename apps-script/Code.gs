// The Stand @ Linden - Google Apps Script
// Version 14 - with order confirmation emails

var SHEET_ID = '1N6Q2TCow0xhh0LoSWEs4HQkiqOBaq_6u2aGYL6IDyl8';
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
      ordersSheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Size', 'Qty', 'Total', 'Status']);
      ordersSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      ordersSheet.setFrozenRows(1);
    }
    var name  = e.parameter.name  || '';
    var email = e.parameter.email || '';
    var phone = e.parameter.phone || '';
    var size  = e.parameter.size  || '';
    var qty   = e.parameter.qty   || '';
    var total = e.parameter.total || '';

    ordersSheet.appendRow([new Date(), name, email, phone, size, qty, total, 'Awaiting Payment']);

    // Send confirmation email to customer
    if (email) {
      sendOrderConfirmation(name, email, phone, size, qty, total);
    }

    // Send notification to admin
    sendAdminNotification(name, email, phone, size, qty, total);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // STAY UPDATED / CONTACT FORM
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

function sendOrderConfirmation(name, email, phone, size, qty, total) {
  var subject = 'Your Stand Supporters T-Shirt Order - Payment Required';
  var body = 'Hi ' + name + ',\n\n' +
    'Thank you for your order! Here is a summary:\n\n' +
    '  Size(s): ' + size + '\n' +
    '  Quantity: ' + qty + '\n' +
    '  Total: ' + total + '\n\n' +
    'PAYMENT INSTRUCTIONS\n' +
    '--------------------\n' +
    'Please make payment of ' + total + ' using one of the following methods:\n\n' +
    '1. SnapScan / Card Payment (instant):\n' +
    '   ' + SNAPSCAN_URL + '\n' +
    '   Use your name as reference.\n\n' +
    '2. EFT / Bank Transfer:\n' +
    '   Bank: INVESTEC\n' +
    '   Account Name: VIDA BUILDING FUND\n' +
    '   Account No: 50022422006\n' +
    '   Branch Code: 580105\n' +
    '   SWIFT/BIC: IVESZAJJ\n' +
    '   Reference: ' + name + ' + STAND\n\n' +
    'PROOF OF PAYMENT\n' +
    '----------------\n' +
    'Once payment is made, please reply to this email with your proof of payment.\n' +
    'We will confirm your order and let you know when your shirt is ready for collection.\n\n' +
    'COLLECTION\n' +
    '----------\n' +
    'Shirts are collected from The Stand, 42 7th Street, Linden, Randburg.\n' +
    'We will contact you with collection details once payment is confirmed.\n\n' +
    'Thank you for supporting The Stand Youth Hub, Linden!\n\n' +
    'The Stand Team\n' +
    SITE_URL;

  GmailApp.sendEmail(email, subject, body, {
    name: 'The Stand @ Linden',
    replyTo: ADMIN_EMAIL
  });
}

function sendAdminNotification(name, email, phone, size, qty, total) {
  var subject = 'New T-Shirt Order: ' + name + ' - ' + total;
  var body = 'New order received on thestandlinden.co.za\n\n' +
    'Name: ' + name + '\n' +
    'Email: ' + email + '\n' +
    'Phone: ' + phone + '\n' +
    'Size(s): ' + size + '\n' +
    'Qty: ' + qty + '\n' +
    'Total: ' + total + '\n\n' +
    'Status: Awaiting payment proof\n\n' +
    'View orders: https://docs.google.com/spreadsheets/d/' + SHEET_ID;

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body, {
    name: 'The Stand Website'
  });
}