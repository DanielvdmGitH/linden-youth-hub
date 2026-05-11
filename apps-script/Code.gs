// The Stand @ Linden - Google Apps Script
// Paste this entire file into your Apps Script editor and deploy as new version

var SHEET_ID = '1N6Q2TCow0xhh0LoSWEs4HQkiqOBaq_6u2aGYL6IDyl8';

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // GET FUNDRAISING STATS
  if (action === 'getStats') {
    var statsSheet = ss.getSheetByName('Stats') || ss.getSheets()[0];
    var raised = statsSheet.getRange('B1').getValue();
    var donors = statsSheet.getRange('B2').getValue();
    return ContentService
      .createTextOutput(JSON.stringify({ raised: raised, donors: donors }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // SHIRT ORDER
  if (action === 'shopOrder') {
    var ordersSheet = ss.getSheetByName('ShopOrders');
    if (!ordersSheet) {
      ordersSheet = ss.insertSheet('ShopOrders');
      ordersSheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Size', 'Qty', 'Total']);
      ordersSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      ordersSheet.setFrozenRows(1);
    }
    ordersSheet.appendRow([
      new Date(),
      e.parameter.name  || '',
      e.parameter.email || '',
      e.parameter.phone || '',
      e.parameter.size  || '',
      e.parameter.qty   || '',
      e.parameter.total || ''
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // STAY UPDATED / CONTACT FORM
  if (action === 'stayUpdated' || action === 'contact') {
    var contactSheet = ss.getSheetByName('Contacts') || ss.getSheets()[0];
    contactSheet.appendRow([
      new Date(),
      e.parameter.firstName || '',
      e.parameter.lastName  || '',
      e.parameter.email     || '',
      action
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // FALLBACK
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}