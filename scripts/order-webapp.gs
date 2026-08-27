// Paste this into the Apps Script editor bound to the order log spreadsheet
// (https://docs.google.com/spreadsheets/d/14IpxQ_Idp1K4BJPZWAsDHxCHZu2BzHi490JiDbP3d90/edit)
// — Extensions > Apps Script — then deploy it as a Web App (Deploy > New
// deployment > type "Web app", execute as "Me", access "Anyone"). Paste the
// resulting /exec URL into ORDER_ENDPOINT near the top of both order.html's
// and order-status.html's <script> blocks.
//
// Sheet columns (header row 2):
// A 날짜&시각 | B 주문상품 | C 개수 | D 이름 | E 전화번호 | F 주소 | G 주문가격 | H 주문ID | I 상태
// Column H doubles as the lookup key AND the cell that carries edit-request
// notes: notes are attached as a Sheets cell Note (메모), never written into
// a normal cell value, so they never show up as visible sheet data.
// Column I (상태) is a dropdown — 입금대기/입금확인/취소 — editable directly
// in the sheet, or from admin.html (doGet ?adminKey=... lists every order;
// doPost {action:'updateStatus'} changes one). Either way the value is
// served back to order-status.html so customers see it update live.
//
// Before admin.html will work you must set a Script Property named
// ADMIN_KEY (Project Settings > Script properties in the Apps Script
// editor) to whatever password you want to gate the admin page with.

var COL = { TIMESTAMP: 1, PRODUCT: 2, QTY: 3, NAME: 4, PHONE: 5, ADDRESS: 6, PRICE: 7, ORDER_ID: 8, STATUS: 9 };
var STATUS_OPTIONS = ['입금대기', '입금확인', '취소'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (data.action === 'note') {
      return writeNote_(sheet, data);
    }

    if (data.action === 'updateStatus') {
      return updateStatus_(sheet, data);
    }

    return createOrder_(sheet, data);
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var params = (e && e.parameter) || {};

    if (params.adminKey) {
      if (!isValidAdminKey_(params.adminKey)) return jsonOutput_({ ok: false, error: '인증에 실패했습니다.' });
      return jsonOutput_({ ok: true, orders: listAllOrders_(sheet) });
    }

    if (params.orderId) {
      return jsonOutput_({ ok: true, order: findOrderById_(sheet, params.orderId) });
    }

    if (params.phone) {
      return jsonOutput_({ ok: true, orders: findOrdersByPhone_(sheet, params.phone) });
    }

    return jsonOutput_({ ok: false, error: 'orderId 또는 phone 파라미터가 필요합니다.' });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function createOrder_(sheet, data) {
  ensureHeader_(sheet);

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var orderId = Utilities.getUuid();

  sheet.appendRow([
    timestamp,
    String(data.product || ''),
    Number(data.quantity) || 0,
    String(data.name || ''),
    String(data.phone || ''),
    String(data.address || ''),
    Number(data.totalPrice) || 0,
    orderId,
    STATUS_OPTIONS[0]
  ]);

  var statusCell = sheet.getRange(sheet.getLastRow(), COL.STATUS);
  statusCell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build()
  );

  return jsonOutput_({ ok: true, orderId: orderId });
}

function writeNote_(sheet, data) {
  var row = findRowByOrderId_(sheet, data.orderId);
  if (!row) return jsonOutput_({ ok: false, error: '주문을 찾을 수 없습니다.' });

  var note = String(data.note || '').trim();
  if (!note) return jsonOutput_({ ok: false, error: '내용을 입력해주세요.' });

  var cell = sheet.getRange(row, COL.ORDER_ID);
  var existing = cell.getNote();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var entry = '[' + stamp + ' 수정요청] ' + note;
  cell.setNote(existing ? existing + '\n' + entry : entry);

  return jsonOutput_({ ok: true });
}

function updateStatus_(sheet, data) {
  if (!isValidAdminKey_(data.adminKey)) return jsonOutput_({ ok: false, error: '인증에 실패했습니다.' });

  var row = findRowByOrderId_(sheet, data.orderId);
  if (!row) return jsonOutput_({ ok: false, error: '주문을 찾을 수 없습니다.' });

  var status = String(data.status || '');
  if (STATUS_OPTIONS.indexOf(status) === -1) return jsonOutput_({ ok: false, error: '올바르지 않은 상태값입니다.' });

  sheet.getRange(row, COL.STATUS).setValue(status);
  return jsonOutput_({ ok: true });
}

function listAllOrders_(sheet) {
  var values = sheet.getDataRange().getValues();
  var orders = [];
  for (var i = values.length - 1; i >= 0; i--) {
    if (!values[i][COL.ORDER_ID - 1]) continue;
    var order = rowToOrder_(sheet, i + 1, values[i]);
    order.note = sheet.getRange(i + 1, COL.ORDER_ID).getNote() || '';
    orders.push(order);
  }
  return orders;
}

function isValidAdminKey_(key) {
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!expected && String(key || '') === expected;
}

function ensureHeader_(sheet) {
  var orderIdHeader = sheet.getRange(2, COL.ORDER_ID);
  if (!orderIdHeader.getValue()) orderIdHeader.setValue('주문ID');

  var statusHeader = sheet.getRange(2, COL.STATUS);
  if (!statusHeader.getValue()) statusHeader.setValue('상태');
}

function findRowByOrderId_(sheet, orderId) {
  if (!orderId) return null;
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][COL.ORDER_ID - 1]) === String(orderId)) return i + 1;
  }
  return null;
}

function findOrderById_(sheet, orderId) {
  var row = findRowByOrderId_(sheet, orderId);
  return row ? rowToOrder_(sheet, row) : null;
}

function findOrdersByPhone_(sheet, phone) {
  var normalized = normalizePhone_(phone);
  if (!normalized) return [];

  var values = sheet.getDataRange().getValues();
  var orders = [];
  for (var i = values.length - 1; i >= 0; i--) {
    if (normalizePhone_(values[i][COL.PHONE - 1]) === normalized) {
      orders.push(rowToOrder_(sheet, i + 1, values[i]));
    }
  }
  return orders;
}

function rowToOrder_(sheet, row, rowValues) {
  var values = rowValues || sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
  return {
    timestamp: formatTimestamp_(values[COL.TIMESTAMP - 1]),
    product: String(values[COL.PRODUCT - 1] || ''),
    quantity: Number(values[COL.QTY - 1]) || 0,
    name: String(values[COL.NAME - 1] || ''),
    phone: String(values[COL.PHONE - 1] || ''),
    address: String(values[COL.ADDRESS - 1] || ''),
    totalPrice: Number(values[COL.PRICE - 1]) || 0,
    orderId: String(values[COL.ORDER_ID - 1] || ''),
    status: String(values[COL.STATUS - 1] || STATUS_OPTIONS[0])
  };
}

function formatTimestamp_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '');
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/[^0-9]/g, '');
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
