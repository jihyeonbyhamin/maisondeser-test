// Paste this into the Apps Script editor bound to the order log spreadsheet
// (https://docs.google.com/spreadsheets/d/14IpxQ_Idp1K4BJPZWAsDHxCHZu2BzHi490JiDbP3d90/edit)
// — Extensions > Apps Script — then deploy it as a Web App (Deploy > New
// deployment > type "Web app", execute as "Me", access "Anyone"). Paste the
// resulting /exec URL into ORDER_ENDPOINT near the top of both order.html's
// and order-status.html's <script> blocks.
//
// Sheet columns (header row 2):
// A 날짜&시각 | B 주문상품 | C 개수 | D 이름 | E 전화번호 | F 주소 | G 주문가격 | H 주문ID
// Column H doubles as the lookup key AND the cell that carries edit-request
// notes: notes are attached as a Sheets cell Note (메모), never written into
// a normal cell value, so they never show up as visible sheet data.

var COL = { TIMESTAMP: 1, PRODUCT: 2, QTY: 3, NAME: 4, PHONE: 5, ADDRESS: 6, PRICE: 7, ORDER_ID: 8 };

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (data.action === 'note') {
      return writeNote_(sheet, data);
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
    orderId
  ]);

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

function ensureHeader_(sheet) {
  var headerCell = sheet.getRange(2, COL.ORDER_ID);
  if (!headerCell.getValue()) headerCell.setValue('주문ID');
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
  var values = rowValues || sheet.getRange(row, 1, 1, COL.ORDER_ID).getValues()[0];
  return {
    timestamp: formatTimestamp_(values[COL.TIMESTAMP - 1]),
    product: String(values[COL.PRODUCT - 1] || ''),
    quantity: Number(values[COL.QTY - 1]) || 0,
    name: String(values[COL.NAME - 1] || ''),
    phone: String(values[COL.PHONE - 1] || ''),
    address: String(values[COL.ADDRESS - 1] || ''),
    totalPrice: Number(values[COL.PRICE - 1]) || 0,
    orderId: String(values[COL.ORDER_ID - 1] || '')
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
