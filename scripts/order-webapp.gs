// Paste this into the Apps Script editor bound to the order log spreadsheet
// (https://docs.google.com/spreadsheets/d/14IpxQ_Idp1K4BJPZWAsDHxCHZu2BzHi490JiDbP3d90/edit)
// — Extensions > Apps Script — then deploy it as a Web App (Deploy > New
// deployment > type "Web app", execute as "Me", access "Anyone"). Paste the
// resulting /exec URL into ORDER_ENDPOINT near the top of order.html's
// <script> block.
//
// Appends one row per submission, matching the sheet's existing header:
// 날짜&시각 | 주문상품 | 개수 | 이름 | 전화번호 | 주소 | 주문가격
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      timestamp,
      String(data.product || ''),
      Number(data.quantity) || 0,
      String(data.name || ''),
      String(data.phone || ''),
      String(data.address || ''),
      Number(data.totalPrice) || 0
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
