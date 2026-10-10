/**
 * Dollar Store Oman – save catalogue orders to this Google Sheet + email alert.
 * Setup (5 min):
 * 1. Open the Google Sheet "DS Orders – Dollar Store Oman". Menu: Extensions > Apps Script.
 * 2. Delete the sample code, paste this whole file, Save.
 * 3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy, allow access.
 * 4. Copy the Web app URL and paste it into index.html:  const SHEET_URL = "https://script.google.com/macros/s/.../exec";
 */
const ALERT_EMAIL = "m.tayyab.r@outlook.com";   // where new-order alerts go

function doPost(e) {
  const o = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Orders") || ss.getSheets()[0].setName("Orders");
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Date","Order #","WhatsApp","Name","Area / notes","Items","Qty","Items OMR","Delivery OMR","Total OMR","Language","Status"]);
    sh.setFrozenRows(1);
  }
  const items = o.items.map(i => `${i.qty} × ${i.name}`).join("\n");
  const qty = o.items.reduce((a, i) => a + i.qty, 0);
  sh.appendRow([new Date(o.date), o.id, "+" + o.phone, o.name, o.area, items, qty,
                Number(o.subtotal), Number(o.delivery), Number(o.total), o.lang === "ar" ? "Arabic" : "English", "New"]);
  MailApp.sendEmail(ALERT_EMAIL, `New order ${o.id} – ${o.total} OMR`,
    o.text + `\n\nReply on WhatsApp: https://api.whatsapp.com/send?phone=${o.phone}`);
  return ContentService.createTextOutput("ok");
}

// Open the Web app URL in a browser: it should say "DS Orders is connected".
function doGet() {
  return ContentService.createTextOutput("DS Orders is connected");
}
