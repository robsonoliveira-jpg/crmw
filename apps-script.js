// ============================================================
// COLE ESTE CÓDIGO no Google Apps Script
// script.google.com → Novo projeto → cole aqui → Deploy
// ============================================================

const SHEET_NAME = "Leads";

function doPost(e) {
  const sheet = getSheet();
  const data = JSON.parse(e.postData.contents);
  sheet.appendRow([data.nome, data.email, data.whatsapp, data.data]);
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.action === "get") {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    const leads = rows.slice(1).map(r => ({
      nome: r[0], email: r[1], whatsapp: r[2], data: r[3]
    })).reverse();
    return ContentService
      .createTextOutput(JSON.stringify({ leads }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Nome", "Email", "WhatsApp", "Data"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  return sheet;
}
