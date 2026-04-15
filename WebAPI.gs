const PICKUP_SS_ID = '1EBlWYTlCCQqlWfyNWXsb2VcNxKH6YCxyLL4OXg4_NAA';

function doGet(e) {
  const type = e.parameter.type;
  let data;

  if (type === 'symbolstats') {
    data = getSheetAsJson_('SymbolStats');
  } else if (type === 'banlist') {
    data = getSheetAsJson_('BanList');
  } else if (type === 'riskstate') {
    data = getSheetAsJson_('RiskState');
  } else {
    data = { error: 'invalid type' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetAsJson_(sheetName) {
  const ss = SpreadsheetApp.openById(PICKUP_SS_ID);
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}