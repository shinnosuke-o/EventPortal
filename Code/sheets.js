/***********************
 * HEADERS / LIST (used by publicSheet etc)
 ***********************/
const SHEET_HEADER_ROW = 1;
const SHEET_HEADER_COL = 1;

function listEventSheets_() {
  const ss = openSS_(CONFIG.MASTER_SS_ID);
  return ss.getSheets()
    .map(s => s.getName())
    .filter(name => !String(name).includes(CONFIG.EXCLUDE_MARK));
}

function api_listEventSheets() {
  return safeApi_(() => ({ ok: true, data: listEventSheets_() }));
}

function api_getSheetHeaders(payload) {
  return safeApi_(() => {
    const sheetName = payload?.sheetName;
    if (!sheetName) return { ok: false, message: 'sheetName is required' };

    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sheet = mustSheet_(ss, sheetName);

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return { ok: true, data: [] };

    const header = sheet.getRange(SHEET_HEADER_ROW, SHEET_HEADER_COL, 1, lastCol).getValues()[0];
    const data = header
      .map((name, i) => ({ index: i + 1, name: String(name || '').trim() }))
      .filter(x => x.name !== '');
    return { ok: true, data };
  });
}




