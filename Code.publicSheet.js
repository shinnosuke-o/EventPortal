/***********************
 * PUBLIC CHECK SHEET
 ***********************/
const PUBLIC_SHEET_SUFFIX = '出欠確認_公開用';
const PUBLIC_QUERY_HEADER_ROWS = 1;
const PUBLIC_IMPORTRANGE_RANGE = 'A:ZZ';
const PUBLIC_COL_PREFIX = 'Col';

function api_createPublicCheckSheet(payload) {
  return safeApi_(() => {
    guard_(payload || {});
    const sheetName = payload?.sheetName;
    const includeCols = payload?.includeCols || [];
    if (!sheetName) throw new Error('シート選択は必須です。');

    const master = openSS_(CONFIG.MASTER_SS_ID);
    const srcSheet = master.getSheetByName(sheetName);
    if (!srcSheet) throw new Error('元シートが見つかりません：' + sheetName);

    const eventName = sheetName;

    const cols = includeCols
      .filter(x => x.include)
      .map(x => x.index)
      .filter(n => Number.isFinite(n) && n >= 1);

    if (cols.length === 0) throw new Error('表示列が0件です。');

    const selectCols = cols.map(n => PUBLIC_COL_PREFIX + n).join(',');
    const query = `select ${selectCols}`;

    const destTitle = `${eventName}${PUBLIC_SHEET_SUFFIX}`;
    const dest = SpreadsheetApp.create(destTitle);
    const destId = dest.getId();
    const destUrl = dest.getUrl();

    const first = dest.getSheets()[0];
    first.setName(eventName);

    const rangeA1 = `${escapeSheetName_(eventName)}!${PUBLIC_IMPORTRANGE_RANGE}`;
    const formula = `=QUERY(IMPORTRANGE("${CONFIG.MASTER_SS_ID}","${rangeA1}"),"${query}",PUBLIC_QUERY_HEADER_ROWS)`;
    first.getRange(1, 1).setFormula(formula);

    moveFileToFolder_(destId, CONFIG.PUBLIC_FOLDER_ID);
    writePublicUrlToGuide_(eventName, destUrl);

    return {
      ok: true,
      url: destUrl,
      message: '作成しました。初回は「アクセスを許可」が必要です。'
    };
  });
}






