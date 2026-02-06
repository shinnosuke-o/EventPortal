/***********************
 * GUIDE SHEET
 ***********************/
const GUIDE_START_ROW = 3;
const GUIDE_EVENT_COL = 2;
const GUIDE_DETAIL_COLS = 9;
const GUIDE_PUBLIC_URL_COL = 10;
const GUIDE_APPEND_COLS = 10;

function api_listGuideEvents() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sh = mustSheet_(ss, CONFIG.GUIDE_SHEET_NAME);

    const lastRow = sh.getLastRow();
    if (lastRow < GUIDE_START_ROW) return { ok: true, data: [] };

    const values = sh.getRange(GUIDE_START_ROW, GUIDE_EVENT_COL, lastRow - (GUIDE_START_ROW - 1), 1).getValues(); // B3:B
    const seen = new Set();
    const out = [];

    for (const r of values) {
      const name = norm_(r[0]);
      if (!name) continue;
      if (seen.has(name)) continue;
      seen.add(name);
      out.unshift(name);
    }
    return { ok: true, data: out };
  });
}

function api_getGuideEvent(payload) {
  return safeApi_(() => {
    const eventName = String(payload?.eventName || '').trim();
    if (!eventName) return { ok: false, message: 'eventName is required', data: null };

    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sh = mustSheet_(ss, CONFIG.GUIDE_SHEET_NAME);

    const lastRow = sh.getRange("B:B").getValues().filter(String).length;
    if (lastRow < GUIDE_START_ROW) return { ok: false, message: 'イベント一覧がありません。', data: null };

    // B3:J（9列）
    const values = sh.getRange(GUIDE_START_ROW, GUIDE_EVENT_COL, lastRow - (GUIDE_START_ROW - 1), GUIDE_DETAIL_COLS).getValues();

    let row = null;
    for (const r of values) {
      if (String(r[0] || '').trim() === eventName) { row = r; break; }
    }
    if (!row) return { ok: false, message: 'イベントが見つかりません：' + eventName, data: null };

    return {
      ok: true,
      message: '',
      data: {
        eventName: String(row[0] || '').trim(), // B
        dateFrom: toIso_(row[1]),               // C
        dateTo: toIso_(row[2]),                 // D
        place: String(row[3] || ''),            // E
        deadline: toIso_(row[4]),               // F
        practice: String(row[5] || ''),         // G
        condition: String(row[6] || ''),        // H
        formUrl: String(row[7] || ''),          // I
        publicUrl: String(row[8] || '')         // J
      }
    };
  });
}

function writePublicUrlToGuide_(eventName, publicUrl) {
  const ss = openSS_(CONFIG.MASTER_SS_ID);
  const sheet = mustSheet_(ss, CONFIG.GUIDE_SHEET_NAME);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('案内テンプレシートにデータがありません。');

  const b = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B
  let row = -1;
  for (let i = 0; i < b.length; i++) {
    if (String(b[i][0]).trim() === String(eventName).trim()) {
      row = i + 2;
      break;
    }
  }
  if (row === -1) throw new Error(`案内テンプレにイベント名が見つかりません：${eventName}`);

  sheet.getRange(row, GUIDE_PUBLIC_URL_COL).setValue(publicUrl); // J
}

function appendEventToGuideTemplate_(params) {
  const {
    ssId,
    sheetName,
    eventName,
    dateFrom,
    dateTo,
    place,
    deadline,
    practice,
    condition,
    publishedUrl,
    notify
  } = params;

  const ss = openSS_(ssId);
  const sheet = mustSheet_(ss, sheetName);

  const appendRow = appendRowByColumnB_(sheet, 2);

  const start = dateFrom || null;
  const end = dateTo || dateFrom || null;

  const notifyBool = !!notify;
  const notifiedFlag = !notifyBool; // 要望仕様

  const values = [[
    eventName || '',     // B
    start,               // C
    end,                 // D
    place || '',         // E
    deadline || null,    // F
    practice || '',      // G
    condition || '',     // H
    publishedUrl || '',  // I
    '',                  // J
    notifiedFlag         // K
  ]];

  sheet.getRange(appendRow, GUIDE_EVENT_COL, 1, GUIDE_APPEND_COLS).setValues(values);
}














