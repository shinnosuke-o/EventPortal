/***********************
 * TASKS
 ***********************/
const TASK_DATA_START_ROW = 2;
const TASK_DATA_START_COL = 2;
const TASK_DATA_COLS = 8;

function api_getTasks() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = mustSheet_(ss, CONFIG.TASK_SHEET_NAME);

    const lastRow = sh.getLastRow();
    if (lastRow < TASK_DATA_START_ROW) return { ok: true, message: '', tasks: [] };

    const values = sh.getRange(TASK_DATA_START_ROW, TASK_DATA_START_COL, lastRow - (TASK_DATA_START_ROW - 1), TASK_DATA_COLS).getValues(); // B:I
    const tasks = [];

    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      const title = String(r[0] || '').trim();
      if (!title) continue;

      tasks.push({
        rowNumber: i + TASK_DATA_START_ROW,
        task: title,
        category: String(r[1] || ''),
        genre: String(r[2] || ''),
        assignee: String(r[3] || ''),
        due: toIso_(r[4]),
        priority: String(r[5] || ''),
        status: String(r[6] || ''),
        note: String(r[7] || '')
      });
    }

    return { ok: true, message: '', tasks };
  });
}

function api_upsertTask(payload) {
  return writeApi_('api_upsertTask', payload || {}, () => {
    const mode = String(payload?.mode || '').trim(); // create/edit
    const data = payload?.data || {};

    const task = String(data.task || '').trim();
    if (!task) return { ok: false, message: 'タスク内容は必須です。', rowNumber: null };

    const rowValues = [[
      task,
      String(data.category || ''),
      String(data.genre || ''),
      (Array.isArray(data.assignee) ? data.assignee : [])
        .map(x => String(x || '').trim()).filter(Boolean).join(','),
      parseDateYmd_(data.due),
      String(data.priority || ''),
      String(data.status || ''),
      String(data.note || '')
    ]]; // B:I

    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = mustSheet_(ss, CONFIG.TASK_SHEET_NAME);

    if (mode === 'edit') {
      const rowNumber = Number(payload?.rowNumber);
      if (!Number.isFinite(rowNumber) || rowNumber < TASK_DATA_START_ROW) {
        return { ok: false, message: '修正対象行が不正です。', rowNumber: null };
      }
      sh.getRange(rowNumber, TASK_DATA_START_COL, 1, TASK_DATA_COLS).setValues(rowValues);
      return { ok: true, message: '更新しました。', rowNumber };
    }

    const appendRow = appendRowByColumnB_(sh, TASK_DATA_START_ROW);
    sh.getRange(appendRow, TASK_DATA_START_COL, 1, TASK_DATA_COLS).setValues(rowValues);
    return { ok: true, message: '追加しました。', rowNumber: appendRow };
  });
}

/***********************
 * HOLIDAYS (JP public holidays)
 * - Cabinet Office CSV (no Calendar scopes required)
 ***********************/
function api_getHolidays(payload) {
  return safeApi_(() => {
    const year = Number(payload?.year);
    const month = Number(payload?.month); // 1..12
    if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) {
      return { ok: false, message: 'year/month が不正です。', days: [] };
    }

    const days = getJpPublicHolidaysByMonth_(year, month);
    return { ok: true, message: '', days };
  });
}

/**
 * Get Japanese public holidays ("国民の祝日") for a given month.
 * Source: Cabinet Office CSV. (syukujitsu.csv)
 */
function getJpPublicHolidaysByMonth_(year, month) {
  const key = `JP_HOLIDAYS_${year}`; // cache per year
  const prop = PropertiesService.getScriptProperties();
  const cached = prop.getProperty(key);

  let list = null; // [{ymd, name}]
  if (cached) {
    try { list = JSON.parse(cached); } catch(e) { list = null; }
  }

  if (!Array.isArray(list)) {
    // Cabinet Office CSV (filename historically changed; currently syukujitsu.csv is used)
    const url = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code !== 200) throw new Error(`祝日CSVの取得に失敗しました（HTTP ${code}）`);

    // CSV: "国民の祝日・休日月日","国民の祝日・休日名称"
    // date format: YYYY/M/D
    const csv = res.getContentText('MS932'); // 内閣府はShift_JIS系が多い
    const rows = Utilities.parseCsv(csv);

    list = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const dateStr = String(r[0] || '').replace(/"/g, '').trim();
      const name = String(r[1] || '').replace(/"/g, '').trim();
      if (!dateStr || !name || dateStr === '国民の祝日・休日月日') continue;

      // YYYY/M/D -> YYYY-MM-DD
      const m = dateStr.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      if (!m) continue;
      const y = Number(m[1]);
      const mo = Number(m[2]);
      const d = Number(m[3]);
      if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(d)) continue;

      const ymd = `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      list.push({ ymd, name });
    }

    // cache (keep 7 days)
    prop.setProperty(key, JSON.stringify(list));
    prop.setProperty(`${key}_TS`, String(Date.now()));
  }

  const prefix = `${year}-${String(month).padStart(2, '0')}-`;
  return list
    .filter(x => x && typeof x.ymd === 'string' && x.ymd.startsWith(prefix))
    .map(x => x.ymd);
}