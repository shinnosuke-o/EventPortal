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



