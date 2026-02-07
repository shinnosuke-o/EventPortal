/***********************
 * TASKS
 ***********************/
const TASK_DATA_START_ROW = 2;
const TASK_DATA_START_COL = 2;
const TASK_DATA_COLS = 8;

function api_getTasks(payload) {
  return safeApi_(() => {
    guard_(payload || {});
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
  return safeApi_(() => {
    guard_(payload || {});
    const mode = String(payload?.mode || '').trim();
    const data = payload?.data || {};

    const task = String(data.task || '').trim();
    if (!task) return { ok: false, message: 'タスク内容は必須です。', rowNumber: null };

    const category = String(data.category || '');
    const genre = String(data.genre || '');
    const assigneeArr = Array.isArray(data.assignee) ? data.assignee : [];
    const assignee = assigneeArr.map(x => String(x || '').trim()).filter(Boolean).join(',');
    const due = data.due ? new Date(String(data.due) + 'T00:00:00') : null;
    const priority = String(data.priority || '');
    const status = String(data.status || '');
    const note = String(data.note || '');

    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = mustSheet_(ss, CONFIG.TASK_SHEET_NAME);

    const rowValues = [[ task, category, genre, assignee, due, priority, status, note ]];

    if (mode === 'edit') {
      const rowNumber = Number(payload?.rowNumber);
      if (!Number.isFinite(rowNumber) || rowNumber < TASK_DATA_START_ROW) {
        return { ok: false, message: '修正対象行が不正です。', rowNumber: null };
      }
      sh.getRange(rowNumber, TASK_DATA_START_COL, 1, TASK_DATA_COLS).setValues(rowValues);
      return { ok: true, message: '更新しました。', rowNumber };
    }

    // create（B列ベースで末尾探索）
    const appendRow = appendRowByColumnB_(sh, TASK_DATA_START_ROW);
    sh.getRange(appendRow, TASK_DATA_START_COL, 1, TASK_DATA_COLS).setValues(rowValues);

    return { ok: true, message: '追加しました。', rowNumber: appendRow };
  });
}



