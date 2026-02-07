/***********************
 * ATTENDANCE SUMMARY
 ***********************/
function api_listAttendanceEventSheets() {
  return safeApi_(() => ({ ok: true, data: listEventSheets_() }));
}

function api_listDancerMasterSheets() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const names = ss.getSheets().map(s => s.getName());
    const data = names.filter(n => String(n).startsWith(CONFIG.DANCER_MASTER_PREFIX));
    return { ok: true, data };
  });
}

function api_getSheetHeadersRow1(payload) {
  return safeApi_(() => {
    const sheetName = String(payload?.sheetName || '').trim();
    if (!sheetName) return { ok: false, message: 'sheetName is required' };

    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: 'シートが見つかりません：' + sheetName };

    const lastCol = sh.getLastColumn();
    if (lastCol < 1) return { ok: true, data: [] };

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = header
      .map((name, i) => ({ index: i + 1, name: String(name || '').trim() }))
      .filter(x => x.name !== '');

    return { ok: true, data };
  });
}

function findHeaderIndex_(headerArr, keywords) {
  const keys = (keywords || []).map(String);
  for (let i = 0; i < headerArr.length; i++) {
    const h = String(headerArr[i] || '');
    if (keys.some(k => h.includes(k))) return i + 1;
  }
  return -1;
}

// 正規化
function normalizeName_(s) {
  let x = String(s || '');
  x = x.replace(/[\s　]+/g, '');

  const map = {
    '髙':'高',
    '﨑':'崎',
    '塚':'塚',
    '神':'神',
  };
  x = x.split('').map(ch => map[ch] || ch).join('');

  if (x.normalize) x = x.normalize('NFKC');
  return x.trim();
}

function normalizeYosana_(s) {
  let x = String(s || '');
  x = x.replace(/[\s　]+/g, '');
  if (x.normalize) x = x.normalize('NFKC');
  x = x.toLowerCase();
  return x.trim();
}

function levenshteinLE1_(a, b) {
  if (a === b) return true;
  const la = a.length, lb = b.length;
  if (Math.abs(la - lb) > 1) return false;

  if (la === lb) {
    let diff = 0;
    for (let i = 0; i < la; i++) {
      if (a[i] !== b[i]) { diff++; if (diff > 1) return false; }
    }
    return true;
  }

  const s1 = la < lb ? a : b;
  const s2 = la < lb ? b : a;
  let i = 0, j = 0, diff = 0;
  while (i < s1.length && j < s2.length) {
    if (s1[i] === s2[j]) { i++; j++; continue; }
    diff++; if (diff > 1) return false;
    j++;
  }
  return true;
}

function api_aggregateAttendance(payload) {
  return safeApi_(() => {
    const eventSheetName  = String(payload?.eventSheetName || '').trim();
    const masterSheetName = String(payload?.masterSheetName || '').trim();
    const cols = payload?.cols || {};

    if (!eventSheetName)  return { ok: false, message: 'eventSheetName is required' };
    if (!masterSheetName) return { ok: false, message: 'masterSheetName is required' };

    const requiredKeys = ['ts','name','yosana','attend'];
    for (const k of requiredKeys) {
      const v = Number(cols[k]);
      if (!Number.isFinite(v) || v < 1) return { ok: false, message: `列指定が不正です：${k}` };
    }

    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const ev = mustSheet_(ss, eventSheetName);
    const ms = mustSheet_(ss, masterSheetName);

    const evLastRow = ev.getLastRow();
    const evLastCol = ev.getLastColumn();
    if (evLastRow < 2) {
      return { ok: true, participants: [], undecided: [], unanswered: [], review: [], summary: { totalMaster: 0, totalAnswered: 0 } };
    }

    const evValues = ev.getRange(2, 1, evLastRow - 1, evLastCol).getValues();

    const msLastRow = ms.getLastRow();
    const msLastCol = ms.getLastColumn();
    if (msLastRow < 2) return { ok: false, message: 'マスタにデータがありません。' };

    const msHeader = ms.getRange(1, 1, 1, msLastCol).getValues()[0].map(x => String(x || '').trim());
    const msNameCol = findHeaderIndex_(msHeader, ['名前','氏名']);
    const msYosaCol = findHeaderIndex_(msHeader, ['よさ名']);
    if (msNameCol < 1) return { ok: false, message: 'マスタの「名前/氏名」列が見つかりません。' };
    if (msYosaCol < 1) return { ok: false, message: 'マスタの「よさ名」列が見つかりません。' };

    const msValues = ms.getRange(2, 1, msLastRow - 1, msLastCol).getValues();

    // 1) 回答：同一人物（正規化名）で最新採用
    const latestByName = new Map();
    for (const r of evValues) {
      const ts = r[cols.ts - 1];
      const name = r[cols.name - 1];
      const yosana = r[cols.yosana - 1];
      const attend = r[cols.attend - 1];

      if (!ts || !name) continue;

      const normName = normalizeName_(String(name));
      if (!normName) continue;

      const cur = latestByName.get(normName);
      const curTs = cur?.ts ? new Date(cur.ts).getTime() : -1;
      const newTs = new Date(ts).getTime();

      if (!cur || (Number.isFinite(newTs) && newTs >= curTs)) {
        latestByName.set(normName, {
          ts,
          name: String(name || '').trim(),
          yosana: String(yosana || '').trim(),
          gender: cols.gender ? String(r[cols.gender - 1] || '').trim() : '',
          attend: String(attend || '').trim(),
          undecidedReason: cols.undecided ? String(r[cols.undecided - 1] || '').trim() : '',
          _normName: normName,
          _normYosana: normalizeYosana_(String(yosana || '')),
        });
      }
    }

    // 2) マスタ正規化
    const masters = msValues.map(row => {
      const name = String(row[msNameCol - 1] || '').trim();
      const yosana = String(row[msYosaCol - 1] || '').trim();
      return {
        name,
        yosana,
        _normName: normalizeName_(name),
        _normYosana: normalizeYosana_(yosana),
      };
    }).filter(x => x._normName || x._normYosana);

    // 3) マスタ→回答突合
    const matchedMaster = new Map(); // masterIdx -> {res, needsReview}
    const usedResponseKey = new Set(); // normName used

    for (let i = 0; i < masters.length; i++) {
      const m = masters[i];
      if (!m._normName) continue;

      const direct = latestByName.get(m._normName);
      if (direct) {
        matchedMaster.set(i, { res: direct, needsReview: false });
        usedResponseKey.add(direct._normName);
        continue;
      }

      let best = null;
      for (const [k, res] of latestByName.entries()) {
        if (usedResponseKey.has(k)) continue;
        if (levenshteinLE1_(m._normName, k)) { best = res; break; }
      }
      if (best) {
        matchedMaster.set(i, { res: best, needsReview: true });
        usedResponseKey.add(best._normName);
      }
    }

    // 4) 未回答候補 → よさ名で要確認一致
    const review = [];
    for (let i = 0; i < masters.length; i++) {
      if (matchedMaster.has(i)) continue;

      const m = masters[i];
      if (!m._normYosana) continue;

      let found = null;
      for (const [, res] of latestByName.entries()) {
        if (!res._normYosana) continue;
        if (levenshteinLE1_(m._normYosana, res._normYosana)) { found = res; break; }
      }
      if (found) {
        matchedMaster.set(i, { res: found, needsReview: true });
        review.push({
          masterName: m.name,
          masterYosana: m.yosana,
          responseName: found.name,
          responseYosana: found.yosana,
          attend: found.attend
        });
      }
    }

    // 5) 出力組み立て（マスタ側）
    const participants = [];
    const undecided = [];
    const unanswered = [];

    for (let i = 0; i < masters.length; i++) {
      const m = masters[i];
      const match = matchedMaster.get(i);

      if (!match) {
        unanswered.push({ name: m.name, yosana: m.yosana });
        continue;
      }

      const r = match.res;
      const a = String(r.attend || '').trim();
      const isDancer  = a.includes('踊り子');
      const isStaff   = a.includes('スタッフ');
      const isPending = a.includes('未定');

      if (isDancer || isStaff) {
        participants.push({
          name: r.name || m.name,
          yosana: r.yosana || m.yosana,
          gender: r.gender || '',
          attendType: isStaff ? 'スタッフ' : '踊り子',
          needsReview: !!match.needsReview
        });
      } else if (isPending) {
        undecided.push({
          name: r.name || m.name,
          yosana: r.yosana || m.yosana,
          reason: r.undecidedReason || '',
          needsReview: !!match.needsReview
        });
      }
    }

    // 6) ★マスタ未登録だが回答に存在：参加/未定ならそのまま表示
    const usedResponseNormNames = new Set();
    for (const [, match] of matchedMaster.entries()) {
      if (match?.res?._normName) usedResponseNormNames.add(match.res._normName);
    }

    for (const [normName, r] of latestByName.entries()) {
      if (usedResponseNormNames.has(normName)) continue;

      const a = String(r.attend || '').trim();
      const isDancer  = a.includes('踊り子');
      const isStaff   = a.includes('スタッフ');
      const isPending = a.includes('未定');

      if (isDancer || isStaff) {
        participants.push({
          name: r.name,
          yosana: r.yosana,
          gender: r.gender || '',
          attendType: isStaff ? 'スタッフ' : '踊り子',
          needsReview: false
        });
      } else if (isPending) {
        undecided.push({
          name: r.name,
          yosana: r.yosana,
          reason: r.undecidedReason || '',
          needsReview: false
        });
      }
    }

    return {
      ok: true,
      signature: 'ATTEND_AGG_v20260203_refactor_01',
      participants,
      undecided,
      unanswered,
      review,
      summary: {
        totalMaster: masters.length,
        totalAnswered: masters.length - unanswered.length
      }
    };
  });
}



/***********************
 * PARTICIPANTS DB (list/create/import)
 ***********************/
const PDB_HEADER_ROW = 1;
const PDB_DATA_START_ROW = 2;
const PDB_MAX_COLS = 7; // A:G
const PDB_PIVOT_START_ROW_INDEX = 1; // I2 (0-based)
const PDB_PIVOT_START_COL_INDEX = 8; // I2 (0-based)
const PDB_PIVOT_ROW_OFFSET = 4; // E
const PDB_PIVOT_COL_OFFSET = 3; // D
const PDB_PIVOT_VAL_OFFSET = 1; // B

function buildParticipantsPivot_(sheetId, endRowIndex) {
  return {
    source: {
      sheetId,
      startRowIndex: 0,
      endRowIndex,
      startColumnIndex: 0,
      endColumnIndex: PDB_MAX_COLS
    },
    rows: [{
      sourceColumnOffset: PDB_PIVOT_ROW_OFFSET,
      sortOrder: 'DESCENDING',
      showTotals: true
    }],
    columns: [{
      sourceColumnOffset: PDB_PIVOT_COL_OFFSET,
      sortOrder: 'DESCENDING',
      showTotals: true
    }],
    values: [{
      summarizeFunction: 'COUNTA',
      sourceColumnOffset: PDB_PIVOT_VAL_OFFSET
    }]
  };
}

function updateParticipantsPivot_(ssId, sheetId, endRowIndex) {
  if (typeof Sheets === 'undefined') return;
  const requests = [{
    updateCells: {
      start: { sheetId, rowIndex: PDB_PIVOT_START_ROW_INDEX, columnIndex: PDB_PIVOT_START_COL_INDEX },
      rows: [{ values: [{ pivotTable: buildParticipantsPivot_(sheetId, endRowIndex) }] }],
      fields: 'pivotTable'
    }
  }];
  Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
}

function api_listParticipantDbSheets() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
    return { ok: true, sheets: ss.getSheets().map(s => s.getName()) };
  });
}

// 互換：古いクライアントが呼ぶ可能性
function api_listParticipantsDbSheets() {
  const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
  return ss.getSheets().map(s => s.getName());
}

function api_createParticipantsDbSheet(payload) {
  return safeApi_(() => {
    guard_(payload || {});
    return withWriteLockAndAudit_('api_createParticipantsDbSheet', payload || {}, () => {
    const sheetName = String(payload?.sheetName || '').trim();
    if (!sheetName) return { ok: false, message: 'sheetName is required' };

    const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
    if (ss.getSheetByName(sheetName)) {
      return { ok: false, message: '同名のシートが既に存在します：' + sheetName };
    }

    const sh = ss.insertSheet(sheetName);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(1);

    const headers = ['SEQ','名前','よさ名','性別','参加区分','登録区分','備考'];
    sh.getRange(PDB_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(PDB_HEADER_ROW);

    // Table作成（Sheets Advanced Service）
    if (typeof Sheets === 'undefined') {
      return { ok: false, message: 'Sheets API（高度なGoogleサービス）が無効です。「サービス」から Sheets API を追加してください。' };
    }

    const sheetId = sh.getSheetId();
    const requests = [{
      addTable: {
        table: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: PDB_MAX_COLS
          },
          name: sheetName
        }
      }
    }];

    Sheets.Spreadsheets.batchUpdate({ requests }, ss.getId());
    updateParticipantsPivot_(ss.getId(), sheetId, 1);

    return { ok: true, message: 'シートを作成しました：' + sheetName };
  });
  });
}

// 互換：別名（古い関数名）
function api_createParticipantDbSheet(payload) {
  return api_createParticipantsDbSheet(payload);
}

function api_importParticipants(payload) {
  return safeApi_(() => {
    guard_(payload || {});
    return withWriteLockAndAudit_('api_importParticipants', payload || {}, () => {
    const destSheetName = String(payload?.destSheetName || '').trim();
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];

    if (!destSheetName) return { ok: false, message: 'destSheetName is required' };
    if (rows.length === 0) return { ok: false, message: 'rows is empty' };

    const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
    const sh = ss.getSheetByName(destSheetName);
    if (!sh) return { ok: false, message: 'インポート先シートが見つかりません：' + destSheetName };

    const lastDataRow = findLastDataRowByCols_(sh, PDB_DATA_START_ROW, PDB_MAX_COLS); // A:G を見て最終行判定
    const startRow = Math.max(PDB_DATA_START_ROW, lastDataRow + 1);

    const values = rows.map(r => ([
      '', // SEQ（後で式）
      String(r.name || ''),
      String(r.yosana || ''),
      String(r.gender || ''),
      String(r.attendType || ''),
      String(r.regType || '正規'),
      '' // 備考
    ]));

    sh.getRange(startRow, 1, values.length, PDB_MAX_COLS).setValues(values);

    sh.getRange(startRow, 1, values.length, 1).setFormulas(
      Array.from({ length: values.length }, () => ['=ROW()-1'])
    );

    // ピボットのソース範囲を「実データ行」までに更新
    if (typeof Sheets !== 'undefined') {
      const newLast = findLastDataRowByCols_(sh, PDB_DATA_START_ROW, PDB_MAX_COLS);
      const endRowIndex = Math.max(1, newLast); // 0-based, endRowIndex は最終行+1
      updateParticipantsPivot_(ss.getId(), sh.getSheetId(), endRowIndex);
    }

    return { ok: true, message: values.length + '件インポートしました' };
  });
  });
}
