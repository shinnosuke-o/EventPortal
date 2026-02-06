/***********************
 * CONFIG
 ***********************/
const CONFIG = {
  TZ: 'Asia/Tokyo',

  MASTER_SS_ID: '1BL08rl8USs2cbIgjyD_9EDP9xL_ilRKwF1pStHv7El4',
  GUIDE_SHEET_NAME: '★案内テンプレ自動作成★',

  PUBLIC_FOLDER_ID: '1-8mgI0sqmySqEbANonvkdl8_83V6M9Eb',

  TASK_SS_ID: '1-wyfN0krk_1o46WNlnh9pFNVkYQnYG3exC09E2XErT0',
  TASK_SHEET_NAME: '全体タスク',

  PARTICIPANTS_DB_ID: '1krJtAbxISUZOIIiByRcySrFi9ua5ZdwZzRsWSimrm8w',

  // 出欠フォーム作成の固定（後でUI化してもOK）
  ATTEND_FORM_FOLDER_ID: '1w7QeWmi0rWRSkVf9fsNTe6_dTlNm9VYa',

  // シートフィルタ
  EXCLUDE_MARK: '★',
  DANCER_MASTER_PREFIX: '★踊り子一覧',

  ACCOUNT_SHEET_NAME: '会計',
  RECEIPT_FOLDER_ID: '1op_LjvPTGh6qhC0kRkpp-ULjupoECqML', // 領収書保存先

  ALLOW_FOLDER_ID: '1A3toeK8PDXkTYmQVMJr6NGHNjRVMTkIp', // ←指定フォルダID
};


/***********************
 * UTIL
 ***********************/
function safeApi_(fn) {
  try {
    return fn();
  } catch (e) {
    return { ok: false, message: e?.message || String(e) };
  }
}

function openSS_(id) {
  return SpreadsheetApp.openById(id);
}

function mustSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません：' + name);
  return sh;
}

function escapeSheetName_(name) {
  // A1 notation: single quotes required if name has spaces/symbols. Single quotes are escaped by doubling.
  return "'" + String(name || '').replace(/'/g, "''") + "'";
}

function toIso_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return v.toISOString();
  }
  return String(v);
}

function norm_(v) {
  return String(v ?? '')
    .replace(/\u3000/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();
}

function appendRowByColumnB_(sheet, startRow /* usually 2 */) {
  // B列ベースで「実データの最終行」を探して末尾+1を返す
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  const values = sheet.getRange(startRow, 2, lastRow - (startRow - 1), 1).getValues(); // B
  let last = startRow - 1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v !== '' && v !== null) last = startRow + i;
  }
  return Math.max(startRow, last + 1);
}

function findLastDataRowByCols_(sheet, startRow, endCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;

  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, endCol).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (row.some(v => v !== '' && v !== null)) return startRow + i;
  }
  return startRow - 1;
}

function formatMDW_(d) {
  const wd = ['日','月','火','水','木','金','土'];
  const md = Utilities.formatDate(d, CONFIG.TZ, 'M/d');
  return `${md}(${wd[d.getDay()]})`;
}

function sameDay_(d1, d2) {
  if (!d1 || !d2) return false;
  return Utilities.formatDate(d1, CONFIG.TZ, 'yyyyMMdd') === Utilities.formatDate(d2, CONFIG.TZ, 'yyyyMMdd');
}

function ymd_(d) {
  if (!d) return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd');
  return Utilities.formatDate(d, CONFIG.TZ, 'yyyyMMdd');
}


/***********************
 * WEB APP / HTML
 ***********************/
/**
 * フォルダにアクセスできるか判定
 * - そのフォルダの「閲覧者/コメント可/編集者」なら true
 * - アクセス権が無ければ例外になるので false
 */
function canAccessAllowFolder_(){
  try{
    DriveApp.getFolderById(CONFIG.ALLOW_FOLDER_ID).getName(); // 権限チェック目的
    return true;
  }catch(e){
    return false;
  }
}

function doGet(e) {
  if(!canAccessAllowFolder_()){
    return HtmlService.createHtmlOutput(`
      <div style="font-family:system-ui; padding:24px; line-height:1.7;">
        <h2>アクセス権がありません</h2>
        <p>
          このWebアプリは指定フォルダにアクセスできるユーザー（イベント班）のみ利用できます。<br>
          正しいGoogleアカウントでログインしているか確認してください。
        </p>
      </div>
    `).setTitle('Access Denied');
  }

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.page = page;
  const faviconUrl = 'https://drive.google.com/uc?export=view&id=1s-dJkHqSazwVEBBFWGnSCJIHW9TlCSk3&.png';
  // iOSホーム画面アイコンは apple-touch-icon を優先する
  // const appleTouchIconUrl = 'https://drive.google.com/uc?export=view&id=1s-dJkHqSazwVEBBFWGnSCJIHW9TlCSk3';
  tpl.faviconUrl = faviconUrl;
  tpl.appleTouchIconUrl = faviconUrl;

  // ★追加：ホームへの絶対URL
  tpl.homeUrl = getPageUrl_('home');
  

  return tpl.evaluate()
    .setTitle('イベ班ポータル(Event Portal)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl(faviconUrl);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function api_getPageHtml(page) {
  return safeApi_(() => {
    const allowed = ['home', 'formWizard', 'publicSheet', 'noticeTemplate', 'attendanceSummary', 'task', 'accounting'];
    const p = allowed.includes(page) ? page : 'home';
    const html = HtmlService.createHtmlOutputFromFile('pages/' + p).getContent();
    return { ok: true, data: html };
  });
}

function getPageUrl_(page) {
  const base = ScriptApp.getService().getUrl();
  return base + '?page=' + encodeURIComponent(page);
}


/***********************
 * DRIVE (Advanced Service)
 ***********************/
function assertAuthorized_(){
  if(!canAccessAllowFolder_()){
    throw new Error('UNAUTHORIZED');
  }
}

function requireDrive_() {
  if (typeof Drive === 'undefined') {
    throw new Error(
      'Drive API（高度なGoogleサービス）が無効です。「サービス」から Drive API を追加してください。'
    );
  }
}

function moveFileToFolder_(fileId, folderId) {
  requireDrive_();

  const file = Drive.Files.get(fileId, {
    fields: 'parents',
    supportsAllDrives: true
  });

  const prevParents = (file.parents || []).map(p => p.id).join(',');

  Drive.Files.update(
    {},      // resource
    fileId,
    null,    // mediaData
    {
      addParents: folderId,
      removeParents: prevParents,
      supportsAllDrives: true
    }
  );
}

function api_getCurrentUser(params){
  try{
    const accessToken = params && params.accessToken;
    if(!accessToken) return { ok:false, message:'NO_TOKEN' };

    const res = UrlFetchApp.fetch('https://openidconnect.googleapis.com/v1/userinfo', {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    if(res.getResponseCode() !== 200){
      return { ok:false, message:'userinfo failed: ' + res.getContentText() };
    }

    const data = JSON.parse(res.getContentText());
    return {
      ok:true,
      email: data.email || '',
      name: data.name || '',
      picture: data.picture || ''
    };
  }catch(e){
    return { ok:false, message: e.message || String(e) };
  }
}


// tokeninfo で検証（最短・堅牢）
function verifyIdToken_(idToken) {
  const clientId = PropertiesService.getScriptProperties().getProperty('GIS_CLIENT_ID');
  if (!clientId) return { ok:false, message:'GIS_CLIENT_ID が未設定です' };

  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  const body = res.getContentText() || '{}';

  let data = {};
  try { data = JSON.parse(body); } catch (e) {}

  if (code !== 200) {
    return { ok:false, message:'tokeninfo failed: ' + body };
  }
  if (data.aud !== clientId) {
    return { ok:false, message:'aud mismatch' };
  }
  const verified = (data.email_verified === true || data.email_verified === 'true');
  if (!verified) {
    return { ok:false, message:'email not verified' };
  }

  return {
    ok:true,
    email: data.email,
    name: data.name,
    picture: data.picture,
    sub: data.sub
  };
}


/***********************
 * GUIDE SHEET
 ***********************/
function api_listGuideEvents() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sh = mustSheet_(ss, CONFIG.GUIDE_SHEET_NAME);

    const lastRow = sh.getLastRow();
    if (lastRow < 3) return { ok: true, data: [] };

    const values = sh.getRange(3, 2, lastRow - 2, 1).getValues(); // B3:B
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
    if (lastRow < 3) return { ok: false, message: 'イベント一覧がありません。', data: null };

    // B3:J（9列）
    const values = sh.getRange(3, 2, lastRow - 2, 9).getValues();

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

  sheet.getRange(row, 10).setValue(publicUrl); // J
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

  sheet.getRange(appendRow, 2, 1, 10).setValues(values);
}

function api_pingGuide() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.MASTER_SS_ID);
    const sh = ss.getSheetByName(CONFIG.GUIDE_SHEET_NAME);
    return {
      ok: true,
      time: new Date().toISOString(),
      canOpenMaster: !!ss,
      hasGuideSheet: !!sh,
      guideLastRow: sh ? sh.getLastRow() : null
    };
  });
}


/***********************
 * FORM CREATE
 ***********************/
function buildFormDescription_({ dateFrom, dateTo, place, deadline, practice, condition }) {
  const dateStr = (() => {
    if (!dateFrom && !dateTo) return '';
    if (dateFrom && dateTo) {
      if (sameDay_(dateFrom, dateTo)) return formatMDW_(dateFrom);
      return `${formatMDW_(dateFrom)}〜${formatMDW_(dateTo)}`;
    }
    if (dateFrom) return formatMDW_(dateFrom);
    return formatMDW_(dateTo);
  })();

  const deadlineStr = deadline ? formatMDW_(deadline) : '';

  return [
    `${dateStr}@${place || ''}`,
    '',
    `⚠️回答期限：${deadlineStr}⚠️`,
    '未定・不参加の場合も必ず回答お願いします！',
    '',
    '■練習日(予定)',
    practice || '',
    '■参加条件',
    condition || ''
  ].join('\n');
}

function buildFormFileName_({ eventName, dateFrom, dateTo }) {
  const fromStr = ymd_(dateFrom);
  const toStr = ymd_(dateTo || dateFrom);

  const same = (dateFrom && (dateTo || dateFrom))
    ? ymd_(dateFrom) === ymd_(dateTo || dateFrom)
    : true;

  return same
    ? `${fromStr}_${eventName}出欠確認フォーム`
    : `${fromStr}-${toStr}_${eventName}出欠確認フォーム`;
}

function findNewSheetAfterDestination_(ss, beforeSheetIds, maxTry, waitMs) {
  const before = new Set(beforeSheetIds);

  for (let i = 0; i < maxTry; i++) {
    SpreadsheetApp.flush();
    const sheets = ss.getSheets();
    const added = sheets.find(s => !before.has(s.getSheetId()));
    if (added) return added;
    Utilities.sleep(waitMs);
  }
  return null;
}

function moveSheetToIndex_(ss, sheet, index1based) {
  const sheets = ss.getSheets();
  const count = sheets.length;
  const dest = Math.min(Math.max(1, index1based), count);

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(dest);
}

function addItemToForm_(form, r) {
  const title = (r.title || '').trim();
  const type = (r.type || '').trim();
  const required = !!r.required;
  const help = (r.help || '').trim();

  const choices = [r.opt1, r.opt2, r.opt3, r.opt4, r.opt5]
    .map(v => (v || '').trim())
    .filter(v => v !== '');

  let item;

  switch (type) {
    case '記述式（短文）':
      item = form.addTextItem().setTitle(title);
      break;
    case '段落':
      item = form.addParagraphTextItem().setTitle(title);
      break;
    case 'ラジオボタン':
      item = form.addMultipleChoiceItem().setTitle(title);
      item.setChoiceValues(choices.length ? choices : ['選択肢を入力してください']);
      break;
    case 'チェックボックス':
      item = form.addCheckboxItem().setTitle(title);
      item.setChoiceValues(choices.length ? choices : ['選択肢を入力してください']);
      break;
    case 'プルダウン':
      item = form.addListItem().setTitle(title);
      item.setChoiceValues(choices.length ? choices : ['選択肢を入力してください']);
      break;
    case '日付':
      item = form.addDateItem().setTitle(title);
      break;
    case '時刻':
      item = form.addTimeItem().setTitle(title);
      break;
    default:
      item = form.addTextItem().setTitle(title + '（※タイプ不明）');
  }

  if (help) item.setHelpText(help);
  item.setRequired(required);
}

function api_createAttendanceForm(payload) {
  return safeApi_(() => {
    const DEST_SS_ID = CONFIG.MASTER_SS_ID;
    const FOLDER_ID = CONFIG.ATTEND_FORM_FOLDER_ID;

    const b = payload?.basic || {};
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];

    const eventName = (b.eventName || '').trim();
    if (!eventName) return { ok: false, message: 'イベント名は必須です。' };

    const dateFrom = b.dateFrom ? new Date(b.dateFrom) : null;
    const dateTo = b.dateTo ? new Date(b.dateTo) : null;
    const deadline = b.deadline ? new Date(b.deadline) : null;

    if (dateFrom && dateTo && dateFrom.getTime() > dateTo.getTime()) {
      return { ok: false, message: '開催日のFrom/Toが不正です（From <= To）。' };
    }
    if (deadline && dateFrom && deadline.getTime() > dateFrom.getTime()) {
      return { ok: false, message: '回答期限は開催日(From)以前にしてください。' };
    }

    const ss = openSS_(DEST_SS_ID);
    if (ss.getSheets().some(s => s.getName() === eventName)) {
      return { ok: false, message: `同名シート「${eventName}」が既に存在します。` };
    }

    // 入力検証（空行は無視）
    const activeRows = rows
      .map((r, i) => ({ ...r, _i: i + 1 }))
      .filter(r => (r.title || '').trim() !== '' || (r.type || '').trim() !== '');

    for (const r of activeRows) {
      if (!(r.title || '').trim()) return { ok: false, message: `フォーム内容：${r._i}行目「質問内容」が必須です。` };
      if (!(r.type || '').trim()) return { ok: false, message: `フォーム内容：${r._i}行目「質問タイプ」が必須です。` };
    }

    // フォーム作成
    const formTitle = `${eventName}出欠確認`;
    const form = FormApp.create(formTitle).setTitle(formTitle);

    const place = (b.place || '').trim();
    const practice = (b.practice || '').trim();
    const condition = (b.condition || '').trim();

    const desc = buildFormDescription_({ dateFrom, dateTo, place, deadline, practice, condition });
    form.setDescription(desc);

    // Drive上のファイル名
    DriveApp.getFileById(form.getId()).setName(buildFormFileName_({ eventName, dateFrom, dateTo }));

    // setDestination 差分で回答シート取得
    const beforeSheetIds = ss.getSheets().map(s => s.getSheetId());
    form.setDestination(FormApp.DestinationType.SPREADSHEET, DEST_SS_ID);

    const responseSheet = findNewSheetAfterDestination_(ss, beforeSheetIds, 15, 400);
    if (!responseSheet) throw new Error('回答連携シートを特定できませんでした。');

    responseSheet.setName(eventName);
    moveSheetToIndex_(ss, responseSheet, 6);

    // フォルダへ移動（共有ドライブ対応）
    moveFileToFolder_(form.getId(), FOLDER_ID);

    // 質問追加
    for (const r of activeRows) addItemToForm_(form, r);

    // 案内テンプレ追記
    appendEventToGuideTemplate_({
      ssId: DEST_SS_ID,
      sheetName: CONFIG.GUIDE_SHEET_NAME,
      eventName,
      dateFrom,
      dateTo,
      place,
      deadline,
      practice,
      condition,
      publishedUrl: form.getPublishedUrl(),
      notify: !!b.notify
    });

    return {
      ok: true,
      message: 'フォームを作成しました。',
      formId: form.getId(),
      editUrl: form.getEditUrl(),
      publishedUrl: form.getPublishedUrl()
    };
  });
}


/***********************
 * PUBLIC CHECK SHEET
 ***********************/
function api_createPublicCheckSheet(payload) {
  return safeApi_(() => {
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

    const selectCols = cols.map(n => 'Col' + n).join(',');
    const query = `select ${selectCols}`;

    const destTitle = `${eventName}出欠確認_公開用`;
    const dest = SpreadsheetApp.create(destTitle);
    const destId = dest.getId();
    const destUrl = dest.getUrl();

    const first = dest.getSheets()[0];
    first.setName(eventName);

    const rangeA1 = `${escapeSheetName_(eventName)}!A:ZZ`;
    const formula = `=QUERY(IMPORTRANGE("${CONFIG.MASTER_SS_ID}","${rangeA1}"),"${query}",1)`;
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


/***********************
 * HEADERS / LIST (used by publicSheet etc)
 ***********************/
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

    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = header
      .map((name, i) => ({ index: i + 1, name: String(name || '').trim() }))
      .filter(x => x.name !== '');
    return { ok: true, data };
  });
}


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
 * 既存API名が二重だったので「片方に統一」しつつ互換用の別名も残す
 ***********************/
function api_listParticipantDbSheets() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
    return { ok: true, sheets: ss.getSheets().map(s => s.getName()) };
  });
}

// 互換：古いクライアントが呼ぶ可能性
function api_listParticipantsDbSheets() {
  // 以前は配列返しだったが、呼び出し箇所が少ないなら ok形式に合わせても良い。
  // ただし互換優先で「配列」を維持したい場合は return ss.getSheets().map... に戻してください。
  const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
  return ss.getSheets().map(s => s.getName());
}

function api_createParticipantsDbSheet(payload) {
  return safeApi_(() => {
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
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);

    // Table作成（Sheets Advanced Service）
    if (typeof Sheets === 'undefined') {
      return { ok: false, message: 'Sheets API（高度なGoogleサービス）が無効です。「サービス」から Sheets API を追加してください。' };
    }

    const sheetId = sh.getSheetId();
    const maxCols = 7; // A:G

    const requests = [{
      addTable: {
        table: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: maxCols
          },
          name: sheetName
        }
      }
    },{
      updateCells: {
        start: { sheetId, rowIndex: 1, columnIndex: 8 }, // I2
        rows: [{
          values: [{
            pivotTable: {
              source: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: 1, // 初期はヘッダ行のみ
                startColumnIndex: 0,
                endColumnIndex: maxCols
              },
              rows: [{
                sourceColumnOffset: 4, // 参加区分（E列）
                sortOrder: 'DESCENDING',
                showTotals: true
              }],
              columns: [{
                sourceColumnOffset: 3, // 性別（D列）
                sortOrder: 'DESCENDING',
                showTotals: true
              }],
              values: [{
                summarizeFunction: 'COUNTA',
                sourceColumnOffset: 1 // 名前（B列）
              }]
            }
          }]
        }],
        fields: 'pivotTable'
      }
    }];

    Sheets.Spreadsheets.batchUpdate({ requests }, ss.getId());

    return { ok: true, message: 'シートを作成しました：' + sheetName };
  });
}

// 互換：別名（古い関数名）
function api_createParticipantDbSheet(payload) {
  return api_createParticipantsDbSheet(payload);
}

function api_importParticipants(payload) {
  return safeApi_(() => {
    const destSheetName = String(payload?.destSheetName || '').trim();
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];

    if (!destSheetName) return { ok: false, message: 'destSheetName is required' };
    if (rows.length === 0) return { ok: false, message: 'rows is empty' };

    const ss = openSS_(CONFIG.PARTICIPANTS_DB_ID);
    const sh = ss.getSheetByName(destSheetName);
    if (!sh) return { ok: false, message: 'インポート先シートが見つかりません：' + destSheetName };

    const lastDataRow = findLastDataRowByCols_(sh, 2, 7); // A:G を見て最終行判定
    const startRow = Math.max(2, lastDataRow + 1);

    const values = rows.map(r => ([
      '', // SEQ（後で式）
      String(r.name || ''),
      String(r.yosana || ''),
      String(r.gender || ''),
      String(r.attendType || ''),
      String(r.regType || '正規'),
      '' // 備考
    ]));

    sh.getRange(startRow, 1, values.length, 7).setValues(values);

    sh.getRange(startRow, 1, values.length, 1).setFormulas(
      Array.from({ length: values.length }, () => [`=ROW()-1`])
    );

    // ピボットのソース範囲を「実データ行」までに更新
    if (typeof Sheets !== 'undefined') {
      const sheetId = sh.getSheetId();
      const newLast = findLastDataRowByCols_(sh, 2, 7);
      const endRowIndex = Math.max(1, newLast); // 0-based, endRowIndex は最終行+1
      const requests = [{
        updateCells: {
          start: { sheetId, rowIndex: 1, columnIndex: 8 }, // I2
          rows: [{
            values: [{
              pivotTable: {
                source: {
                  sheetId,
                  startRowIndex: 0,
                  endRowIndex,
                  startColumnIndex: 0,
                  endColumnIndex: 7
                },
                rows: [{
                  sourceColumnOffset: 4,
                  sortOrder: 'DESCENDING',
                  showTotals: true
                }],
                columns: [{
                  sourceColumnOffset: 3,
                  sortOrder: 'DESCENDING',
                  showTotals: true
                }],
                values: [{
                  summarizeFunction: 'COUNTA',
                  sourceColumnOffset: 1
                }]
              }
            }]
          }],
          fields: 'pivotTable'
        }
      }];
      Sheets.Spreadsheets.batchUpdate({ requests }, ss.getId());
    }

    return { ok: true, message: `${values.length}件インポートしました` };
  });
}


/***********************
 * TASKS
 ***********************/
function api_getTasks() {
  return safeApi_(() => {
    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = mustSheet_(ss, CONFIG.TASK_SHEET_NAME);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, message: '', tasks: [] };

    const values = sh.getRange(2, 2, lastRow - 1, 8).getValues(); // B:I
    const tasks = [];

    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      const title = String(r[0] || '').trim();
      if (!title) continue;

      tasks.push({
        rowNumber: i + 2,
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
      if (!Number.isFinite(rowNumber) || rowNumber < 2) {
        return { ok: false, message: '修正対象行が不正です。', rowNumber: null };
      }
      sh.getRange(rowNumber, 2, 1, 8).setValues(rowValues);
      return { ok: true, message: '更新しました。', rowNumber };
    }

    // create（B列ベースで末尾探索）
    const appendRow = appendRowByColumnB_(sh, 2);
    sh.getRange(appendRow, 2, 1, 8).setValues(rowValues);

    return { ok: true, message: '追加しました。', rowNumber: appendRow };
  });
}

/***********************
 * ACCOUNTING
 ***********************/
function api_listExpenses() {
  const out = { ok:false, message:'', expenses:[] };

  try {
    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = ss.getSheetByName(CONFIG.ACCOUNT_SHEET_NAME);
    if (!sh) return { ok:false, message:'シートが見つかりません：' + CONFIG.ACCOUNT_SHEET_NAME, expenses:[] };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, expenses:[] };

    // ★B列（経費内容）で「本当の最終行」を決める（装飾だけの行を除外）
    const bVals = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B
    let lastDataRow = 1; // 実行時は 2行目〜なので「シート上の行番号」を作る
    for (let i = bVals.length - 1; i >= 0; i--) {
      if (String(bVals[i][0] ?? '').trim() !== '') {
        lastDataRow = i + 2; // シート行番号
        break;
      }
    }
    if (lastDataRow < 2) return { ok:true, expenses:[] };

    // ★A:I（9列）を読む：A=SEQ, B=経費内容, C=説明, D=支払者, E=支払日, F=領収書, G=ステータス, H=精算依頼日, I=精算日
    const numRows = lastDataRow - 1; // 2行目〜lastDataRow
    const values = sh.getRange(2, 1, numRows, 9).getValues(); // A2:I

    const expenses = [];
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      const sheetRow = i + 2; // ★実シート行番号（絶対にズレない）

      const title = String(r[1] ?? '').trim(); // B
      if (!title) continue; // 空行は無視（ただし rowNumber は sheetRow で保持してるのでズレない）

      expenses.push({
        seq: (r[0] !== '' && r[0] != null) ? Number(r[0]) : null, // A
        rowNumber: sheetRow,                 // ★実行用：実シート行番号
        title: title,                        // B
        desc:  String(r[2] ?? ''),           // C
        payer: String(r[3] ?? ''),           // D
        payDate: toIso_(r[4]),               // E
        receiptUrl: String(r[5] ?? ''),      // F
        status: String(r[6] ?? ''),          // G
        requestDate: toIso_(r[7]),           // H
        settleDate: toIso_(r[8])             // I
      });
    }

    out.ok = true;
    out.expenses = expenses;
    return out;

  } catch(e){
    out.ok = false;
    out.message = e?.message || String(e);
    return out;
  }
}

function api_uploadReceipt(payload){
  try{
    const mime = String(payload?.mime||'');
    const base64 = String(payload?.base64||'');
    if(!mime || !base64) return { ok:false, message:'file is required' };

    const payDate = String(payload?.payDate||'yyyyMMdd').trim(); // YYYY-MM-DD
    const payer = String(payload?.payer||'支払者').trim();
    const title = String(payload?.title||'経費内容').trim();

    const oldReceiptUrl = String(payload.oldReceiptUrl || '').trim();
    if (oldReceiptUrl) {
      // 旧ファイル削除 → 失敗したら中断（誤って二重管理にならないように）
      deleteReceiptByUrl_(oldReceiptUrl);
    }

    const yyyymmdd = payDate ? payDate.replaceAll('-','') : Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd');

    // ファイル名に使えない文字除去（/ \ : * ? " < > | など）
    const safe = (s)=>String(s||'')
      .replace(/[\\\/:\*\?"<>\|]/g,'')
      .replace(/\s+/g,' ')
      .trim()
      .slice(0, 40); // 長すぎ防止

    const ext = mime.includes('pdf') ? 'pdf' : 'png'; // 必要なら mime で細分化
    const fileName = `${yyyymmdd}_${safe(payer)}_${safe(title)}.${ext}`;

    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mime, fileName);

    const folder = DriveApp.getFolderById(CONFIG.RECEIPT_FOLDER_ID);
    const file = folder.createFile(blob);

    // 必要なら共有設定（閲覧用リンク）
    // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { ok:true, url:file.getUrl(), id:file.getId(), name:fileName };

  }catch(e){
    return { ok:false, message:e?.message||String(e) };
  }
}

function api_upsertExpense(payload){
  const out = { ok:false, message:'', rowNumber:null };

  try{
    const mode = String(payload?.mode || 'create'); // create/edit
    const data = payload?.data || {};

    const title = String(data.title||'').trim();
    if(!title) return { ok:false, message:'経費内容は必須です。' };

    const desc = String(data.desc||'');
    const payer = String(data.payer||'');
    const payDate = data.payDate ? new Date(String(data.payDate)+'T00:00:00') : null;

    const receiptUrl = String(data.receiptUrl||'');
    const status = String(data.status||'精算依頼前');

    let requestDate = data.requestDate ? new Date(String(data.requestDate)+'T00:00:00') : null;
    let settleDate  = data.settleDate ? new Date(String(data.settleDate)+'T00:00:00') : null;

    // ステータスに応じた日付自動補完（好みで）
    const today = new Date();
    if(status === '精算依頼済み' && !requestDate) requestDate = today;
    if(status === '精算済み' && !settleDate) settleDate = today;

    const ss = SpreadsheetApp.openById(CONFIG.TASK_SS_ID);
    const sh = ss.getSheetByName(CONFIG.ACCOUNT_SHEET_NAME);
    if(!sh) return { ok:false, message:'シートが見つかりません：' + CONFIG.ACCOUNT_SHEET_NAME };

    const rowValues = [[ title, desc, payer, payDate, receiptUrl, status, requestDate, settleDate ]]; // B:I

    if(mode === 'edit'){
      const rowNumber = Number(payload?.rowNumber);
      if(!Number.isFinite(rowNumber) || rowNumber < 2) return { ok:false, message:'修正対象行が不正です。' };
      sh.getRange(rowNumber, 2, 1, 8).setValues(rowValues);

      out.ok = true;
      out.rowNumber = rowNumber;
      out.message = '更新しました。';
      return out;
    }

    // create: B列（経費内容）基準で末尾追加
    const lastDataRow = findLastDataRowByCol_(sh, 2, 2); // col=2(B), startRow=2
    const appendRow = Math.max(2, lastDataRow + 1);

    sh.getRange(appendRow, 2, 1, 8).setValues(rowValues);

    out.ok = true;
    out.rowNumber = appendRow;
    out.message = '追加しました。';
    return out;

  }catch(e){
    return { ok:false, message:e?.message || String(e) };
  }
}

function api_bulkUpdateExpenseStatus(payload){
  try{
    const rowNumbers = Array.isArray(payload?.rowNumbers) ? payload.rowNumbers : [];
    const newStatus = String(payload?.newStatus || '').trim();
    if(!rowNumbers.length) return { ok:false, message:'rowNumbers is empty' };
    if(!newStatus) return { ok:false, message:'newStatus is required' };

    const ss = SpreadsheetApp.openById(CONFIG.TASK_SS_ID);
    const sh = ss.getSheetByName(CONFIG.ACCOUNT_SHEET_NAME);
    if(!sh) return { ok:false, message:'シートが見つかりません：' + CONFIG.ACCOUNT_SHEET_NAME };

    const today = new Date();

    for(const rn of rowNumbers){
      const row = Number(rn);
      if(!Number.isFinite(row) || row < 2) continue;

      // G=ステータス(7列目/B起点だと6番目) → ここはRangeで直接指定が安全
      // B起点のため、G列は「2+5=7列目」ではなく、シートの列でGは7。
      sh.getRange(row, 7).setValue(newStatus); // G

      if(newStatus === '精算依頼済み'){
        sh.getRange(row, 8).setValue(today); // H 精算依頼日
      }
      if(newStatus === '精算済み'){
        sh.getRange(row, 9).setValue(today); // I 精算日
      }
    }

    return { ok:true, message:`${rowNumbers.length}件更新しました` };
  }catch(e){
    return { ok:false, message:e?.message || String(e) };
  }
}

function findLastDataRowByCol_(sh, col, startRow){
  const last = sh.getLastRow();
  if (last < startRow) return startRow - 1;

  // B列だけ取得して、最後に値がある行を探す
  const vals = sh.getRange(startRow, col, last - startRow + 1, 1).getValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    const v = vals[i][0];
    if (v !== '' && v !== null) return startRow + i;
  }
  return startRow - 1;
}

function getDriveFileIdFromUrl_(url){
  const s = String(url || '').trim();
  if(!s) return '';

  // patterns:
  // https://drive.google.com/file/d/<ID>/view
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if(m) return m[1];

  // https://drive.google.com/open?id=<ID>
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if(m) return m[1];

  // https://drive.google.com/uc?id=<ID>
  m = s.match(/\/uc\?id=([a-zA-Z0-9_-]{10,})/);
  if(m) return m[1];

  return '';
}

function deleteReceiptByUrl_(oldUrl){
  const fileId = getDriveFileIdFromUrl_(oldUrl);
  if(!fileId) throw new Error('既存領収書URLからファイルIDを取得できませんでした。');

  if (typeof Drive === 'undefined') {
    throw new Error('Drive API（高度なGoogleサービス）が無効です。「サービス」から Drive API を追加してください。');
  }

  // shared drive 対応
  Drive.Files.remove(fileId, { supportsAllDrives: true });
}


/***********************
 * DEBUG
 ***********************/
function api_debugReturn() {
  return { ok: true, time: new Date().toISOString() };
}
