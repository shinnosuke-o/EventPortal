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
  APP_ICON_URL: 'https://drive.google.com/uc?export=view&id=1s-dJkHqSazwVEBBFWGnSCJIHW9TlCSk3&.png',
  AUDIT_SS_ID: '1e-JnrPLY_IDg7T_C6M9DvYoQyWn3mCIPM9j4AfahIbc',
  AUDIT_SHEET_NAME: '監査ログ',
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


/**
 * 共通ガード（サーバ側）
 * - Webアプリ利用許可フォルダにアクセスできないユーザーは拒否
 * - 変更系APIはここを必ず通す（各API冒頭で呼ぶ）
 */
function guard_(payload){
  if (!canAccessAllowFolder_()) throw new Error('アクセス権がありません。正しいGoogleアカウントでログインしてください。');
  // ここに将来CSRF等を追加する場合も guard_ に集約
  return true;
}

/**
 * 監査ログ（変更系のみ）
 * - 監査ログSSの「監査ログ」シートに追記
 */
function auditLog_(entry){
  try{
    const ss = SpreadsheetApp.openById(CONFIG.AUDIT_SS_ID);
    const sh = ss.getSheetByName(CONFIG.AUDIT_SHEET_NAME) || ss.insertSheet(CONFIG.AUDIT_SHEET_NAME);

    // ヘッダが無ければ作成
    if (sh.getLastRow() === 0){
      sh.appendRow(['timestamp','userEmail','fn','action','target','ok','message','payload']);
    }

    sh.appendRow([
      entry.timestamp || new Date(),
      entry.userEmail || '',
      entry.fn || '',
      entry.action || '',
      entry.target || '',
      entry.ok ? 'TRUE' : 'FALSE',
      entry.message || '',
      entry.payload || ''
    ]);
  }catch(e){
    // 監査ログ失敗で本処理を落とさない（運用優先）
    console.error('auditLog_ failed:', e);
  }
}

function getUserEmail_(){
  try{
    const email = Session.getActiveUser().getEmail();
    return email || '';
  }catch(e){
    return '';
  }
}

function summarizePayload_(payload){
  try{
    const seen = new WeakSet();
    const replacer = (k,v)=>{
      // base64 / 大きいデータを縮約
      if (typeof v === 'string'){
        if (v.length > 2000) return `[string:${v.length}]`;
        // ありがちなキー名
        const key = String(k||'').toLowerCase();
        if (key.includes('base64') || key.includes('filedata') || key.includes('blob') || key.includes('content')){
          return `[string:${v.length}]`;
        }
        return v;
      }
      if (v && typeof v === 'object'){
        if (seen.has(v)) return '[circular]';
        seen.add(v);
      }
      return v;
    };
    return JSON.stringify(payload ?? {}, replacer);
  }catch(e){
    return '[unserializable]';
  }
}

function inferAction_(fnName, payload){
  const n = String(fnName||'');
  if (n.includes('upload')) return 'UPLOAD';
  if (n.includes('import')) return 'IMPORT';
  if (n.includes('create')) return 'CREATE';
  if (n.includes('bulk') || n.includes('update')) return 'UPDATE';
  if (n.includes('upsert')) return 'UPSERT';
  return 'WRITE';
}

function inferTarget_(fnName, payload){
  const p = payload || {};
  // よくあるターゲット候補
  if (p.rowNumber) return `row:${p.rowNumber}`;
  if (p?.data?.rowNumber) return `row:${p.data.rowNumber}`;
  if (p.sheetName) return `sheet:${p.sheetName}`;
  if (p.eventSheetName) return `sheet:${p.eventSheetName}`;
  if (p.eventName) return `event:${p.eventName}`;
  if (p.fileId) return `file:${p.fileId}`;
  if (p.url) return `url:${p.url}`;
  return '';
}

/**
 * 変更系API用：排他＋監査ログ
 */
function withWriteLockAndAudit_(fnName, payload, handler){
  const lock = LockService.getScriptLock();
  const email = getUserEmail_();
  const action = inferAction_(fnName, payload);
  const target = inferTarget_(fnName, payload);
  const started = Date.now();

  let result;
  let ok = false;
  let message = '';

  // 30秒待つ（必要なら調整）
  lock.waitLock(30000);
  try{
    result = handler();
    ok = !!result?.ok;
    message = String(result?.message || '');
    return result;
  }catch(e){
    ok = false;
    message = e?.message || String(e);
    throw e;
  }finally{
    try{ lock.releaseLock(); }catch(e){}
    auditLog_({
      timestamp: new Date(),
      userEmail: email,
      fn: fnName,
      action,
      target,
      ok,
      message,
      payload: summarizePayload_(payload),
      ms: Date.now() - started
    });
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

/***********************
 * USER (who is accessing)
 ***********************/
function getUserEmail_(){
  // NOTE:
  // - Webアプリの「実行ユーザー」が USER_ACCESSING のとき、これでアクセス中ユーザーが取れます。
  // - 個人Gmail等で Session から取れない場合は、OAuth token から userinfo を引いて補完します。
  try{
    const a = Session.getActiveUser().getEmail();
    if (a) return a;
  }catch(e){}
  try{
    const e = Session.getEffectiveUser().getEmail();
    if (e) return e;
  }catch(e){}
  try{
    const cache = CacheService.getUserCache();
    const cached = cache.get('me_email');
    if (cached) return cached;

    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300){
      const obj = JSON.parse(resp.getContentText() || '{}');
      const email = String(obj.email || '').trim();
      if (email){
        cache.put('me_email', email, 21600); // 6h
        return email;
      }
    }
  }catch(e){}
  return '';
}

function getActorEmail_(payload){
  const email = getUserEmail_();
  if(email) return email;
  const c = String(payload && payload.__clientEmail || '').trim();
  return c;
}

function api_getMe(){
  return safeApi_(() => {
    if(!canAccessAllowFolder_()) return { ok:false, message:'アクセス権限がありません。' };
    const email = getUserEmail_();
    return { ok:true, email };
  });
}


function getBaseUrl_(){
  return ScriptApp.getService().getUrl();
}

function getManifestUrl_(){
  return getBaseUrl_() + '?manifest=1';
}

function getServiceWorkerUrl_(){
  return getBaseUrl_() + '?sw=1';
}

function buildManifest_(){
  const startUrl = getPageUrl_('home');
  const iconUrl = CONFIG.APP_ICON_URL;
  const manifest = {
    name: 'イベント班ポータル',
    short_name: 'EventPortal',
    start_url: startUrl,
    display: 'standalone',
    background_color: '#0b1220',
    theme_color: '#0b1220',
    icons: [
      { src: iconUrl, sizes: '192x192', type: 'image/png' },
      { src: iconUrl, sizes: '512x512', type: 'image/png' }
    ]
  };
  return ContentService
    .createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildServiceWorker_(){
  const startUrl = getPageUrl_('home');
  const sw = `const CACHE_NAME = 'eventportal-pwa-v1';\n` +
    `const START_URL = ${JSON.stringify(startUrl)};\n` +
    `self.addEventListener('install', (event) => {\n` +
    `  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll([START_URL])));\n` +
    `});\n` +
    `self.addEventListener('activate', (event) => {\n` +
    `  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))));\n` +
    `});\n` +
    `self.addEventListener('fetch', (event) => {\n` +
    `  if (event.request.method !== 'GET') return;\n` +
    `  if (event.request.mode === 'navigate') {\n` +
    `    event.respondWith(fetch(event.request).catch(() => caches.match(START_URL)));\n` +
    `  }\n` +
    `});\n`;
  return ContentService
    .createTextOutput(sw)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doGet(e) {
  if(!canAccessAllowFolder_()){
    // IMPORTANT:
    // - Avoid authuser links here. They can trigger script.google.com/accounts recursion and Bad Request 400.
    // - Use ServiceLogin with prompt=select_account to force account selection.
    const continueUrl = getBaseUrl_() + '?page=home&ts=' + Date.now();
    const loginUrl =
      'https://accounts.google.com/ServiceLogin?service=script&passive=true&prompt=select_account&continue=' +
      encodeURIComponent(continueUrl);
    const chooserUrl =
      'https://accounts.google.com/AccountChooser?service=lso&continue=' + encodeURIComponent(continueUrl);
    const logoutUrl = 'https://accounts.google.com/Logout?continue=' + encodeURIComponent(continueUrl);

    return HtmlService.createHtmlOutput(`
      <div style="font-family:system-ui; padding:24px; line-height:1.7; max-width:720px; margin:0 auto;">
        <h2 style="margin:0 0 12px;">アクセス権がありません</h2>
        <p style="margin:0 0 16px;">
          このWebアプリは指定フォルダにアクセスできるユーザーのみ利用できます。<br>
          <b>別のGoogleアカウントで開き直す</b>には、下のボタンを押してください。
        </p>

        <div style="display:flex; gap:12px; flex-wrap:wrap; margin:16px 0 20px;">
          <a href="${loginUrl}" target="_top"
             style="display:inline-block; padding:10px 14px; border-radius:10px; background:#111; color:#fff; text-decoration:none;">
            別のGoogleアカウントで開く（推奨）
          </a>
          <a href="${chooserUrl}" target="_top"
             style="display:inline-block; padding:10px 14px; border-radius:10px; border:1px solid #111; color:#111; text-decoration:none;">
            アカウント選択画面を開く
          </a>
          <a href="${continueUrl}" target="_top"
             style="display:inline-block; padding:10px 14px; border-radius:10px; border:1px solid #ccc; color:#111; text-decoration:none;">
            再読み込み
          </a>
        </div>

        <details style="margin:14px 0;">
          <summary style="cursor:pointer;">うまく切り替わらない場合</summary>
          <div style="margin-top:10px;">
            <div style="margin:8px 0;">
              <div style="font-size:13px; color:#444;">すべてのGoogleアカウントからサインアウトしてやり直す</div>
              <div><a href="${logoutUrl}" target="_top">Googleからサインアウトして開き直す</a></div>
            </div>
            <div style="margin:8px 0; font-size:13px; color:#444;">
              ※アカウント選択後も権限エラーになる場合、そのアカウントに指定フォルダの閲覧権限が付与されていません。
            </div>
          </div>
        </details>
      </div>
    `).setTitle('Access Denied');
  }

  if (e && e.parameter) {
    if (e.parameter.manifest === '1') return buildManifest_();
    if (e.parameter.sw === '1') return buildServiceWorker_();
  }

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.page = page;
  const faviconUrl = CONFIG.APP_ICON_URL;
  // iOSホーム画面アイコンは apple-touch-icon を優先する
  // const appleTouchIconUrl = 'https://drive.google.com/uc?export=view&id=1s-dJkHqSazwVEBBFWGnSCJIHW9TlCSk3';
  tpl.faviconUrl = faviconUrl;
  tpl.appleTouchIconUrl = faviconUrl;

  // ★追加：ホームへの絶対URL
  tpl.homeUrl = getPageUrl_('home');
  tpl.manifestUrl = getManifestUrl_();
  tpl.swUrl = getServiceWorkerUrl_();
  

  return tpl.evaluate()
    .setTitle('イベ班ポータル(Event Portal)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
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




// tokeninfo で検証（最短・堅牢）






