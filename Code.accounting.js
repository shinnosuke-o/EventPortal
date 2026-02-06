/***********************
 * ACCOUNTING
 ***********************/
const ACCOUNT_DATA_START_ROW = 2;
const ACCOUNT_COL = {
  SEQ: 1,
  TITLE: 2,
  DESC: 3,
  PAYER: 4,
  PAY_DATE: 5,
  RECEIPT: 6,
  STATUS: 7,
  REQUEST_DATE: 8,
  SETTLE_DATE: 9
};
const ACCOUNT_COL_COUNT = 9; // A:I
const ACCOUNT_DATA_COLS_NO_SEQ = 8; // B:I
const ACCOUNT_STATUS_DEFAULT = '精算依頼前';
const ACCOUNT_STATUS_REQUESTED = '精算依頼済み';
const ACCOUNT_STATUS_DONE = '精算済み';
const ACCOUNT_PAYDATE_PLACEHOLDER = 'yyyyMMdd';
const ACCOUNT_FILENAME_MAX = 40;

function api_listExpenses() {
  const out = { ok:false, message:'', expenses:[] };

  try {
    const ss = openSS_(CONFIG.TASK_SS_ID);
    const sh = ss.getSheetByName(CONFIG.ACCOUNT_SHEET_NAME);
    if (!sh) return { ok:false, message:'シートが見つかりません：' + CONFIG.ACCOUNT_SHEET_NAME, expenses:[] };

    const lastRow = sh.getLastRow();
    if (lastRow < ACCOUNT_DATA_START_ROW) return { ok:true, expenses:[] };

    // ★B列（経費内容）で「本当の最終行」を決める（装飾だけの行を除外）
    const bVals = sh.getRange(
      ACCOUNT_DATA_START_ROW,
      ACCOUNT_COL.TITLE,
      lastRow - (ACCOUNT_DATA_START_ROW - 1),
      1
    ).getValues(); // B2:B
    let lastDataRow = ACCOUNT_DATA_START_ROW - 1; // 実行時は 2行目〜なので「シート上の行番号」を作る
    for (let i = bVals.length - 1; i >= 0; i--) {
      if (String(bVals[i][0] ?? '').trim() !== '') {
        lastDataRow = ACCOUNT_DATA_START_ROW + i; // シート行番号
        break;
      }
    }
    if (lastDataRow < ACCOUNT_DATA_START_ROW) return { ok:true, expenses:[] };

    // ★A:I（9列）を読む：A=SEQ, B=経費内容, C=説明, D=支払者, E=支払日, F=領収書, G=ステータス, H=精算依頼日, I=精算日
    const numRows = lastDataRow - (ACCOUNT_DATA_START_ROW - 1); // 2行目〜lastDataRow
    const values = sh.getRange(ACCOUNT_DATA_START_ROW, ACCOUNT_COL.SEQ, numRows, ACCOUNT_COL_COUNT).getValues(); // A2:I

    const expenses = [];
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      const sheetRow = i + ACCOUNT_DATA_START_ROW; // ★実シート行番号（絶対にズレない）

      const title = String(r[ACCOUNT_COL.TITLE - 1] ?? '').trim(); // B
      if (!title) continue; // 空行は無視（ただし rowNumber は sheetRow で保持してるのでズレない）

      expenses.push({
        seq: (r[ACCOUNT_COL.SEQ - 1] !== '' && r[ACCOUNT_COL.SEQ - 1] != null) ? Number(r[ACCOUNT_COL.SEQ - 1]) : null, // A
        rowNumber: sheetRow,                 // ★実行用：実シート行番号
        title: title,                        // B
        desc:  String(r[ACCOUNT_COL.DESC - 1] ?? ''),           // C
        payer: String(r[ACCOUNT_COL.PAYER - 1] ?? ''),          // D
        payDate: toIso_(r[ACCOUNT_COL.PAY_DATE - 1]),           // E
        receiptUrl: String(r[ACCOUNT_COL.RECEIPT - 1] ?? ''),   // F
        status: String(r[ACCOUNT_COL.STATUS - 1] ?? ''),        // G
        requestDate: toIso_(r[ACCOUNT_COL.REQUEST_DATE - 1]),   // H
        settleDate: toIso_(r[ACCOUNT_COL.SETTLE_DATE - 1])      // I
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

    const payDate = String(payload?.payDate||ACCOUNT_PAYDATE_PLACEHOLDER).trim(); // YYYY-MM-DD
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
      .replace(/[\\\/\:\*\?"\<\>\|]/g,'')
      .replace(/\s+/g,' ')
      .trim()
      .slice(0, ACCOUNT_FILENAME_MAX); // 長すぎ防止

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
    const status = String(data.status||ACCOUNT_STATUS_DEFAULT);

    let requestDate = data.requestDate ? new Date(String(data.requestDate)+'T00:00:00') : null;
    let settleDate  = data.settleDate ? new Date(String(data.settleDate)+'T00:00:00') : null;

    // ステータスに応じた日付自動補完（好みで）
    const today = new Date();
    if(status === ACCOUNT_STATUS_REQUESTED && !requestDate) requestDate = today;
    if(status === ACCOUNT_STATUS_DONE && !settleDate) settleDate = today;

    const ss = SpreadsheetApp.openById(CONFIG.TASK_SS_ID);
    const sh = ss.getSheetByName(CONFIG.ACCOUNT_SHEET_NAME);
    if(!sh) return { ok:false, message:'シートが見つかりません：' + CONFIG.ACCOUNT_SHEET_NAME };

    const rowValues = [[ title, desc, payer, payDate, receiptUrl, status, requestDate, settleDate ]]; // B:I

    if(mode === 'edit'){
      const rowNumber = Number(payload?.rowNumber);
      if(!Number.isFinite(rowNumber) || rowNumber < ACCOUNT_DATA_START_ROW) return { ok:false, message:'修正対象行が不正です。' };
      sh.getRange(rowNumber, ACCOUNT_COL.TITLE, 1, ACCOUNT_DATA_COLS_NO_SEQ).setValues(rowValues);

      out.ok = true;
      out.rowNumber = rowNumber;
      out.message = '更新しました。';
      return out;
    }

    // create: B列（経費内容）基準で末尾追加
    const lastDataRow = findLastDataRowByCol_(sh, ACCOUNT_COL.TITLE, ACCOUNT_DATA_START_ROW); // col=2(B), startRow=2
    const appendRow = Math.max(ACCOUNT_DATA_START_ROW, lastDataRow + 1);

    sh.getRange(appendRow, ACCOUNT_COL.TITLE, 1, ACCOUNT_DATA_COLS_NO_SEQ).setValues(rowValues);

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
      if(!Number.isFinite(row) || row < ACCOUNT_DATA_START_ROW) continue;

      // G=ステータス(7列目/B起点だと6番目) → ここはRangeで直接指定が安全
      // B起点のため、G列は「2+5=7列目」ではなく、シートの列でGは7。
      sh.getRange(row, ACCOUNT_COL.STATUS).setValue(newStatus); // G

      if(newStatus === ACCOUNT_STATUS_REQUESTED){
        sh.getRange(row, ACCOUNT_COL.REQUEST_DATE).setValue(today); // H 精算依頼日
      }
      if(newStatus === ACCOUNT_STATUS_DONE){
        sh.getRange(row, ACCOUNT_COL.SETTLE_DATE).setValue(today); // I 精算日
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
