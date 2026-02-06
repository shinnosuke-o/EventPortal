/***********************
 * FORM CREATE
 ***********************/
const FORM_TITLE_SUFFIX = '出欠確認';
const FORM_FILE_SUFFIX = '出欠確認フォーム';
const FORM_CHOICE_PLACEHOLDER = '選択肢を入力してください';
const FORM_RESPONSE_SHEET_INDEX = 6;
const FORM_FIND_SHEET_MAX_TRY = 15;
const FORM_FIND_SHEET_WAIT_MS = 400;

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
    ? `${fromStr}_${eventName}${FORM_FILE_SUFFIX}`
    : `${fromStr}-${toStr}_${eventName}${FORM_FILE_SUFFIX}`;
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
      item.setChoiceValues(choices.length ? choices : [FORM_CHOICE_PLACEHOLDER]);
      break;
    case 'チェックボックス':
      item = form.addCheckboxItem().setTitle(title);
      item.setChoiceValues(choices.length ? choices : [FORM_CHOICE_PLACEHOLDER]);
      break;
    case 'プルダウン':
      item = form.addListItem().setTitle(title);
      item.setChoiceValues(choices.length ? choices : [FORM_CHOICE_PLACEHOLDER]);
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
    const formTitle = `${eventName}${FORM_TITLE_SUFFIX}`;
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

    const responseSheet = findNewSheetAfterDestination_(ss, beforeSheetIds, FORM_FIND_SHEET_MAX_TRY, FORM_FIND_SHEET_WAIT_MS);
    if (!responseSheet) throw new Error('回答連携シートを特定できませんでした。');

    responseSheet.setName(eventName);
    moveSheetToIndex_(ss, responseSheet, FORM_RESPONSE_SHEET_INDEX);

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

