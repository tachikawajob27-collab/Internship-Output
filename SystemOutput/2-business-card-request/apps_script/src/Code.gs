/**
 * =========================================================================
 * Business Card Request Notifier (Public-safe)
 * - Form submit: notify Slack once per row and mark as "済"
 * - Edit: if key fields are edited BEFORE notified, notify and mark as "済"
 *
 * Security:
 * - Slack Webhook URL is stored in Script Properties (SLACK_WEBHOOK_URL)
 * - Slack message does NOT include phone/email/address (avoid PII in channels)
 * =========================================================================
 */

/** =======================
 * Config
 * ======================= */
const CONFIG = {
  sheetName: 'フォームの回答 1',     // フォーム回答が入るシートタブ名（完全一致）
  headerRow: 1,
  notifiedHeader: '通知済み',       // 無ければ自動追加
  notifiedValue: '済',

  // 氏名を取得する候補ヘッダー（上から順に探索）
  keyHeaders: ['お名前（漢字フルネーム）', 'お名前'],

  // 通知判定に使う「編集トリガー対象」ヘッダー（複数列対応）
  // ※このいずれかが入力/更新された時点で、未通知ならSlack通知する
  triggerHeaders: [
    'お名前（漢字フルネーム）',
    '役職',
    '電話番号',
    'メールアドレス',
    '名刺送付先の住所',
  ],

  // Script Properties key
  slackWebhookPropKey: 'SLACK_WEBHOOK_URL',

  // エラーログ（公開用でも残すと運用説明がしやすい）
  errorLogSheetName: 'ERROR_LOG',
};

/** =======================
 * Triggers
 * ======================= */

/**
 * On form submit (installable trigger recommended)
 * - Notify Slack for the submitted row (if not notified yet)
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.sheetName) return;

    const row = e.range.getRow();
    if (row <= CONFIG.headerRow) return;

    processRowIfNeeded_(sheet, row);
  } catch (err) {
    logError_(err);
  }
}

/**
 * On edit (installable trigger recommended)
 * - If any triggerHeaders column is edited AND row is not notified, notify
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.sheetName) return;

    const row = e.range.getRow();
    if (row <= CONFIG.headerRow) return;

    const editedCol = e.range.getColumn();

    // ヘッダー整備 & マップ作成
    const { headerMap, lastCol } = ensureHeaders_(sheet);

    // 編集された列のヘッダー名を特定
    const editedHeader = getHeaderNameByCol_(headerMap, editedCol);
    if (!editedHeader) return;

    // トリガー対象列以外は無視
    if (!CONFIG.triggerHeaders.includes(editedHeader)) return;

    // 既に通知済みなら無視
    const notifiedCol = headerMap[CONFIG.notifiedHeader];
    const notified = notifiedCol
      ? String(sheet.getRange(row, notifiedCol).getValue() ?? '').trim()
      : '';
    if (notified === CONFIG.notifiedValue) return;

    // 値が追加/変更された場合のみ通知（空→空、同値は除外）
    const newVal = (e.value ?? '') === null ? '' : String(e.value ?? '').trim();
    const oldVal = (e.oldValue ?? '') === null ? '' : String(e.oldValue ?? '').trim();
    if (!newVal || newVal === oldVal) return;

    // 行として必要情報が揃ったかは processRowIfNeeded_ 側で判定
    processRowIfNeeded_(sheet, row, { headerMap, lastCol });
  } catch (err) {
    logError_(err);
  }
}

/** =======================
 * Core
 * ======================= */

/**
 * Notify Slack if:
 * - key name exists
 * - not already notified
 * Then mark "通知済み" as "済"
 */
function processRowIfNeeded_(sheet, row, cache) {
  const { headerMap, lastCol } = cache || ensureHeaders_(sheet);

  const keyCol = getKeyCol_(headerMap);
  const notifiedCol = headerMap[CONFIG.notifiedHeader];

  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const name = String(values[keyCol - 1] ?? '').trim();
  const notified = notifiedCol ? String(values[notifiedCol - 1] ?? '').trim() : '';

  // 氏名（キー）が無い場合は通知しない（入力途中を想定）
  if (!name) return;

  // 通知済みならスキップ
  if (notified === CONFIG.notifiedValue) return;

  const payload = buildSlackPayload_(sheet, row, name);

  if (postToSlack_(payload)) {
    if (notifiedCol) sheet.getRange(row, notifiedCol).setValue(CONFIG.notifiedValue);
  }
}

/** =======================
 * Header utilities
 * ======================= */

function ensureHeaders_(sheet) {
  const headerRange = sheet.getRange(CONFIG.headerRow, 1, 1, Math.max(1, sheet.getLastColumn()));
  let headers = headerRange.getValues()[0].map(v => String(v ?? '').trim());

  // "通知済み" 列が無ければ末尾に追加
  if (!headers.includes(CONFIG.notifiedHeader)) {
    sheet.getRange(CONFIG.headerRow, headers.length + 1).setValue(CONFIG.notifiedHeader);
    headers.push(CONFIG.notifiedHeader);
  }

  const headerMap = {};
  headers.forEach((h, i) => { if (h) headerMap[h] = i + 1; });

  return { headerMap, lastCol: headers.length };
}

function getKeyCol_(headerMap) {
  for (const name of (CONFIG.keyHeaders || [])) {
    if (name && headerMap[name]) return headerMap[name];
  }
  throw new Error('キー列（keyHeaders）が見つかりません。ヘッダー名の綴りを確認してください。');
}

function getHeaderNameByCol_(headerMap, col) {
  for (const h of Object.keys(headerMap)) {
    if (headerMap[h] === col) return h;
  }
  return '';
}

/** =======================
 * Slack
 * ======================= */

function buildSlackPayload_(sheet, row, primaryName) {
  const ss = sheet.getParent();
  const sheetUrl = ss.getUrl();
  const directUrl = `${sheetUrl}#gid=${sheet.getSheetId()}&range=${row}:${row}`;

  // PII（電話/メール/住所）はSlack本文に載せない方針（公開用）
  const messageText = `${primaryName}さんから新規の名刺作成依頼がありました。`;

  return {
    text: messageText,
    blocks: [
      {
        type: 'section',
        text: { type: 'mrkdwn', text: messageText },
      },
      {
        type: 'section',
        text: { type: 'mrkdwn', text: `<${directUrl}|スプレッドシートで該当行を開く>` },
      },
    ],
  };
}

function postToSlack_(payload) {
  const url = PropertiesService.getScriptProperties().getProperty(CONFIG.slackWebhookPropKey);
  if (!url) throw new Error('スクリプトプロパティ SLACK_WEBHOOK_URL が未設定です。');

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  return code >= 200 && code < 300;
}

/** =======================
 * Manual tests (optional)
 * ======================= */

function testNotifyActiveRow() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`シート「${CONFIG.sheetName}」が見つかりません。`);

  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) throw new Error('通知したい行のセルをクリックしてから実行してください。');

  let row = activeRange.getRow();
  if (row <= CONFIG.headerRow) row = sheet.getLastRow();
  if (row <= CONFIG.headerRow) throw new Error('データ行がありません。');

  processRowIfNeeded_(sheet, row);
}

function pingSlack() {
  const url = PropertiesService.getScriptProperties().getProperty(CONFIG.slackWebhookPropKey);
  if (!url) throw new Error('SLACK_WEBHOOK_URL が未設定です。');

  const payload = { text: '✅ GAS → Slack 疎通テストOK（Business Card Request Notifier）' };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  return code >= 200 && code < 300;
}

/** =======================
 * Logging
 * ======================= */

function logError_(err) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CONFIG.errorLogSheetName) || ss.insertSheet(CONFIG.errorLogSheetName);
    sh.appendRow([new Date(), String(err && err.stack ? err.stack : err)]);
  } catch (e) {
    // 最後の砦：ログすら書けない場合
    console.error('[ERROR]', err);
  }
}
