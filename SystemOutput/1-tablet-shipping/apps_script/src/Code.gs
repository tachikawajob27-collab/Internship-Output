/**
 * =========================================================================
 * Tablet Shipping Automation (Public-safe)
 * - Form submit: initialize status + checkbox column
 * - Checkbox ON: generate B2 Cloud CSV row + update status
 * - From/To edit: sync asset ledger (update existing row or append new row)
 *
 * Security:
 * - Slack webhook is NOT hard-coded. Use Script Properties: SLACK_WEBHOOK_URL
 * =========================================================================
 */

/** =======================
 *  Config
 *  ======================= */
const CONFIG = {
  SHEETS: {
    RESP:  'タブレット発送管理シート',
    B2:    'B2クラウド出力用CSV',
    ASSET: '社内用資産管理シート',
    ERROR: 'ERROR_LOG',
  },

  STATUS: {
    INITIAL: '未対応',
    CREATED: '送り状作成済',
  },

  // Header names (must match the first row exactly)
  HEADER: {
    timestamp: 'タイムスタンプ',
    qty:       '必要台数',
    baseAddr:  'お届け先住所',
    shipAddr:  '送付先住所',
    recipient: '送付先企業名/氏名',
    agent:     '貴社/貴方の代理店コードをご記入ください',
    phone:     'お届け先電話番号',
    zip:       'お届け先郵便番号',
    building:  'お届け先建物名（アパートマンション名）',
    status:    '発送ステータス',
    checkbox:  '発送対象',
    fromNum:   'タブレット番号 From',
    toNum:     'タブレット番号 To',
  },

  // Column index (1-based) for specific inputs (faster than header search)
  INPUT_COL: {
    AGENT_CODE: 6,   // F
    FROM_NUM:   11,  // K
    TO_NUM:     12,  // L
  },

  // Asset sheet columns (1-based)
  OUTPUT_COL: {
    ASSET_ID:   1,  // A: Asset No. (e.g., W00013201)
    STORE_NAME: 4,  // D
    STORE_ID:   5,  // E
    AGENT_NAME: 6,  // F
    AGENT_CODE: 7,  // G
  },

  // B2 Cloud CSV header (13 cols)
  B2_HEADERS: [
    'お届け先郵便番号',
    'お届け先住所',
    'お届け先建物名',
    'お届け先名',
    'お届け先敬称',
    'お届け先電話番号',
    'ご依頼主郵便番号',
    'ご依頼主住所',
    'ご依頼主名',
    'ご依頼主電話番号',
    '品名１',
    '個数１',
    '発送予定日',
  ],

  // Values used in B2 rows (public-safe defaults; change in your environment)
  B2_DEFAULTS: {
    SENDER_NAME: 'ロケットナウ',
    ITEM_NAME:   'タブレット',
    HONORIFIC:   '様',
  },
};

/** =======================
 *  Entry points (Triggers)
 *  ======================= */

/**
 * Trigger: On form submit
 * - Set initial status
 * - Ensure checkbox column exists + validation
 * - Initialize checkbox to false
 */
function onFormSubmit(e) {
  try {
    const sheet = getResponseSheet_(e);
    const row   = e && e.range ? e.range.getRow() : sheet.getLastRow();

    const colStatus = getColByHeader_(sheet, CONFIG.HEADER.status);
    const colCb     = ensureCheckboxColumn_(sheet, CONFIG.HEADER.checkbox);

    sheet.getRange(row, colStatus).setValue(CONFIG.STATUS.INITIAL);
    sheet.getRange(row, colCb).setValue(false);

    // Optional: Slack notify (public-safe; requires script property)
    // notifySlackNewRequest_(sheet, row);

  } catch (err) {
    logError_(err);
  }
}

/**
 * Trigger: On edit
 * - Checkbox ON -> create B2 output row + update status
 * - From/To edit -> sync asset ledger
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEETS.RESP) return;

    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    if (editedRow <= 1) return;

    const colStatus = getColByHeader_(sheet, CONFIG.HEADER.status);
    const colCb     = getColByHeader_(sheet, CONFIG.HEADER.checkbox);

    // Feature 1: B2 export when checkbox becomes TRUE
    if (editedCol === colCb) {
      const newVal = String(e.value || '').toUpperCase();
      const oldVal = String(e.oldValue || '').toUpperCase();
      if (newVal === 'TRUE' && oldVal !== 'TRUE') {
        createAndExportB2Row_(sheet, editedRow, colStatus);
      }
      return;
    }

    // Feature 2: Asset ledger sync when From/To is edited
    if (editedCol === CONFIG.INPUT_COL.FROM_NUM || editedCol === CONFIG.INPUT_COL.TO_NUM) {
      syncAssetsFromRange_(sheet, editedRow);
      return;
    }

  } catch (err) {
    logError_(err);
  }
}

/** =======================
 *  Feature 1: B2 Export
 *  ======================= */

function createAndExportB2Row_(srcSheet, row, statusColIndex) {
  // Skip if already created
  const currentStatus = String(srcSheet.getRange(row, statusColIndex).getDisplayValue() || '').trim();
  if (currentStatus === CONFIG.STATUS.CREATED) return;

  const b2Sheet = ensureB2SheetReady_();

  const c = {
    qty:      getColByHeader_(srcSheet, CONFIG.HEADER.qty),
    rcpt:     getColByHeader_(srcSheet, CONFIG.HEADER.recipient),
    baseAddr: getColByHeader_(srcSheet, CONFIG.HEADER.baseAddr),
    shipAddr: getColByHeader_(srcSheet, CONFIG.HEADER.shipAddr),
    phone:    getColByHeader_(srcSheet, CONFIG.HEADER.phone),
    zip:      getColByHeader_(srcSheet, CONFIG.HEADER.zip),
    building: getColByHeader_(srcSheet, CONFIG.HEADER.building),
  };

  const qty          = srcSheet.getRange(row, c.qty).getDisplayValue();
  const recipient    = srcSheet.getRange(row, c.rcpt).getDisplayValue();
  const baseAddr     = srcSheet.getRange(row, c.baseAddr).getDisplayValue();
  const shipAddr     = srcSheet.getRange(row, c.shipAddr).getDisplayValue();
  const phone        = srcSheet.getRange(row, c.phone).getDisplayValue();
  const zip          = srcSheet.getRange(row, c.zip).getDisplayValue();
  const buildingName = srcSheet.getRange(row, c.building).getDisplayValue();
  const address      = shipAddr || baseAddr || '';

  const shippingDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const rowData = [
    zip,
    address,
    buildingName,
    recipient,
    CONFIG.B2_DEFAULTS.HONORIFIC,
    phone,
    '', // sender zip (optional)
    '', // sender address (optional)
    CONFIG.B2_DEFAULTS.SENDER_NAME,
    '', // sender phone (optional)
    CONFIG.B2_DEFAULTS.ITEM_NAME,
    qty,
    shippingDate,
  ];

  // Duplicate prevention (key fields)
  if (isDuplicateB2_(b2Sheet, rowData)) {
    srcSheet.getRange(row, statusColIndex).setValue(CONFIG.STATUS.CREATED);
    return;
  }

  b2Sheet.appendRow(rowData);
  srcSheet.getRange(row, statusColIndex).setValue(CONFIG.STATUS.CREATED);
}

function ensureB2SheetReady_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEETS.B2) || ss.insertSheet(CONFIG.SHEETS.B2);

  // Ensure header
  const firstRow = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  const hasHeader = firstRow.some(v => String(v).trim() !== '');
  if (!hasHeader) {
    sh.getRange(1, 1, 1, CONFIG.B2_HEADERS.length).setValues([CONFIG.B2_HEADERS]);
  }

  // Format zip/phone as text to preserve leading zeros
  const maxRows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 1,  maxRows, 1).setNumberFormat('@');   // A: zip
  sh.getRange(2, 6,  maxRows, 1).setNumberFormat('@');   // F: phone
  sh.getRange(2, 7,  maxRows, 1).setNumberFormat('@');   // G: sender zip
  sh.getRange(2, 10, maxRows, 1).setNumberFormat('@');   // J: sender phone

  return sh;
}

function isDuplicateB2_(b2Sheet, rowData) {
  const lastRow = b2Sheet.getLastRow();
  if (lastRow < 2) return false;

  const values = b2Sheet.getRange(2, 1, lastRow - 1, rowData.length).getValues();

  // Compare selected fields: zip, address, building, recipient, item, qty
  return values.some(r =>
    String(r[0]  || '') === String(rowData[0]  || '') && // zip
    String(r[1]  || '') === String(rowData[1]  || '') && // address
    String(r[2]  || '') === String(rowData[2]  || '') && // building
    String(r[3]  || '') === String(rowData[3]  || '') && // recipient
    String(r[10] || '') === String(rowData[10] || '') && // item
    String(r[11] || '') === String(rowData[11] || '')    // qty
  );
}

/** =======================
 *  Feature 2: Asset Ledger Sync
 *  ======================= */

function syncAssetsFromRange_(inputSheet, inputRow) {
  try {
    const fromStr = String(inputSheet.getRange(inputRow, CONFIG.INPUT_COL.FROM_NUM).getDisplayValue() || '').trim();
    const toStr   = String(inputSheet.getRange(inputRow, CONFIG.INPUT_COL.TO_NUM).getDisplayValue()   || '').trim();
    if (!fromStr || !toStr) return;

    // Parse "W00013201"
    const m = fromStr.match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return;

    const prefix        = m[1];
    const fromNumeric   = parseInt(m[2], 10);
    const paddingLength = m[2].length;

    // Allow both "W00013210" or "00013210"
    const toNumeric = parseInt(String(toStr).replace(prefix, ''), 10);
    if (!Number.isFinite(fromNumeric) || !Number.isFinite(toNumeric) || fromNumeric > toNumeric) return;

    const agentCode = String(inputSheet.getRange(inputRow, CONFIG.INPUT_COL.AGENT_CODE).getDisplayValue() || '').trim() || '??';

    // Public-safe placeholders (replace with actual mapping in your environment)
    const storeName = '??';
    const storeId   = '??';
    const agentName = '??';

    const ss = inputSheet.getParent();
    const assetSheet = ss.getSheetByName(CONFIG.SHEETS.ASSET);
    if (!assetSheet) throw new Error(`出力先シート「${CONFIG.SHEETS.ASSET}」が見つかりません。`);

    const idToRow = buildAssetIdToRowMap_(assetSheet, CONFIG.OUTPUT_COL.ASSET_ID);

    const maxCol = Math.max(...Object.values(CONFIG.OUTPUT_COL));

    for (let n = fromNumeric; n <= toNumeric; n++) {
      const assetId = prefix + String(n).padStart(paddingLength, '0');

      const rowIndex = idToRow[assetId];
      if (rowIndex) {
        assetSheet.getRange(rowIndex, CONFIG.OUTPUT_COL.ASSET_ID).setValue(assetId);
        assetSheet.getRange(rowIndex, CONFIG.OUTPUT_COL.AGENT_CODE).setValue(agentCode);
        assetSheet.getRange(rowIndex, CONFIG.OUTPUT_COL.AGENT_NAME).setValue(agentName);
        assetSheet.getRange(rowIndex, CONFIG.OUTPUT_COL.STORE_NAME).setValue(storeName);
        assetSheet.getRange(rowIndex, CONFIG.OUTPUT_COL.STORE_ID).setValue(storeId);
      } else {
        const newRow = Array(maxCol).fill('');
        newRow[CONFIG.OUTPUT_COL.ASSET_ID   - 1] = assetId;
        newRow[CONFIG.OUTPUT_COL.AGENT_CODE - 1] = agentCode;
        newRow[CONFIG.OUTPUT_COL.AGENT_NAME - 1] = agentName;
        newRow[CONFIG.OUTPUT_COL.STORE_NAME - 1] = storeName;
        newRow[CONFIG.OUTPUT_COL.STORE_ID   - 1] = storeId;
        assetSheet.appendRow(newRow);

        // Update map to avoid duplicate appends in the same run
        idToRow[assetId] = assetSheet.getLastRow();
      }
    }
  } catch (err) {
    logError_(err);
  }
}

function buildAssetIdToRowMap_(assetSheet, assetIdCol) {
  const lastRow = assetSheet.getLastRow();
  const rows = Math.max(lastRow, 1);

  const ids = assetSheet.getRange(1, assetIdCol, rows, 1).getValues()
    .map(r => String(r[0] || '').trim());

  const map = {};
  ids.forEach((val, idx) => {
    if (val) map[val] = idx + 1; // 1-based row
  });
  return map;
}

/** =======================
 *  Optional: Slack Notification (public-safe)
 *  ======================= */

function notifySlackNewRequest_(sheet, row) {
  const url = getScriptProperty_('SLACK_WEBHOOK_URL');
  if (!url) return; // no-op if not set

  // Build a minimal message (avoid personal/sensitive data in public repo)
  const recipientCol = getColByHeader_(sheet, CONFIG.HEADER.recipient);
  const qtyCol       = getColByHeader_(sheet, CONFIG.HEADER.qty);
  const recipient    = sheet.getRange(row, recipientCol).getDisplayValue();
  const qty          = sheet.getRange(row, qtyCol).getDisplayValue();

  const payload = {
    text: `タブレット送付依頼を受信しました：宛先=${recipient} / 台数=${qty}`,
  };

  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function getScriptProperty_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/** =======================
 *  Utilities
 *  ======================= */

function getResponseSheet_(e) {
  if (e && e.range) return e.range.getSheet();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEETS.RESP);
  if (!sh) throw new Error(`シート「${CONFIG.SHEETS.RESP}」が見つかりません。`);
  return sh;
}

function getColByHeader_(sheet, headerText) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());
  const idx = headers.findIndex(h => h === String(headerText).trim());
  if (idx === -1) throw new Error(`ヘッダー「${headerText}」が見つかりませんでした。`);
  return idx + 1; // 1-based
}

function ensureCheckboxColumn_(sheet, checkboxHeader) {
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  let col = headers.findIndex(h => String(h).trim() === String(checkboxHeader).trim()) + 1;

  if (col === 0) {
    col = sheet.getLastColumn() + 1;
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, col).setValue(checkboxHeader);
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  const maxRows = sheet.getMaxRows();
  const rng = sheet.getRange(2, col, Math.max(maxRows - 1, 1), 1);
  rng.setDataValidation(rule);

  return col;
}

function logError_(err) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEETS.ERROR) || ss.insertSheet(CONFIG.SHEETS.ERROR);
  sh.appendRow([new Date(), String(err && err.message ? err.message : err)]);
}

/**
 * Helper: For existing data repair
 * - If checkbox is true and status is not created, generate B2 row
 */
function repairExistingRows_() {
  try {
    const sheet = getResponseSheet_();
    const colStatus = getColByHeader_(sheet, CONFIG.HEADER.status);
    const colCb     = getColByHeader_(sheet, CONFIG.HEADER.checkbox);

    const lastRow = sheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) {
      const cbVal = sheet.getRange(row, colCb).getValue();
      const stVal = String(sheet.getRange(row, colStatus).getDisplayValue() || '').trim();
      if (cbVal === true && stVal !== CONFIG.STATUS.CREATED) {
        createAndExportB2Row_(sheet, row, colStatus);
      }
    }
  } catch (err) {
    logError_(err);
  }
}
