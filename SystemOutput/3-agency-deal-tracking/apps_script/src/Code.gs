/**
 * =========================================================================
 * Agency Deal Sheets -> HQ Master Upsert Sync (Public-safe)
 *
 * What it does:
 * 1) Scan Google Sheets files under a target Drive folder (recursive).
 * 2) For each "deal management" sheet, read rows and upsert them into:
 *    - Master sheet (e.g., "Deal Master")
 *    - CP pending sheet (e.g., "CP Pending Agency Records")
 *    using composite key: AgencyID|StoreID
 *
 * Security:
 * - Drive folder ID is stored in Script Properties (AGENCY_ROOT_FOLDER_ID)
 * - No internal IDs/URLs are hard-coded in this public version.
 * =========================================================================
 */

/** =======================
 * Config
 * ======================= */
const CONFIG = {
  // Script Properties key names
  props: {
    rootFolderId: 'AGENCY_ROOT_FOLDER_ID', // <-- set this in Script Properties
  },

  // Master spreadsheet = the spreadsheet this script is bound to
  master: {
    masterSheetName: '案件マスター',          // public-safe name (change as needed)
    cpSheetName:     'CP未反映レコード',      // public-safe name (change as needed)
    headerRow: 1,
  },

  // Process only files whose name includes this keyword
  fileNameMustInclude: '案件管理シート',

  // Source (agency-side) header names
  sourceHeaders: {
    tempStoreId:     'Temp_Store_ID',
    agencyId:        'Agency_ID',
    storeName:       'Store_Name',
    storeId:         'Store_ID',
    tabletDeviceId:  'Tablet_Deviceid',
    tabletSn:        'Tablet_SN',
  },

  // Master/CP header names (destination)
  destHeaders: {
    storeId:         '店舗ID',
    agencyId:        '代理店ID',
    tempStoreId:     'temp店舗ID',
    storeName:       '店舗名',
    tabletDeviceId:  'タブレット端末番号',        // optional in CP sheet
    tabletSn:        'タブレットシリアルナンバー', // optional in CP sheet
  },

  // Safety guards
  maxRowsPerSheet: 2000,
  emptyStreakBreak: 10,

  // Logging
  errorLogSheetName: 'ERROR_LOG',
};

/** =======================
 * Main
 * ======================= */
function syncAgencySheetsToMaster() {
  try {
    const rootFolderId = getProp_(CONFIG.props.rootFolderId);
    if (!rootFolderId) throw new Error(`Script Properties に ${CONFIG.props.rootFolderId} が未設定です。`);

    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const masterSs   = SpreadsheetApp.getActiveSpreadsheet();

    const masterSheet = masterSs.getSheetByName(CONFIG.master.masterSheetName);
    const cpSheet     = masterSs.getSheetByName(CONFIG.master.cpSheetName);

    if (!masterSheet) throw new Error(`マスタシート「${CONFIG.master.masterSheetName}」が見つかりません。`);
    if (!cpSheet)     throw new Error(`CPシート「${CONFIG.master.cpSheetName}」が見つかりません。`);

    // --- Resolve destination header columns ---
    const masterHeaderRow = getHeaderRow_(masterSheet, CONFIG.master.headerRow);
    const masterCols = getHeaderCols_(masterHeaderRow, CONFIG.destHeaders, ['storeId', 'agencyId']);
    if (!masterCols) return;

    const cpHeaderRow = getHeaderRow_(cpSheet, CONFIG.master.headerRow);
    const cpCols = getHeaderCols_(cpHeaderRow, CONFIG.destHeaders, ['storeId', 'agencyId', 'tempStoreId', 'storeName']);
    if (!cpCols) return;

    // --- Build key -> row maps for upsert ---
    const keyToRowMaster = {};
    const lastDataRowMaster = buildKeyMap_(masterSheet, masterCols, keyToRowMaster);

    const keyToRowCp = {};
    const lastDataRowCp = buildKeyMap_(cpSheet, cpCols, keyToRowCp);

    let nextNewRowMaster = (lastDataRowMaster >= 2) ? lastDataRowMaster + 1 : 2;
    let nextNewRowCp     = (lastDataRowCp     >= 2) ? lastDataRowCp     + 1 : 2;

    // --- Collect spreadsheet IDs under folder (recursive) ---
    const fileIds = [];
    collectSpreadsheetIds_(rootFolder, fileIds);

    let insertedMaster = 0;
    let updatedMaster  = 0;
    let processedFiles = 0;

    for (const fileId of fileIds) {
      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName();

      if (fileName.indexOf(CONFIG.fileNameMustInclude) === -1) continue;

      processedFiles++;

      try {
        const ss = SpreadsheetApp.openById(fileId);
        const sheet = ss.getSheets()[0]; // first sheet only

        const srcLastRow = sheet.getLastRow();
        const srcLastCol = sheet.getLastColumn();
        if (srcLastRow < 3) continue;

        const srcValues = sheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();
        const srcHeaderRow = srcValues[0];

        // Resolve source columns (required)
        const srcCols = resolveSourceCols_(srcHeaderRow, CONFIG.sourceHeaders);
        if (!srcCols) {
          // Missing required header -> skip this file
          continue;
        }

        let emptyStreak = 0;
        let processedRows = 0;

        for (let r = 2; r < srcValues.length; r++) { // 3rd row = data start
          if (processedRows >= CONFIG.maxRowsPerSheet) break;
          processedRows++;

          const row = srcValues[r];

          const tempStoreId    = toStr_(row[srcCols.tempStoreId]);
          const agencyId       = toStr_(row[srcCols.agencyId]);
          const storeName      = toStr_(row[srcCols.storeName]);
          const storeId        = toStr_(row[srcCols.storeId]);
          const tabletDeviceId = toStr_(row[srcCols.tabletDeviceId]);
          const tabletSn       = toStr_(row[srcCols.tabletSn]);

          const hasAny = agencyId || storeId || storeName || tabletDeviceId || tabletSn || tempStoreId;

          if (!hasAny) {
            emptyStreak++;
            if (emptyStreak >= CONFIG.emptyStreakBreak) break;
            continue;
          } else {
            emptyStreak = 0;
          }

          // Validate key
          if (!agencyId || !storeId) continue;
          if (agencyId === '#REF!' || storeId === '#REF!') continue;

          const key = makeKey_(agencyId, storeId);

          // Decide master row (upsert)
          let masterRow;
          if (keyToRowMaster[key]) {
            masterRow = keyToRowMaster[key];
            updatedMaster++;
          } else {
            masterRow = nextNewRowMaster++;
            keyToRowMaster[key] = masterRow;
            insertedMaster++;
          }

          // Decide cp row (upsert)
          let cpRow;
          if (keyToRowCp[key]) {
            cpRow = keyToRowCp[key];
          } else {
            cpRow = nextNewRowCp++;
            keyToRowCp[key] = cpRow;
          }

          const data = { storeId, agencyId, tempStoreId, storeName, tabletDeviceId, tabletSn };

          // Write to master & cp
          writeRowToSheet_(masterSheet, masterCols, masterRow, data);
          writeRowToSheet_(cpSheet, cpCols, cpRow, data);
        }
      } catch (eFile) {
        logError_(eFile, `File processing error: ${fileName}`);
        continue;
      }
    }

    Logger.log('完了: 対象ファイル %d / マスタ 追加 %d 行 / 更新 %d 行',
               processedFiles, insertedMaster, updatedMaster);

  } catch (eAll) {
    logError_(eAll, 'syncAgencySheetsToMaster failed');
  }
}

/** =======================
 * Helpers: headers / maps
 * ======================= */

function getHeaderRow_(sheet, headerRowIndex) {
  return sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * headerRow: 1行目配列
 * headerDef: { key: '列名', ... }
 * requiredKeys: 必須キー配列
 */
function getHeaderCols_(headerRow, headerDef, requiredKeys) {
  const cols = {};
  for (const key in headerDef) {
    const name = headerDef[key];
    const idx = headerRow.indexOf(name);
    if (idx === -1) {
      if (requiredKeys && requiredKeys.indexOf(key) !== -1) {
        Logger.log('ヘッダー「%s」が見つからないため処理を中止します。', name);
        return null;
      }
      continue; // optional
    }
    cols[key] = idx + 1; // 1-based
  }
  return cols;
}

/**
 * Resolve source column indices (0-based) for required headers.
 * Return null if any required header is missing.
 */
function resolveSourceCols_(srcHeaderRow, sourceHeaders) {
  const srcCols = {};
  for (const key in sourceHeaders) {
    const name = sourceHeaders[key];
    const idx = srcHeaderRow.indexOf(name);
    if (idx === -1) {
      Logger.log('必須列「%s」が見つからないため、このファイルはスキップします。', name);
      return null;
    }
    srcCols[key] = idx; // 0-based
  }
  return srcCols;
}

/**
 * Build key -> row map for existing destination sheet.
 * Returns last non-empty data row index.
 */
function buildKeyMap_(sheet, cols, keyToRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  let lastDataRow = 1;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowNum = 2 + i;
      const row = values[i];

      const hasAny = row.some(v => toStr_(v) !== '');
      if (hasAny) lastDataRow = rowNum;

      const storeId  = cols.storeId  ? toStr_(row[cols.storeId  - 1]) : '';
      const agencyId = cols.agencyId ? toStr_(row[cols.agencyId - 1]) : '';
      if (!storeId || !agencyId) continue;

      keyToRow[makeKey_(agencyId, storeId)] = rowNum;
    }
  }
  return lastDataRow;
}

/**
 * Write one row to destination sheet.
 * - Only writes non-empty values (prevents overwriting with blanks).
 */
function writeRowToSheet_(sheet, cols, rowIndex, data) {
  writeIfNotEmpty_(sheet, rowIndex, cols.storeId,        data.storeId);
  writeIfNotEmpty_(sheet, rowIndex, cols.agencyId,       data.agencyId);
  writeIfNotEmpty_(sheet, rowIndex, cols.tempStoreId,    data.tempStoreId);
  writeIfNotEmpty_(sheet, rowIndex, cols.storeName,      data.storeName);
  writeIfNotEmpty_(sheet, rowIndex, cols.tabletDeviceId, data.tabletDeviceId);
  writeIfNotEmpty_(sheet, rowIndex, cols.tabletSn,       data.tabletSn);
}

function writeIfNotEmpty_(sheet, rowIndex, colIndex, value) {
  if (!colIndex) return; // destination may not have this column
  const v = toStr_(value);
  if (v === '') return;
  sheet.getRange(rowIndex, colIndex).setValue(v);
}

/** =======================
 * Helpers: Drive / keys / props / logging
 * ======================= */

function collectSpreadsheetIds_(folder, list) {
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) list.push(files.next().getId());

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) collectSpreadsheetIds_(subFolders.next(), list);
}

function makeKey_(agencyId, storeId) {
  return `${toStr_(agencyId)}|${toStr_(storeId)}`;
}

function toStr_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function logError_(err, context) {
  const msg = `[ERROR] ${context || ''}\n${err && err.stack ? err.stack : err}`;
  console.error(msg);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.errorLogSheetName) || ss.insertSheet(CONFIG.errorLogSheetName);
    sh.appendRow([new Date(), context || '', String(err && err.message ? err.message : err)]);
  } catch (_) {
    // ignore
  }
}

/** =======================
 * Optional: test root folder property
 * ======================= */
function testRootFolder_() {
  const id = getProp_(CONFIG.props.rootFolderId);
  if (!id) throw new Error(`Script Properties に ${CONFIG.props.rootFolderId} が未設定です。`);
  const folder = DriveApp.getFolderById(id);
  Logger.log('フォルダ名: %s', folder.getName());
}
