/**
 * 捐血叫號系統 - Google Apps Script 後端
 *
 * 試算表配置（綁定此 Script 的「叫號系統數據」試算表）：
 *   工作表1!A1  目前號碼（工作人員可直接在試算表修改，也可用 admin.html）
 *   工作表1!B1  狀態／公告文字（選填，空白則網頁不顯示）
 *   叫號紀錄    每次號碼變動自動新增一列：時間、號碼、來源（第一次使用時自動建立）
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   ADMIN_PIN   叫號頁的密碼，建議 6 碼以上
 */

const MAIN_SHEET = '工作表1';
const LOG_SHEET = '叫號紀錄';
const NUMBER_CELL = 'A1';
const STATUS_CELL = 'B1';

const CACHE_KEY = 'state';
const CACHE_SECONDS = 5;          // 號碼最多延遲幾秒反映到網頁
const LOG_ROWS_FOR_ESTIMATE = 20; // 用最近幾筆紀錄估算叫號速度
const MAX_GAP_SECONDS = 20 * 60;  // 超過 20 分鐘的間隔視為休息，不納入估算
const MAX_NUMBER = 9999;
const MAX_STATUS_LENGTH = 100;

const PIN_FAIL_KEY = 'pinFail';
const PIN_FAIL_LIMIT = 20;        // 10 分鐘內密碼錯誤超過此次數就暫停登入
const PIN_FAIL_WINDOW_SECONDS = 10 * 60;

/** 捐血者頁面：讀取目前狀態（有快取，人多時也不會每次都讀試算表） */
function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get(CACHE_KEY);
  if (!json) {
    json = JSON.stringify(readState_());
    cache.put(CACHE_KEY, json, CACHE_SECONDS);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/** 工作人員叫號頁：修改號碼或狀態，需要密碼 */
function doPost(e) {
  let result;
  try {
    result = handleAdmin_(JSON.parse(e.postData.contents));
  } catch (err) {
    result = { ok: false, error: '請求格式錯誤' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/** 工作人員直接在試算表改 A1 時，自動寫入叫號紀錄（簡易觸發條件，不需另外設定） */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== MAIN_SHEET) return;
  if (range.getRow() !== 1 || range.getColumn() > 2) return;

  if (range.getColumn() === 1) {
    const value = sheet.getRange(NUMBER_CELL).getValue();
    if (value !== '' && Number.isInteger(Number(value))) appendLog_(e.source, Number(value), '試算表');
  }
  try {
    CacheService.getScriptCache().remove(CACHE_KEY);
  } catch (err) {
    // 快取最多 CACHE_SECONDS 秒後也會自然過期
  }
}

function handleAdmin_(req) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get(PIN_FAIL_KEY) || 0);
  if (fails >= PIN_FAIL_LIMIT) {
    return { ok: false, error: '密碼錯誤次數過多，請 10 分鐘後再試' };
  }

  const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin) return { ok: false, error: '尚未設定 ADMIN_PIN，請參考 README' };
  if (String(req.pin || '') !== pin) {
    cache.put(PIN_FAIL_KEY, String(fails + 1), PIN_FAIL_WINDOW_SECONDS);
    return { ok: false, error: '密碼錯誤' };
  }

  // 多位工作人員同時操作時排隊處理，避免號碼互相覆蓋
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: '系統忙碌中，請再按一次' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getMainSheet_(ss);
    const numberCell = sheet.getRange(NUMBER_CELL);
    const current = Number(numberCell.getValue()) || 0;
    let next = null;

    switch (req.action) {
      case 'login':
        break;
      case 'next':
        next = current + 1;
        break;
      case 'prev':
        next = current - 1;
        break;
      case 'set':
        next = Number(req.value);
        break;
      case 'status':
        sheet.getRange(STATUS_CELL).setValue(String(req.value || '').trim().slice(0, MAX_STATUS_LENGTH));
        break;
      default:
        return { ok: false, error: '未知的操作' };
    }

    if (next !== null) {
      if (!Number.isInteger(next) || next < 0 || next > MAX_NUMBER) {
        return { ok: false, error: '號碼需為 0～' + MAX_NUMBER + ' 的整數' };
      }
      numberCell.setValue(next);
      appendLog_(ss, next, '叫號頁');
    }

    SpreadsheetApp.flush();
    const state = readState_();
    cache.put(CACHE_KEY, JSON.stringify(state), CACHE_SECONDS);
    return { ok: true, state: state };
  } finally {
    lock.releaseLock();
  }
}

function readState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const [number, status] = getMainSheet_(ss).getRange(NUMBER_CELL + ':' + STATUS_CELL).getValues()[0];
  const log = readRecentLog_(ss);
  const last = log[log.length - 1];
  return {
    number: number,
    status: String(status || ''),
    updatedAt: last ? last[0].getTime() : null,
    secPerNumber: estimateSecPerNumber_(log, ss.getSpreadsheetTimeZone()),
  };
}

/** 依今天最近幾次叫號的間隔，估算平均每一號要幾秒；資料不足時回傳 null */
function estimateSecPerNumber_(log, timeZone) {
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const rows = log.filter(row => Utilities.formatDate(row[0], timeZone, 'yyyy-MM-dd') === today);
  let totalSeconds = 0;
  let totalNumbers = 0;
  for (let i = 1; i < rows.length; i++) {
    const seconds = (rows[i][0] - rows[i - 1][0]) / 1000;
    const numbers = rows[i][1] - rows[i - 1][1];
    if (numbers > 0 && numbers <= 5 && seconds <= MAX_GAP_SECONDS) {
      totalSeconds += seconds;
      totalNumbers += numbers;
    }
  }
  return totalNumbers >= 3 ? Math.round(totalSeconds / totalNumbers) : null;
}

function readRecentLog_(ss) {
  const sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const start = Math.max(2, sheet.getLastRow() - LOG_ROWS_FOR_ESTIMATE + 1);
  return sheet.getRange(start, 1, sheet.getLastRow() - start + 1, 2).getValues()
    .filter(row => row[0] instanceof Date && typeof row[1] === 'number');
}

function appendLog_(ss, number, source) {
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(['時間', '號碼', '來源']);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  sheet.appendRow([new Date(), number, source]);
}

function getMainSheet_(ss) {
  return ss.getSheetByName(MAIN_SHEET) || ss.getSheets()[0];
}
