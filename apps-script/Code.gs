/**
 * 捐血叫號系統 - Google Apps Script 後端
 *
 * 試算表配置（綁定此 Script 的「叫號系統數據」試算表）：
 *   工作表1!A1  目前號碼（工作人員可直接在試算表修改，也可用 admin.html）
 *   工作表1!B1  狀態／公告文字（選填，空白則網頁不顯示）
 *   叫號紀錄    每次號碼變動自動新增一列：時間、號碼、來源（第一次使用時自動建立）
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   ADMIN_PIN                  叫號頁的密碼，建議 6 碼以上
 *   LINE_CHANNEL_ACCESS_TOKEN  （選用）LINE 查詢功能的 Channel access token
 *   LINE_WEBHOOK_KEY           （選用）LINE Webhook 網址上的秘密參數 ?key=...
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
const SOON_AHEAD = 3;             // 與叫號頁相同：前面剩幾位時提醒「快輪到您了」
const PAGE_URL = 'https://stntkero.github.io/call-number-system/';

const PIN_FAIL_KEY = 'pinFail';
const PIN_FAIL_LIMIT = 20;        // 10 分鐘內密碼錯誤超過此次數就暫停登入
const PIN_FAIL_WINDOW_SECONDS = 10 * 60;

/** 捐血者頁面：讀取目前狀態 */
function doGet() {
  return ContentService.createTextOutput(getStateJson_()).setMimeType(ContentService.MimeType.JSON);
}

/** 工作人員叫號頁（修改號碼或狀態，需要密碼），以及 LINE 官方帳號的 Webhook */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = null;
  }

  if (body && Array.isArray(body.events)) {
    try {
      handleLineWebhook_(body, e.parameter || {});
    } catch (err) {
      console.error('LINE Webhook 處理失敗', err);
    }
    // LINE 只看 HTTP 狀態碼；用 HtmlService 回應，避免 ContentService 的轉址
    return HtmlService.createHtmlOutput('OK');
  }

  let result;
  try {
    result = body ? handleAdmin_(body) : { ok: false, error: '請求格式錯誤' };
  } catch (err) {
    result = { ok: false, error: '請求格式錯誤' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/** 有快取的目前狀態，人多時也不會每次都讀試算表 */
function getStateJson_() {
  const cache = CacheService.getScriptCache();
  let json = cache.get(CACHE_KEY);
  if (!json) {
    json = JSON.stringify(readState_());
    cache.put(CACHE_KEY, json, CACHE_SECONDS);
  }
  return json;
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

/**
 * LINE 官方帳號：捐血者傳「查詢」或自己的號碼，就回覆目前叫號。
 * 只用回覆訊息（Reply），不計入官方帳號每月的訊息額度；其他訊息不回應，不影響真人聊天。
 */
function handleLineWebhook_(body, params) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const key = props.getProperty('LINE_WEBHOOK_KEY');
  // Apps Script 讀不到 X-Line-Signature 標頭，改用網址上的秘密參數擋掉偽造的請求
  if (!token || !key) {
    console.warn('LINE 查詢未設定：缺少 LINE_CHANNEL_ACCESS_TOKEN 或 LINE_WEBHOOK_KEY');
    return;
  }
  if (params.key !== key) return;

  body.events.forEach(event => {
    if (event.type !== 'message' || !event.message || event.message.type !== 'text' || !event.replyToken) return;
    if (!event.source || event.source.type !== 'user') return; // 只回應一對一聊天
    const query = parseLineQuery_(event.message.text);
    if (!query) return;
    replyLine_(token, event.replyToken, describeQueue_(JSON.parse(getStateJson_()), query.number));
  });
}

/** 看得懂的訊息：「查詢」「叫號」等關鍵字，或只有號碼（例如「230」「查詢 230」「230號」） */
function parseLineQuery_(text) {
  const normalized = String(text)
    .replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFF10 + 48)) // 全形數字轉半形
    .replace(/[\s:：,，、。!！?？~～]+/g, ''); // 去掉空白和標點（不含「.」，避免 2.30 變成 230）
  if (/^(查詢|查询|叫號|叫号|目前叫號|目前叫号|現在幾號|现在几号)(?:了|呢)?$/.test(normalized)) return { number: null };
  const match = normalized.match(/^(?:查詢|查询|我的號碼|我的号码|號碼|号码)?[#＃]?(\d{1,4})[號号]?$/);
  if (match && Number(match[1]) > 0) return { number: Number(match[1]) };
  return null;
}

function describeQueue_(state, myNumber) {
  const lines = [];
  if (state.status) lines.push('【公告】' + state.status);

  const current = Number(state.number);
  if (state.number === '' || state.number === null || !Number.isFinite(current)) {
    lines.push('目前尚未開始叫號。');
  } else {
    lines.push('目前叫號：' + current + ' 號');
    if (myNumber !== null) {
      // 用中性的說法：數字也可能是回答工作人員問題（例如「幾位」）而不是號碼牌
      const remaining = myNumber - current; // 還要再叫幾號才輪到這個號碼
      const ahead = remaining - 1;           // 排在前面、還沒被叫到的人數
      lines.push('查詢號碼：' + myNumber + ' 號');
      if (remaining < 0) lines.push('此號碼已過號，若是您的號碼，請洽現場工作人員。');
      else if (remaining === 0) lines.push('輪到 ' + myNumber + ' 號了！請依現場指示前往。');
      else if (ahead === 0) lines.push('下一位就是 ' + myNumber + ' 號！請回到現場等候。');
      else if (ahead <= SOON_AHEAD) lines.push('前面還有 ' + ahead + ' 位，快輪到了！請回到現場等候。');
      else lines.push('前面還有 ' + ahead + ' 位。');
      if (remaining > 0 && state.secPerNumber) {
        lines.push('預估約 ' + formatWait_(remaining * state.secPerNumber) + '（僅供參考）');
      }
    }
  }

  lines.push('');
  lines.push('即時叫號與聲音提醒：' + PAGE_URL);
  if (myNumber === null) lines.push('傳送您的號碼（例如：230），可查詢前面還有幾位。');
  return lines.join('\n');
}

function formatWait_(seconds) {
  const minutes = Math.max(1, Math.round(seconds / 60));
  if (minutes < 60) return minutes + ' 分鐘';
  return Math.floor(minutes / 60) + ' 小時' + (minutes % 60 ? ' ' + (minutes % 60) + ' 分鐘' : '');
}

function replyLine_(token, replyToken, text) {
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    console.error('LINE 回覆失敗', response.getResponseCode(), response.getContentText());
  }
}
