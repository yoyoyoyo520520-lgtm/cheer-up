/**
 * 「融的小小關心所」後台 v4.0
 * 安裝位置：目標 Google 試算表 → 擴充功能 → Apps Script
 * 更新本檔後：部署 → 管理部署作業 → 鉛筆 → 版本「新版本」→ 部署。
 *
 * 首次設定：在原試算表的 Apps Script 編輯器執行 setupBackend()。
 * POST：寫入狀態與留言；使用 requestId 避免重複。
 * GET：僅回傳收件確認或安全的服務狀態，不回傳私人資料。
 */
var BACKEND_VERSION = '4.0';
var SHEET_NAME = '回覆紀錄';
var PLACES_SHEET_NAME = '想去的地方';
// 新心情與舊工作狀態並存，更新前端期間仍接收舊版回覆。
var ALLOWED_CHOICES = ['心情美麗', '平平淡淡', '有點累了', '心情不好', '有點煩惱', '想找你聊聊', '要去上班了', '還在上班', '準備下班', '下班了', '午休吃飯', '準備開會'];
var SAFE_ERROR_CODES = ['UNSUPPORTED_CHOICE', 'INVALID_REQUEST', 'SPREADSHEET_NOT_CONFIGURED', 'WRITE_FAILED', 'INVALID_PLACE', 'PLACE_WRITE_FAILED'];

// 只從原試算表編輯器手動執行；識別資訊留在 Script Properties。
function setupBackend() {
  try {
    var book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) throw safeError_('SPREADSHEET_NOT_CONFIGURED');
    PropertiesService.getScriptProperties().setProperty('CHEER_SPREADSHEET_ID', book.getId());
    return {ok: true, ready: true, version: BACKEND_VERSION};
  } catch (err) {
    var code = safeCode_(err);
    console.error('setupBackend failed: ' + code);
    return {ok: false, ready: false, version: BACKEND_VERSION, code: code};
  }
}

function getBook_() {
  var id = PropertiesService.getScriptProperties().getProperty('CHEER_SPREADSHEET_ID');
  var book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw safeError_('SPREADSHEET_NOT_CONFIGURED');
  return book;
}

function safeError_(code) {
  var err = new Error(code);
  err.code = code;
  return err;
}

function safeCode_(err, fallback) {
  return err && SAFE_ERROR_CODES.indexOf(err.code) !== -1 ? err.code : (fallback || 'WRITE_FAILED');
}

function validRequestId_(requestId) {
  return typeof requestId === 'string' && /^[a-zA-Z0-9_-]{12,90}$/.test(requestId);
}

function statusError_(code) {
  // error 保持舊版前端可識別；兩欄都只含固定代碼。
  return {ok: false, received: false, code: code, error: code};
}

function doPost(e) {
  var requestId = '';
  var isPlace = false;
  try {
    var data;
    try {
      data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (_) {
      throw safeError_('INVALID_REQUEST');
    }
    if (!data || typeof data !== 'object' || Array.isArray(data) || !validRequestId_(data.requestId)) {
      throw safeError_('INVALID_REQUEST');
    }
    requestId = data.requestId;
    isPlace = data.action === 'place';
    if (data.action != null && data.action !== '' && !isPlace) throw safeError_('INVALID_REQUEST');
    if (isPlace) {
      savePlace_(data.place, requestId);
      return ContentService.createTextOutput('OK');
    }
    var choice = typeof data.choice === 'string' ? data.choice : '';
    if (ALLOWED_CHOICES.indexOf(choice) === -1) throw safeError_('UNSUPPORTED_CHOICE');
    var note = data.note == null ? '' : data.note;
    if (typeof note !== 'string') throw safeError_('INVALID_REQUEST');
    note = note.trim();
    if (Array.from(note).length > 120 || note.length > 600) throw safeError_('INVALID_REQUEST');
    // 不接受試算表公式與控制字元，以免被解讀為公式。
    note = note.replace(/[\u0000-\u001f\u007f]/g, ' ');
    if (/^[=+\-@]/.test(note)) note = "'" + note;

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sheet = getSheet_();
      if (!findRequest_(sheet, requestId)) {
        sheet.appendRow([new Date(), choice, note, requestId]);
        SpreadsheetApp.flush();
      }
      var cache = CacheService.getScriptCache();
      cache.put('cheer_' + requestId, '1', 21600);
      cache.remove('cheer_error_' + requestId);
    } finally {
      lock.releaseLock();
    }
    return ContentService.createTextOutput('OK');
  } catch (err) {
    var code = safeCode_(err, isPlace ? 'PLACE_WRITE_FAILED' : 'WRITE_FAILED');
    if (validRequestId_(requestId)) {
      try {
        CacheService.getScriptCache().put((isPlace ? 'cheer_place_error_' : 'cheer_error_') + requestId, code, 21600);
      } catch (_) {
        // 快取故障不改變固定的公開回應，也不記錄使用者資料。
      }
    }
    console.error('doPost failed: ' + code);
    return ContentService.createTextOutput('ERROR');
  }
}

function doGet(e) {
  var parameter = (e && e.parameter) || {};
  var callback = String(parameter.callback || '');
  if (!/^__cheerAck_[0-9]+_[a-z0-9]+$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }
  var result;
  try {
    if (parameter.action === 'health') {
      result = {ok: true, version: BACKEND_VERSION, allowedChoices: ALLOWED_CHOICES.slice(), ready: true, features: {places: true}};
      try {
        getBook_();
      } catch (err) {
        result.ok = false;
        result.ready = false;
        result.code = safeCode_(err);
      }
    } else if (parameter.action === 'places') {
      result = {ok: true, places: listPlaces_(), version: BACKEND_VERSION};
    } else {
      var requestId = parameter.id;
      var isPlace = parameter.action === 'place_status';
      if ((parameter.action !== 'status' && !isPlace) || !validRequestId_(requestId)) {
        result = statusError_('INVALID_REQUEST');
      } else {
        var cache = CacheService.getScriptCache();
        var received = cache.get((isPlace ? 'cheer_place_' : 'cheer_') + requestId) === '1';
        var lookupCode = '';
        if (!received) {
          try {
            var sheet = getBook_().getSheetByName(isPlace ? PLACES_SHEET_NAME : SHEET_NAME);
            received = !!(sheet && (isPlace ? findPlaceRequest_(sheet, requestId) : findRequest_(sheet, requestId)));
          } catch (err) {
            lookupCode = safeCode_(err, isPlace ? 'PLACE_WRITE_FAILED' : 'WRITE_FAILED');
          }
        }
        if (received) {
          result = {ok: true, received: true, serverTime: new Date().toISOString()};
        } else {
          var cachedCode = cache.get((isPlace ? 'cheer_place_error_' : 'cheer_error_') + requestId);
          var code = SAFE_ERROR_CODES.indexOf(cachedCode) !== -1 ? cachedCode : lookupCode;
          result = code ? statusError_(code) : {ok: true, received: false};
        }
      }
    }
  } catch (err) {
    var code = safeCode_(err, parameter.action === 'places' || parameter.action === 'place_status' ? 'PLACE_WRITE_FAILED' : 'WRITE_FAILED');
    console.error('doGet failed: ' + code);
    result = statusError_(code);
  }
  // JavaScript JSONP: 讓放在 GitHub Pages 的網站能讀取「後台已確認」結果。
  // 回傳內容只有收件成功與否，不公開試算表或留言。
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSheet_() {
  var book = getBook_();
  var sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = book.insertSheet(SHEET_NAME);
  // 舊版原本只有前 2 欄，此處可直接相容升級，不清除既有資料。
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 4).getValue() !== '請求 ID') {
    sheet.getRange(1, 1, 1, 4).setValues([['送出時間', '選擇狀態', '想說的話', '請求 ID']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRequest_(sheet, requestId) {
  if (sheet.getLastRow() < 2) return false;
  var requestRange = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1);
  return !!requestRange.createTextFinder(requestId).matchEntireCell(true).findNext();
}

function cleanPlace_(place) {
  if (typeof place !== 'string') throw safeError_('INVALID_PLACE');
  place = place.replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  var length = Array.from(place).length;
  if (length < 1 || length > 60) throw safeError_('INVALID_PLACE');
  return place;
}

function savePlace_(place, requestId) {
  place = cleanPlace_(place);
  if (/^[=+\-@]/.test(place)) place = "'" + place;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var book = getBook_();
    var sheet = book.getSheetByName(PLACES_SHEET_NAME);
    if (!sheet) sheet = book.insertSheet(PLACES_SHEET_NAME);
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 3).getValue() !== '請求 ID') {
      sheet.getRange(1, 1, 1, 3).setValues([['提出時間', '想去的地方', '請求 ID']]);
      sheet.setFrozenRows(1);
    }
    if (!findPlaceRequest_(sheet, requestId)) {
      sheet.appendRow([new Date(), place, requestId]);
      SpreadsheetApp.flush();
    }
    var cache = CacheService.getScriptCache();
    cache.put('cheer_place_' + requestId, '1', 21600);
    cache.remove('cheer_place_error_' + requestId);
  } finally {
    lock.releaseLock();
  }
}

function findPlaceRequest_(sheet, requestId) {
  if (sheet.getLastRow() < 2) return false;
  return !!sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).createTextFinder(requestId).matchEntireCell(true).findNext();
}

function placeDateISO_(value) {
  if (value == null || value === '' || typeof value === 'boolean') return '';
  var date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function listPlaces_() {
  var sheet = getBook_().getSheetByName(PLACES_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // 以追加順序取最新 50 列；僅讀取地點工作表的三欄。
  var count = Math.min(sheet.getLastRow() - 1, 50);
  var rows = sheet.getRange(sheet.getLastRow() - count + 1, 1, count, 3).getValues();
  return rows.reverse().filter(function (row) {
    return validRequestId_(row[2]) && typeof row[1] === 'string';
  }).map(function (row) {
    // Sheets 可能已消除保護用引號；本機或舊資料若保留則還原顯示。
    var place = row[1].replace(/^'(?=[=+\-@])/, '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
    return {id: row[2], place: place, createdAt: placeDateISO_(row[0])};
  }).filter(function (item) {
    return item.place.length > 0 && Array.from(item.place).length <= 60;
  });
}
