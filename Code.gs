/**
 * 「融的小小關心所」後台 v3.1
 * 安裝位置：目標 Google 試算表 → 擴充功能 → Apps Script
 * 更新本檔後：部署 → 管理部署作業 → 鉛筆 → 版本「新版本」→ 部署。
 *
 * POST：寫入狀態與自願留言；使用 requestId 避免重複。
 * GET：僅確認該 requestId 是否已寫入，不回傳留言或任何私人資料。
 */
var SHEET_NAME = '回覆紀錄';
var ALLOWED_CHOICES = ['要去上班了', '還在上班', '準備下班', '下班了', '午休吃飯', '準備開會'];

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var choice = String(data.choice || '');
    var requestId = String(data.requestId || '');
    var note = String(data.note || '').trim();

    if (ALLOWED_CHOICES.indexOf(choice) === -1) throw new Error('Invalid choice');
    if (!/^[a-zA-Z0-9_-]{12,90}$/.test(requestId)) throw new Error('Invalid request id');
    if (Array.from(note).length > 120) throw new Error('Note is too long');
    if (note.length > 600) throw new Error('Invalid note');
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
      CacheService.getScriptCache().put('cheer_' + requestId, '1', 21600);
    } finally {
      lock.releaseLock();
    }
    return ContentService.createTextOutput('OK');
  } catch (err) {
    console.error('doPost failed: ' + String(err));
    return ContentService.createTextOutput('ERROR');
  }
}

function doGet(e) {
  var parameter = (e && e.parameter) || {};
  var callback = String(parameter.callback || '');
  if (!/^__cheerAck_[0-9]+_[a-z0-9]+$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }
  var result = {ok: true, received: false};
  try {
    var requestId = String(parameter.id || '');
    if (parameter.action !== 'status' || !/^[a-zA-Z0-9_-]{12,90}$/.test(requestId)) {
      result = {ok:false,error:'Invalid request'};
    } else {
      var cached = CacheService.getScriptCache().get('cheer_' + requestId);
      var received = cached === '1';
      if (!received) {
        var book = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = book && book.getSheetByName(SHEET_NAME);
        received = !!(sheet && findRequest_(sheet, requestId));
      }
      result.received = received;
      if (received) result.serverTime = new Date().toISOString();
    }
  } catch (err) {
    console.error('doGet failed: ' + String(err));
    result = {ok:false,error:'Temporary service error'};
  }
  // JavaScript JSONP: 讓放在 GitHub Pages 的網站能讀取「後台已確認」結果。
  // 回傳內容只有收件成功與否，不公開試算表或留言。
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('請先從 Google 試算表的「擴充功能」開啟 Apps Script。');
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
