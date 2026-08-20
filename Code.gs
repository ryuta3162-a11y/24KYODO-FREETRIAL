/**
 * JOYFIT24経堂 見学・体験予約
 * 送信時にスプレッドシートのメール重複を判定（1メール1回）。認証コードなし。
 */

var RESERVE_SHEET_NAME = '予約';

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var result;
  try {
    var action = String(p.action || '');
    if (action === 'checkEmail') {
      result = checkEmail_(p.email);
    } else if (action === 'submit') {
      result = submitReservation_(p);
    } else {
      result = { ok: false, error: 'unknownAction' };
    }
  } catch (err) {
    result = { ok: false, error: 'system', detail: String(err) };
  }
  return jsonp_(p.callback, result);
}

function checkEmail_(rawEmail) {
  var email = normalizeEmail_(rawEmail);
  if (!isValidEmail_(email)) return { ok: false, error: 'invalidEmail' };
  return { ok: true, alreadyBooked: isEmailAlreadyBooked_(email) };
}

function submitReservation_(p) {
  var email = normalizeEmail_(p.email);
  if (!isValidEmail_(email)) return { ok: false, error: 'invalidEmail' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (isEmailAlreadyBooked_(email)) return { ok: false, error: 'alreadyBooked' };

    var name = String(p.name || '').trim();
    var tel = String(p.tel || '').trim();
    var plan = String(p.plan || '').trim();
    var gender = String(p.gender || '').trim();
    var age = String(p.age || '').trim();
    var date = String(p.date || '').trim();
    var time = String(p.time || '').trim();
    if (!name || !tel || !plan || !date || !time) {
      return { ok: false, error: 'missingFields' };
    }

    getReserveSheet_().appendRow([new Date(), plan, name, email, tel, gender, age, date, time]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function isEmailAlreadyBooked_(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var lastRow = sh.getLastRow();
    if (lastRow < 1) continue;
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var start = looksLikeHeader_(values[0]) ? 1 : 0;
    for (var r = start; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cell = String(values[r][c] || '');
        if (cell.indexOf('@') === -1) continue;
        if (normalizeEmail_(cell) === email) return true;
      }
    }
  }
  return false;
}

function looksLikeHeader_(row) {
  var joined = row.join(' ');
  return /メール|email|お名前|受付/i.test(joined);
}

function getReserveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RESERVE_SHEET_NAME);
  if (!sh) sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(['受付日時', '種別', 'お名前', 'メール', '電話', '性別', '年代', '希望日', '時刻']);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function normalizeEmail_(raw) {
  return String(raw || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function jsonp_(callback, obj) {
  var name = String(callback || '').replace(/[^A-Za-z0-9_]/g, '');
  var payload = JSON.stringify(obj);
  if (!name) {
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(name + '(' + payload + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
