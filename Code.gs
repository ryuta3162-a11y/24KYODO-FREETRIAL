/**
 * JOYFIT24経堂 見学・体験予約
 * 送信時にスプレッドシートのメール重複を判定（1メール1回）。認証コードなし。
 */

var RESERVE_SHEET_NAME = '見学体験申請';
var EMAIL_COL = 4; // D列

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
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var values = sh.getRange(2, EMAIL_COL, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (normalizeEmail_(values[i][0]) === email) return true;
  }
  return false;
}

function getReserveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RESERVE_SHEET_NAME);
  if (!sh) sh = ss.getSheets()[0];
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
