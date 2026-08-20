/**
 * JOYFIT24経堂 見学・体験予約（メール認証＋1メール1回）
 *
 * 【使い方】
 * 1. いま予約を受け取っているスプレッドシートの Apps Script を開く
 * 2. このファイルの内容で置き換えて保存
 * 3. 初回は sendTestMail_ を実行して権限承認
 * 4. デプロイ → ウェブアプリ → 新しいバージョン（実行: 自分 / アクセス: 全員）
 *
 * シート「予約」に1行1予約、「_認証」にコードを保存します。
 */

var AUTH_SHEET_NAME = '_認証';
var RESERVE_SHEET_NAME = '予約';
var CODE_TTL_MS = 10 * 60 * 1000;
var SEND_COOLDOWN_MS = 60 * 1000;

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var result;
  try {
    var action = String(p.action || '');
    if (action === 'sendCode') {
      result = sendCode_(p.email);
    } else if (action === 'verifyCode') {
      result = verifyCode_(p.email, p.code);
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

function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'useVerifiedFlow' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendTestMail_() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'JOYFIT24経堂 認証メールテスト', '権限OKです。');
}

function sendCode_(rawEmail) {
  var email = normalizeEmail_(rawEmail);
  if (!isValidEmail_(email)) return { ok: false, error: 'invalidEmail' };
  if (isEmailAlreadyBooked_(email)) return { ok: false, error: 'alreadyBooked' };

  var sh = getAuthSheet_();
  var row = findAuthRow_(sh, email);
  var now = new Date();
  if (row > 0) {
    var lastSent = sh.getRange(row, 5).getValue();
    if (lastSent instanceof Date && now.getTime() - lastSent.getTime() < SEND_COOLDOWN_MS) {
      return { ok: false, error: 'tooSoon' };
    }
  }

  var code = String(Math.floor(100000 + Math.random() * 900000));
  var expiry = new Date(now.getTime() + CODE_TTL_MS);
  if (row > 0) {
    sh.getRange(row, 1, 1, 6).setValues([[email, code, expiry, '', lastSentSafe_(now), '']]);
  } else {
    sh.appendRow([email, code, expiry, '', now, '']);
  }

  MailApp.sendEmail({
    to: email,
    subject: '【JOYFIT24経堂】見学・体験の認証コード',
    body: [
      'JOYFIT24経堂です。',
      '',
      '見学・体験予約の認証コードは次の6桁です。',
      '',
      code,
      '',
      '有効期限は10分です。',
      'このメールに心当たりがない場合は破棄してください。',
      '',
      'JOYFIT24経堂'
    ].join('\n')
  });

  return { ok: true };
}

function verifyCode_(rawEmail, rawCode) {
  var email = normalizeEmail_(rawEmail);
  var code = String(rawCode || '').replace(/\s/g, '');
  if (!isValidEmail_(email) || !/^\d{6}$/.test(code)) {
    return { ok: false, error: 'invalidCode' };
  }
  if (isEmailAlreadyBooked_(email)) return { ok: false, error: 'alreadyBooked' };

  var sh = getAuthSheet_();
  var row = findAuthRow_(sh, email);
  if (row < 1) return { ok: false, error: 'invalidCode' };

  var stored = String(sh.getRange(row, 2).getValue() || '');
  var expiry = sh.getRange(row, 3).getValue();
  if (stored !== code) return { ok: false, error: 'invalidCode' };
  if (!(expiry instanceof Date) || expiry.getTime() < Date.now()) {
    return { ok: false, error: 'expired' };
  }

  var token = Utilities.getUuid().replace(/-/g, '');
  sh.getRange(row, 4).setValue(token);
  sh.getRange(row, 6).setValue(new Date());
  return { ok: true, token: token };
}

function submitReservation_(p) {
  var email = normalizeEmail_(p.email);
  var token = String(p.token || '').trim();
  if (!isValidEmail_(email) || !token) return { ok: false, error: 'notVerified' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (isEmailAlreadyBooked_(email)) return { ok: false, error: 'alreadyBooked' };
    if (!isValidToken_(email, token)) return { ok: false, error: 'notVerified' };

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

    var sh = getReserveSheet_();
    sh.appendRow([new Date(), plan, name, email, tel, gender, age, date, time]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function isValidToken_(email, token) {
  var sh = getAuthSheet_();
  var row = findAuthRow_(sh, email);
  if (row < 1) return false;
  var stored = String(sh.getRange(row, 4).getValue() || '');
  return stored && stored === token;
}

function isEmailAlreadyBooked_(email) {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cell = String(values[r][c] || '');
      if (cell.indexOf('@') === -1) continue;
      if (normalizeEmail_(cell) === email) return true;
    }
  }
  return false;
}

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getReserveSheet_() {
  var ss = getSs_();
  var sh = ss.getSheetByName(RESERVE_SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(['受付日時', '種別', 'お名前', 'メール', '電話', '性別', '年代', '希望日', '時刻']);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAuthSheet_() {
  var ss = getSs_();
  var sh = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(AUTH_SHEET_NAME);
    sh.appendRow(['email', 'code', 'expiry', 'token', 'lastSent', 'verifiedAt']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function findAuthRow_(sh, email) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (normalizeEmail_(vals[i][0]) === email) return i + 2;
  }
  return 0;
}

function normalizeEmail_(raw) {
  return String(raw || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function lastSentSafe_(now) {
  return now;
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
