/**
 * JOYFIT24経堂 見学・体験予約
 * - 送信時にスプレッドシートのメール重複を判定（1メール1回）
 * - 申込完了メールを自動送信
 * - 前日リマインド（時間ベーストリガー）
 */

var RESERVE_SHEET_NAME = '見学体験申請';
var EMAIL_COL = 4; // D列
var DATE_COL = 8; // H列
var TIME_COL = 9; // I列
var REMINDER_SENT_COL = 10; // J列（前日リマインド送信済）
var STORE_NAME = 'JOYFIT24 経堂';

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

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('postData missing');
    }
    var data = JSON.parse(e.postData.contents);
    var result = submitReservation_(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'system', detail: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

    var sh = getReserveSheet_();
    sh.appendRow([new Date(), plan, name, email, tel, gender, age, date, time, '']);
    var row = sh.getLastRow();

    try {
      sendConfirmationEmail_({ plan: plan, name: name, email: email, tel: tel, date: date, time: time });
    } catch (mailErr) {
      Logger.log('confirmation mail failed row=' + row + ' ' + mailErr);
    }

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function sendConfirmationEmail_(data) {
  var subject = '【' + STORE_NAME + '】見学・体験予約を承りました';
  var body = buildReservationMailBody_(data, false);
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: body,
    name: STORE_NAME
  });
}

function sendDayBeforeReminders() {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  ensureReminderHeader_(sh);
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowKey = formatDateKey_(tomorrow);

  var values = sh.getRange(2, 1, lastRow, REMINDER_SENT_COL).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 2;
    var sentFlag = String(row[REMINDER_SENT_COL - 1] || '').trim();
    if (sentFlag) continue;

    var visitDate = parseVisitDate_(row[DATE_COL - 1]);
    if (!visitDate || formatDateKey_(visitDate) !== tomorrowKey) continue;

    var email = normalizeEmail_(row[EMAIL_COL - 1]);
    if (!isValidEmail_(email)) continue;

    var data = {
      plan: String(row[1] || '').trim(),
      name: String(row[2] || '').trim(),
      email: email,
      tel: String(row[4] || '').trim(),
      date: formatDisplayDate_(visitDate),
      time: String(row[TIME_COL - 1] || '').trim()
    };

    try {
      MailApp.sendEmail({
        to: email,
        subject: '【' + STORE_NAME + '】明日のご来館リマインド',
        body: buildReservationMailBody_(data, true),
        name: STORE_NAME
      });
      sh.getRange(rowNum, REMINDER_SENT_COL).setValue(new Date());
    } catch (err) {
      Logger.log('day-before reminder failed row=' + rowNum + ' ' + err);
    }
  }
}

function buildReservationMailBody_(data, isReminder) {
  var lines = [
    data.name + ' 様',
    '',
    STORE_NAME + 'です。',
    isReminder
      ? '明日のご来館について、リマインドのご連絡です。'
      : '以下の内容で見学・体験の予約を承りました。',
    '',
    '■ 種別：' + data.plan,
    '■ 日時：' + data.date + ' ' + data.time,
    '■ お名前：' + data.name,
    '■ 電話番号：' + data.tel,
    '',
    '【ご来館時のお願い】',
    '・スタッフ常駐時間内にご来館ください',
    '  平日 10:00-20:00 / 土日祝 12:00-19:00',
    '・月曜・木曜はスタッフ不在のため予約不可',
    '・体験は60分を目安にご利用ください',
    '・館内は土足でのご利用が可能です',
    '',
    'ご来館をお待ちしております。',
    '',
    STORE_NAME
  ];
  return lines.join('\n');
}

function setupReminderTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDayBeforeReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendDayBeforeReminders')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
}

function sendTestMail_() {
  var email = Session.getActiveUser().getEmail();
  sendConfirmationEmail_({
    plan: '見学',
    name: 'テスト太郎',
    email: email,
    tel: '09012345678',
    date: '2026/08/24',
    time: '14:00'
  });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('見学体験メール')
    .addItem('前日リマインドトリガーを設定', 'setupReminderTrigger')
    .addItem('テストメール送信', 'sendTestMail_')
    .addToUi();
}

function isEmailAlreadyBooked_(email) {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var values = sh.getRange(2, EMAIL_COL, lastRow, EMAIL_COL).getValues();
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

function ensureReminderHeader_(sh) {
  var header = sh.getRange(1, REMINDER_SENT_COL).getValue();
  if (!header) {
    sh.getRange(1, REMINDER_SENT_COL).setValue('前日リマインド送信');
  }
}

function parseVisitDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var text = String(value || '').trim();
  if (!text) return null;
  var normalized = text.replace(/-/g, '/');
  var parts = normalized.split('/');
  if (parts.length >= 3) {
    var y = Number(parts[0]);
    var m = Number(parts[1]) - 1;
    var d = Number(parts[2]);
    var dt = new Date(y, m, d);
    return isNaN(dt.getTime()) ? null : dt;
  }
  var dt = new Date(text);
  return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function formatDateKey_(date) {
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

function formatDisplayDate_(date) {
  return (date.getMonth() + 1) + '/' + date.getDate();
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
