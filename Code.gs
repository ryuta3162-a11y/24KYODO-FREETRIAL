/**
 * JOYFIT24経堂 見学・体験予約 GAS
 * Script ID: 1DzpjNAV0xDt4GVdaVZzMBMVfjEsWyOiOrQh8sOdKZxD6P0jx-q20kRCi
 *
 * 初回セットアップ（スプレッドシートを開いて1回だけ）:
 * 1. メニュー「見学体験メール」→ 初期セットアップ（トリガー一括設定）
 * 2. メニュー「見学体験メール」→ テストメール送信（権限承認）
 * ※前日リマインドは廃止
 */

var RESERVE_SHEET_NAME = '見学体験申請';
var MAIL_LOG_SHEET_NAME = '_メール送信';
var EMAIL_COL = 4; // D列
var DATE_COL = 8; // H列
var TIME_COL = 9; // I列
var RESERVE_LAST_COL = 9; // A〜I のみ
var STORE_NAME = 'JOYFIT24 経堂';
var STORE_NAME_SHORT = 'JOYFIT24経堂';
var STORE_EMAIL = 'jf-kyoudou@okamoto-group.co.jp';
var WEEKDAYS_JA = ['日', '月', '火', '水', '木', '金', '土'];
var MAIL_RETRY_MAX = 2;
var MAIL_RETRY_WAIT_MS = 800;
var TRIGGER_PROP_KEY = 'MAIL_SYSTEM_CONFIGURED_AT';

// 管理者通知先（店舗 + 社員）。お客さま宛メールには社員を出さない
var ADMIN_EMAILS = [
  'jf-kyoudou@okamoto-group.co.jp',
  'r-kusaka@okamoto-group.co.jp',
  'mito-sato@okamoto-group.co.jp',
  'yuka-hachiya@okamoto-group.co.jp'
];

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var result;
  try {
    var action = String(p.action || '');
    if (action === 'checkEmail') {
      result = checkEmail_(p.email);
    } else if (action === 'getBookedSlots') {
      result = getBookedSlots_(p.date);
    } else if (action === 'submit') {
      result = submitReservation_(p);
    } else {
      result = { ok: false, error: 'unknownAction' };
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
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
    var action = String(data.action || 'submit');
    var result;
    if (action === 'checkEmail') {
      result = checkEmail_(data.email);
    } else if (action === 'getBookedSlots') {
      result = getBookedSlots_(data.date);
    } else if (action === 'submit') {
      result = submitReservation_(data);
    } else {
      result = { ok: false, error: 'unknownAction' };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err);
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

/** 指定日の予約済み時刻一覧（UIグレーアウト用・軽量） */
function getBookedSlots_(rawDate) {
  var dateText = formatDateInputValue_(rawDate);
  if (!dateText) return { ok: false, error: 'missingFields' };
  var visitDate = parseVisitDate_(dateText);
  var dateKey = visitDate ? formatDateKey_(visitDate) : dateText;
  return { ok: true, date: dateKey, times: getBookedTimesForDate_(dateKey) };
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

    var timeText = formatTimeValue_(time);
    var dateText = formatDateInputValue_(date);
    var visitDate = parseVisitDate_(dateText);
    if (!timeText || !dateText) return { ok: false, error: 'missingFields' };

    // 同時刻の二重予約防止（画面のグレーアウトと併用）
    if (isSlotAlreadyBooked_(dateText, timeText)) {
      return { ok: false, error: 'slotTaken' };
    }

    var sh = getReserveSheet_();
    sh.appendRow([new Date(), plan, name, email, tel, gender, age, dateText, timeText]);
    var row = sh.getLastRow();
    sh.getRange(row, DATE_COL).setNumberFormat('@').setValue(dateText);
    sh.getRange(row, TIME_COL).setNumberFormat('@').setValue(timeText);
    SpreadsheetApp.flush();

    var booking = {
      row: row,
      plan: plan,
      name: name,
      email: email,
      tel: tel,
      gender: gender,
      age: age,
      date: dateText,
      time: timeText,
      displayDate: visitDate ? formatDisplayDate_(visitDate) : dateText,
      visitDateKey: visitDate ? formatDateKey_(visitDate) : ''
    };

    // メール失敗でも申込自体は成功にする（LPのシステムエラー防止）
    var mailResult = { customer: { ok: false }, admin: { ok: false } };
    try {
      mailResult = sendBookingMails_(booking);
    } catch (mailErr) {
      Logger.log('sendBookingMails_ error row=' + row + ' ' + mailErr);
      mailResult = { customer: { ok: false, error: String(mailErr) }, admin: { ok: false, error: String(mailErr) } };
    }
    try {
      writeMailStatus_(row, booking.email, mailResult);
    } catch (logErr) {
      Logger.log('writeMailStatus_ error: ' + logErr);
    }

    return {
      ok: true,
      mailSent: !!(mailResult.customer && mailResult.customer.ok && mailResult.admin && mailResult.admin.ok)
    };
  } finally {
    lock.releaseLock();
  }
}

function sendBookingMails_(booking) {
  // お客さまへ（当初文面）※社員BCCなし
  var customer = sendMailWithRetry_({
    to: booking.email,
    subject: '【' + STORE_NAME + '】見学・体験予約を承りました',
    body: buildCustomerMailBody_(booking)
  });

  // 店舗宛 + 社員はBCC（スプシリンク付き管理者通知）
  var admin = sendMailWithRetry_({
    to: STORE_EMAIL,
    bcc: getStaffBcc_(),
    subject: '【' + STORE_NAME_SHORT + '】 見学・体験の申し込みがありました',
    body: buildAdminMailBody_(booking)
  });

  return { customer: customer, admin: admin };
}

function sendMailWithRetry_(options) {
  var lastResult = { ok: false, error: 'not attempted' };
  for (var attempt = 1; attempt <= MAIL_RETRY_MAX; attempt++) {
    lastResult = sendMailOnce_(options);
    if (lastResult.ok) {
      lastResult.attempts = attempt;
      return lastResult;
    }
    if (attempt < MAIL_RETRY_MAX) {
      Utilities.sleep(MAIL_RETRY_WAIT_MS * attempt);
    }
  }
  lastResult.attempts = MAIL_RETRY_MAX;
  return lastResult;
}

function sendMailOnce_(options) {
  var to = String(options.to || '').trim();
  var subject = String(options.subject || '').trim();
  var body = String(options.body || '');
  var cc = String(options.cc || '').trim();
  var bcc = String(options.bcc || '').trim();
  if (!to || !subject) {
    return { ok: false, error: 'missing to/subject' };
  }

  // Fromは実行アカウントのまま。Reply-Toのみ店舗（send-as未設定でも落ちない）
  var extras = {
    name: STORE_NAME,
    replyTo: STORE_EMAIL
  };
  if (cc) extras.cc = cc;
  if (bcc) extras.bcc = bcc;

  try {
    GmailApp.sendEmail(to, subject, body, extras);
    return { ok: true, via: 'GmailApp', to: to };
  } catch (gmailErr) {
    try {
      MailApp.sendEmail({
        to: to,
        cc: cc || undefined,
        bcc: bcc || undefined,
        subject: subject,
        body: body,
        name: STORE_NAME,
        replyTo: STORE_EMAIL
      });
      return { ok: true, via: 'MailApp', to: to };
    } catch (mailErr) {
      return { ok: false, error: String(mailErr), gmailError: String(gmailErr) };
    }
  }
}

function buildCustomerMailBody_(data) {
  return [
    data.name + ' 様',
    '',
    STORE_NAME + 'です。',
    '以下の内容で見学・体験の予約を承りました。',
    '■ 種別：' + data.plan,
    '■ 日時：' + data.displayDate + ' ' + data.time,
    '■ お名前：' + data.name,
    '■ 電話番号：' + data.tel,
    '【ご来館時のお願い】',
    '・キャンセル時はこちらのメールにご連絡ください。',
    STORE_EMAIL,
    '・体験は60分を目安にご利用ください',
    '・館内は土足でのご利用が可能です',
    '・当日は入口のインターホンを押してください。',
    'スタッフがご案内します。',
    '',
    'ご来館をお待ちしております。',
    STORE_NAME
  ].join('\n');
}

function buildAdminMailBody_(data) {
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return [
    '【' + STORE_NAME_SHORT + '】 見学・体験の申し込みがありました。下記の内容をご確認ください。',
    '',
    '■ お申し込み内容',
    'プラン：' + data.plan,
    'お名前：' + data.name + ' 様',
    'メールアドレス：' + data.email,
    '電話番号：' + data.tel,
    '性\u3000別：' + (data.gender || '-'),
    '年\u3000代：' + (data.age || '-'),
    'ご希望日時：' + formatAdminDateTime_(data.date, data.time),
    '',
    'スプレッドシートにも記録されています。',
    sheetUrl
  ].join('\n');
}

function retryFailedMails() {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var logMap = getMailLogMap_();
  var keys = Object.keys(logMap);
  for (var k = 0; k < keys.length; k++) {
    var rowNum = Number(keys[k]);
    var status = logMap[rowNum];
    if (isMailStatusOk_(status.customer) && isMailStatusOk_(status.admin)) continue;
    if (rowNum < 2 || rowNum > lastRow) continue;

    var row = sh.getRange(rowNum, 1, rowNum, RESERVE_LAST_COL).getValues()[0];
    var booking = buildBookingFromValues_(row, rowNum, parseVisitDate_(row[DATE_COL - 1]));
    if (!booking) continue;

    try {
      var mailResult = sendBookingMails_(booking);
      writeMailStatus_(rowNum, booking.email, mailResult);
    } catch (err) {
      Logger.log('retry fail row=' + rowNum + ' ' + err);
    }
  }
}

function setupAllTriggers() {
  // 前日リマインドは廃止。既存トリガーも削除する
  deleteTriggersForHandlers_(['sendDayBeforeReminders', 'retryFailedMails']);
  ScriptApp.newTrigger('retryFailedMails').timeBased().everyHours(1).create();
  clearLegacyMailColumns_(getReserveSheet_());
  getMailLogSheet_();
  PropertiesService.getScriptProperties().setProperty(TRIGGER_PROP_KEY, new Date().toISOString());
  SpreadsheetApp.getUi().alert(
    '初期セットアップ完了\n\n・毎時：送信失敗の自動リトライ\n・前日リマインド：削除済み\n\n続けて「テストメール送信」を実行してください。'
  );
}

function sendTestMail_() {
  var me = Session.getActiveUser().getEmail();
  if (!me) throw new Error('実行ユーザーのメールアドレスを取得できません');

  var booking = {
    row: 0,
    plan: '見学',
    name: 'テスト太郎',
    email: me,
    tel: '09012345678',
    gender: '男性',
    age: '20代',
    date: '2026-08-28',
    time: '13:30',
    displayDate: '8/28'
  };

  var customer = sendMailWithRetry_({
    to: me,
    subject: '【' + STORE_NAME + '】テスト：申込者メール',
    body: buildCustomerMailBody_(booking)
  });
  var admin = sendMailWithRetry_({
    to: STORE_EMAIL,
    bcc: getStaffBcc_(),
    subject: '【' + STORE_NAME_SHORT + '】テスト：管理者通知',
    body: buildAdminMailBody_(booking)
  });

  SpreadsheetApp.getUi().alert(
    'テスト送信結果\n' +
    'お客さま宛: ' + (customer.ok ? 'OK (' + customer.via + ')' : customer.error) + '\n' +
    '店舗・管理者: ' + (admin.ok ? 'OK (' + admin.via + ')' : admin.error)
  );
}

function resendMissingBookingMails() {
  retryFailedMails();
  SpreadsheetApp.getUi().alert('未送信・エラー分の再送処理を実行しました。');
}

function onOpen() {
  try { clearLegacyMailColumns_(getReserveSheet_()); } catch (e) {}
  SpreadsheetApp.getUi()
    .createMenu('見学体験メール')
    .addItem('初期セットアップ（トリガー一括設定）', 'setupAllTriggers')
    .addItem('テストメール送信（権限承認）', 'sendTestMail_')
    .addSeparator()
    .addItem('未送信・エラー分を再送', 'resendMissingBookingMails')
    .addToUi();
}

function writeMailStatus_(row, email, mailResult) {
  var log = upsertMailLogRow_(row, email);
  var sh = log.sheet;
  var logRow = log.row;
  sh.getRange(logRow, 3).setValue(
    mailResult.customer && mailResult.customer.ok
      ? new Date()
      : ('ERROR: ' + ((mailResult.customer && mailResult.customer.error) || 'unknown'))
  );
  sh.getRange(logRow, 4).setValue(
    mailResult.admin && mailResult.admin.ok
      ? new Date()
      : ('ERROR: ' + ((mailResult.admin && mailResult.admin.error) || 'unknown'))
  );
}

function getMailLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MAIL_LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MAIL_LOG_SHEET_NAME);
    sh.hideSheet();
  }
  var headers = sh.getRange(1, 1, 1, 4).getValues()[0];
  if (!headers[0]) sh.getRange(1, 1).setValue('予約行');
  if (!headers[1]) sh.getRange(1, 2).setValue('メール');
  if (!headers[2]) sh.getRange(1, 3).setValue('申込者送信');
  if (!headers[3]) sh.getRange(1, 4).setValue('管理者送信');
  return sh;
}

function getMailLogMap_() {
  var sh = getMailLogSheet_();
  var lastRow = sh.getLastRow();
  var map = {};
  if (lastRow < 2) return map;
  var values = sh.getRange(2, 1, lastRow, 4).getValues();
  for (var i = 0; i < values.length; i++) {
    var reserveRow = Number(values[i][0]);
    if (!reserveRow) continue;
    map[reserveRow] = {
      logRow: i + 2,
      email: String(values[i][1] || ''),
      customer: values[i][2],
      admin: values[i][3]
    };
  }
  return map;
}

function upsertMailLogRow_(reserveRow, email) {
  var sh = getMailLogSheet_();
  var map = getMailLogMap_();
  if (map[reserveRow]) {
    if (email) sh.getRange(map[reserveRow].logRow, 2).setValue(email);
    return { sheet: sh, row: map[reserveRow].logRow };
  }
  sh.appendRow([reserveRow, email || '', '', '', '']);
  return { sheet: sh, row: sh.getLastRow() };
}

function clearLegacyMailColumns_(sh) {
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  if (lastCol < 10) return;
  sh.getRange(1, 10, Math.max(sh.getLastRow(), 1), lastCol - 9).clearContent();
}

function buildBookingFromValues_(row, rowNum, visitDate) {
  var email = normalizeEmail_(row[EMAIL_COL - 1]);
  var name = String(row[2] || '').trim();
  if (!isValidEmail_(email) || !name) return null;

  var resolvedVisitDate = visitDate || parseVisitDate_(row[DATE_COL - 1]);
  var dateText = resolvedVisitDate ? formatDateKey_(resolvedVisitDate) : formatDateInputValue_(row[DATE_COL - 1]);
  var timeText = formatTimeValue_(row[TIME_COL - 1]);

  return {
    row: rowNum,
    plan: String(row[1] || '').trim(),
    name: name,
    email: email,
    tel: formatTelValue_(row[4]),
    gender: String(row[5] || '').trim(),
    age: String(row[6] || '').trim(),
    date: dateText,
    time: timeText,
    displayDate: resolvedVisitDate ? formatDisplayDate_(resolvedVisitDate) : formatDisplayDateFromText_(dateText),
    visitDateKey: resolvedVisitDate ? formatDateKey_(resolvedVisitDate) : ''
  };
}

function getStaffBcc_() {
  return getAdminRecipients_([STORE_EMAIL]).join(',');
}

function getAdminRecipients_(excludeEmails) {
  var skip = {};
  var excludes = excludeEmails || [];
  for (var j = 0; j < excludes.length; j++) {
    skip[normalizeEmail_(excludes[j])] = true;
  }
  var seen = {};
  var list = [];
  for (var i = 0; i < ADMIN_EMAILS.length; i++) {
    var email = normalizeEmail_(ADMIN_EMAILS[i]);
    if (!isValidEmail_(email) || seen[email] || skip[email]) continue;
    seen[email] = true;
    list.push(email);
  }
  return list;
}

function isMailStatusOk_(value) {
  if (!value) return false;
  if (value instanceof Date && !isNaN(value.getTime())) return true;
  var text = String(value).trim();
  return text && text.indexOf('ERROR') !== 0;
}

function deleteTriggersForHandlers_(handlers) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (handlers.indexOf(triggers[i].getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
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

function getBookedTimesForDate_(dateKey) {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, DATE_COL, lastRow, TIME_COL).getValues();
  var seen = {};
  var times = [];
  for (var i = 0; i < values.length; i++) {
    var rowDate = parseVisitDate_(values[i][0]);
    if (!rowDate || formatDateKey_(rowDate) !== dateKey) continue;
    var t = formatTimeValue_(values[i][1]);
    if (!t || seen[t]) continue;
    seen[t] = true;
    times.push(t);
  }
  times.sort();
  return times;
}

function isSlotAlreadyBooked_(dateText, timeText) {
  var visitDate = parseVisitDate_(dateText);
  var dateKey = visitDate ? formatDateKey_(visitDate) : formatDateInputValue_(dateText);
  var wantTime = formatTimeValue_(timeText);
  if (!dateKey || !wantTime) return false;
  var booked = getBookedTimesForDate_(dateKey);
  for (var i = 0; i < booked.length; i++) {
    if (booked[i] === wantTime) return true;
  }
  return false;
}

function getReserveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(RESERVE_SHEET_NAME);
  if (!sh) sh = ss.getSheets()[0];
  return sh;
}

function parseVisitDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return startOfDay_(value);
  }
  return parseVisitDateFromText_(String(value || '').trim());
}

function parseVisitDateFromText_(text) {
  if (!text) return null;
  var normalized = text.replace(/-/g, '/');
  var parts = normalized.split('/');
  if (parts.length >= 3) {
    var y = Number(parts[0]);
    var m = Number(parts[1]) - 1;
    var d = Number(parts[2]);
    var dt = new Date(y, m, d);
    return isNaN(dt.getTime()) ? null : startOfDay_(dt);
  }
  var dt = new Date(text);
  return isNaN(dt.getTime()) ? null : startOfDay_(dt);
}

function formatDisplayDateFromText_(text) {
  var dt = parseVisitDateFromText_(text);
  return dt ? formatDisplayDate_(dt) : text;
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

function formatTimeValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Tokyo', 'HH:mm');
  }
  if (typeof value === 'number' && isFinite(value)) {
    var totalMinutes = Math.round((value % 1) * 24 * 60);
    if (totalMinutes < 0) totalMinutes = 0;
    var hh = Math.floor(totalMinutes / 60) % 24;
    var mm = totalMinutes % 60;
    return ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2);
  }
  var text = String(value || '').trim();
  if (!text) return '';
  var hm = text.match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:\s|$)/);
  if (hm) {
    return ('0' + Number(hm[1])).slice(-2) + ':' + ('0' + Number(hm[2])).slice(-2);
  }
  return text;
}

function formatDateInputValue_(value) {
  var dt = parseVisitDate_(value);
  if (dt) return formatDateKey_(dt);
  return String(value || '').trim();
}

function formatTelValue_(value) {
  if (typeof value === 'number' && isFinite(value)) {
    var digits = String(Math.floor(Math.abs(value)));
    if (digits.length === 10 && /^[789]0/.test(digits)) digits = '0' + digits;
    return digits;
  }
  return String(value || '').trim();
}

function formatAdminDateTime_(dateValue, timeValue) {
  var dt = parseVisitDate_(dateValue);
  var time = formatTimeValue_(timeValue);
  if (!dt) {
    var fallback = formatDisplayDateFromText_(String(dateValue || '').trim());
    return (fallback + (time ? ' ' + time : '')).trim();
  }
  var m = ('0' + (dt.getMonth() + 1)).slice(-2);
  var d = ('0' + dt.getDate()).slice(-2);
  var w = WEEKDAYS_JA[dt.getDay()];
  return m + '/' + d + ' (' + w + ')' + (time ? ' ' + time : '');
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
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
