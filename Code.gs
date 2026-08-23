/**
 * JOYFIT24経堂 見学・体験予約 GAS
 * Script ID: 1DzpjNAV0xDt4GVdaVZzMBMVfjEsWyOiOrQh8sOdKZxD6P0jx-q20kRCi
 *
 * 初回セットアップ:
 * 1. スプレッドシート → 拡張機能 → Apps Script
 * 2. sendTestMail_ を実行してメール権限を承認
 * 3. メニュー「見学体験メール」→ 前日リマインドトリガーを設定
 */

var RESERVE_SHEET_NAME = '見学体験申請';
var EMAIL_COL = 4; // D列
var DATE_COL = 8; // H列
var TIME_COL = 9; // I列
var CUSTOMER_MAIL_COL = 10; // J列
var ADMIN_MAIL_COL = 11; // K列
var REMINDER_SENT_COL = 12; // L列
var STORE_NAME = 'JOYFIT24 経堂';

// 管理者通知先（店舗メール + 管理者）
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
    ensureMailHeaders_(sh);
    sh.appendRow([new Date(), plan, name, email, tel, gender, age, date, time, '', '', '']);
    var row = sh.getLastRow();

    var booking = {
      row: row,
      plan: plan,
      name: name,
      email: email,
      tel: tel,
      gender: gender,
      age: age,
      date: date,
      time: time,
      displayDate: formatDisplayDateFromText_(date)
    };

    var mailResult = sendBookingMails_(booking);
    writeMailStatus_(sh, row, mailResult);

    return {
      ok: true,
      mailSent: mailResult.customer.ok && mailResult.admin.ok,
      mailDetail: mailResult
    };
  } finally {
    lock.releaseLock();
  }
}

function sendBookingMails_(booking) {
  var customerSubject = '【' + STORE_NAME + '】見学・体験予約を承りました';
  var adminSubject = '【' + STORE_NAME + '】見学・体験の新規予約（管理者通知）';

  var customerBody = buildCustomerMailBody_(booking, false);
  var adminBody = buildAdminMailBody_(booking);

  var customer = sendMail_({
    to: booking.email,
    subject: customerSubject,
    body: customerBody
  });

  var admin = sendMail_({
    to: getAdminRecipients_().join(','),
    subject: adminSubject,
    body: adminBody
  });

  return { customer: customer, admin: admin };
}

function sendMail_(options) {
  var to = String(options.to || '').trim();
  var subject = String(options.subject || '').trim();
  var body = String(options.body || '');
  if (!to || !subject) {
    return { ok: false, error: 'missing to/subject' };
  }

  try {
    GmailApp.sendEmail(to, subject, body, { name: STORE_NAME });
    return { ok: true, via: 'GmailApp', to: to };
  } catch (gmailErr) {
    try {
      MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body,
        name: STORE_NAME
      });
      return { ok: true, via: 'MailApp', to: to };
    } catch (mailErr) {
      return { ok: false, error: String(mailErr), gmailError: String(gmailErr) };
    }
  }
}

function buildCustomerMailBody_(data, isReminder) {
  var lines = [
    data.name + ' 様',
    '',
    STORE_NAME + 'です。',
    isReminder
      ? '明日のご来館について、リマインドのご連絡です。'
      : '以下の内容で見学・体験の予約を承りました。',
    '',
    '■ 種別：' + data.plan,
    '■ 日時：' + data.displayDate + ' ' + data.time,
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

function buildAdminMailBody_(data) {
  var lines = [
    STORE_NAME + ' 管理者各位',
    '',
    '見学・体験の新規予約が入りました。',
    '',
    '■ 種別：' + data.plan,
    '■ 日時：' + data.displayDate + ' ' + data.time,
    '■ お名前：' + data.name,
    '■ メール：' + data.email,
    '■ 電話：' + data.tel,
    '■ 性別：' + (data.gender || '-'),
    '■ 年代：' + (data.age || '-'),
    '■ シート行：' + data.row,
    '',
    'スプレッドシート「' + RESERVE_SHEET_NAME + '」をご確認ください。'
  ];
  return lines.join('\n');
}

function sendDayBeforeReminders() {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  ensureMailHeaders_(sh);
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowKey = formatDateKey_(tomorrow);

  var values = sh.getRange(2, 1, lastRow, REMINDER_SENT_COL).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 2;
    if (String(row[REMINDER_SENT_COL - 1] || '').trim()) continue;

    var visitDate = parseVisitDate_(row[DATE_COL - 1]);
    if (!visitDate || formatDateKey_(visitDate) !== tomorrowKey) continue;

    var email = normalizeEmail_(row[EMAIL_COL - 1]);
    if (!isValidEmail_(email)) continue;

    var booking = {
      row: rowNum,
      plan: String(row[1] || '').trim(),
      name: String(row[2] || '').trim(),
      email: email,
      tel: String(row[4] || '').trim(),
      gender: String(row[5] || '').trim(),
      age: String(row[6] || '').trim(),
      date: String(row[DATE_COL - 1] || '').trim(),
      time: String(row[TIME_COL - 1] || '').trim(),
      displayDate: formatDisplayDate_(visitDate)
    };

    var customer = sendMail_({
      to: booking.email,
      subject: '【' + STORE_NAME + '】明日のご来館リマインド',
      body: buildCustomerMailBody_(booking, true)
    });
    var admin = sendMail_({
      to: getAdminRecipients_().join(','),
      subject: '【' + STORE_NAME + '】明日来館予定（管理者リマインド）',
      body: buildAdminMailBody_(booking)
    });

    if (customer.ok && admin.ok) {
      sh.getRange(rowNum, REMINDER_SENT_COL).setValue(new Date());
    } else {
      sh.getRange(rowNum, REMINDER_SENT_COL).setValue('ERROR: ' + (customer.error || admin.error || 'unknown'));
    }
  }
}

function resendMissingBookingMails() {
  var sh = getReserveSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  ensureMailHeaders_(sh);
  var values = sh.getRange(2, 1, lastRow, REMINDER_SENT_COL).getValues();
  var count = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 2;
    var customerStatus = String(row[CUSTOMER_MAIL_COL - 1] || '').trim();
    var adminStatus = String(row[ADMIN_MAIL_COL - 1] || '').trim();
    if (customerStatus && adminStatus && customerStatus.indexOf('ERROR') !== 0 && adminStatus.indexOf('ERROR') !== 0) {
      continue;
    }

    var email = normalizeEmail_(row[EMAIL_COL - 1]);
    if (!isValidEmail_(email)) continue;

    var booking = {
      row: rowNum,
      plan: String(row[1] || '').trim(),
      name: String(row[2] || '').trim(),
      email: email,
      tel: String(row[4] || '').trim(),
      gender: String(row[5] || '').trim(),
      age: String(row[6] || '').trim(),
      date: String(row[DATE_COL - 1] || '').trim(),
      time: String(row[TIME_COL - 1] || '').trim(),
      displayDate: formatDisplayDateFromText_(String(row[DATE_COL - 1] || '').trim())
    };

    var mailResult = sendBookingMails_(booking);
    writeMailStatus_(sh, rowNum, mailResult);
    if (mailResult.customer.ok && mailResult.admin.ok) count++;
  }

  SpreadsheetApp.getUi().alert('未送信メール再送完了: ' + count + '件');
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
    date: '2026-08-25',
    time: '13:30',
    displayDate: '8/25'
  };

  var customer = sendMail_({
    to: me,
    subject: '【' + STORE_NAME + '】テスト：申込者メール',
    body: buildCustomerMailBody_(booking, false)
  });
  var admin = sendMail_({
    to: getAdminRecipients_().join(','),
    subject: '【' + STORE_NAME + '】テスト：管理者通知',
    body: buildAdminMailBody_(booking)
  });

  SpreadsheetApp.getUi().alert(
    'テスト送信結果\n' +
    '申込者: ' + (customer.ok ? 'OK (' + customer.via + ')' : customer.error) + '\n' +
    '管理者: ' + (admin.ok ? 'OK (' + admin.via + ')' : admin.error)
  );
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
  SpreadsheetApp.getUi().alert('前日18時のリマインドトリガーを設定しました');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('見学体験メール')
    .addItem('テストメール送信（権限承認）', 'sendTestMail_')
    .addItem('未送信分を再送', 'resendMissingBookingMails')
    .addSeparator()
    .addItem('前日リマインドトリガーを設定', 'setupReminderTrigger')
    .addToUi();
}

function writeMailStatus_(sh, row, mailResult) {
  var customerVal = mailResult.customer.ok
    ? new Date()
    : ('ERROR: ' + (mailResult.customer.error || 'unknown'));
  var adminVal = mailResult.admin.ok
    ? new Date()
    : ('ERROR: ' + (mailResult.admin.error || 'unknown'));
  sh.getRange(row, CUSTOMER_MAIL_COL).setValue(customerVal);
  sh.getRange(row, ADMIN_MAIL_COL).setValue(adminVal);
}

function getAdminRecipients_() {
  var list = ADMIN_EMAILS.slice();
  var me = Session.getActiveUser().getEmail();
  if (me && list.indexOf(me) === -1) list.push(me);
  return list.filter(function(email) {
    return isValidEmail_(normalizeEmail_(email));
  });
}

function ensureMailHeaders_(sh) {
  var headers = sh.getRange(1, 1, 1, REMINDER_SENT_COL).getValues()[0];
  if (!headers[CUSTOMER_MAIL_COL - 1]) sh.getRange(1, CUSTOMER_MAIL_COL).setValue('申込者メール');
  if (!headers[ADMIN_MAIL_COL - 1]) sh.getRange(1, ADMIN_MAIL_COL).setValue('管理者通知');
  if (!headers[REMINDER_SENT_COL - 1]) sh.getRange(1, REMINDER_SENT_COL).setValue('前日リマインド');
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

function parseVisitDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
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
    return isNaN(dt.getTime()) ? null : dt;
  }
  var dt = new Date(text);
  return isNaN(dt.getTime()) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
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
