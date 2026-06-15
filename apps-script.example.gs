/* ============================================================
   מערכת זימון אימונים — אביאל | Backend (Google Apps Script)
   - שומר הכול בגוגל שיטס (נוצר אוטומטית בהרצה ראשונה)
   - מייל לאביאל עם כפתורי אישור/דחייה על כל בקשה חדשה
   - באישור: אירוע ביומן גוגל + זימון אוטומטי למייל המתאמן
   - דוח בוקר יומי עם האימונים של היום
   פריסה: Deploy → New deployment → Web app →
           Execute as: Me | Who has access: Anyone
   ============================================================ */

var ADMIN_KEY    = 'YOUR_ADMIN_KEY';   // קוד הניהול שלך — כל מספר/מחרוזת סודית
var NOTIFY_EMAIL = 'you@example.com';   // המייל שמקבל התראות
var TZ           = 'Asia/Jerusalem';
var APP_URL      = 'https://aviel112.github.io/booking/';

var DAY_NAMES = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];
var MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

var DEFAULT_SETTINGS = {
  slotMin: 60,
  bufferMin: 0,
  minNoticeH: 3,
  maxDaysAhead: 60,
  phone: '',
  meetLink: '',
  twilioSid: '',
  twilioToken: '',
  twilioFrom: '',
  remind1On: true,
  remind1H: 24,
  remind2On: true,
  remind2H: 3,
  reqConfirm: true,
  checkCalendar: true,
  services: [
    { id: 's1', name: 'אימון אישי', dur: 60, desc: 'אימון אחד-על-אחד', free: false, on: true }
  ],
  weekTemplate: [
    {on:true,  start:'07:00', end:'15:00'},
    {on:true,  start:'07:00', end:'15:00'},
    {on:true,  start:'07:00', end:'15:00'},
    {on:true,  start:'07:00', end:'15:00'},
    {on:true,  start:'07:00', end:'15:00'},
    {on:true,  start:'07:00', end:'12:00'},
    {on:false, start:'07:00', end:'12:00'}
  ]
};

/* ---------------- storage ---------------- */
function ss_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss = SpreadsheetApp.create('מערכת זימון אימונים — אביאל');
  props.setProperty('SS_ID', ss.getId());
  return ss;
}
function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}
var BK_HEAD = ['id','date','time','name','phone','email','note','status','createdAt','eventId','confirmed','r1','r2','serviceId','serviceName','dur'];
var SUM_HEAD = ['id','bookingId','date','clientName','clientEmail','measurements','homeworkClient','homeworkTrainer','changes','requests','updates','createdAt','sentAt'];
function bkSheet_()  { return sheet_('בקשות ופגישות', BK_HEAD); }
function avSheet_()  { return sheet_('זמינות', ['date','start','end']); }
function cfgSheet_() { return sheet_('הגדרות', ['key','value']); }
function sumSheet_() { return sheet_('סיכומי פגישות', SUM_HEAD); }
function getSummaries_() {
  return rows_(sumSheet_(), SUM_HEAD).map(function (r) { r.date = normDate_(r.date); return r; });
}

function rows_(sh, head) {
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var o = {_row: i + 1};
    for (var j = 0; j < head.length; j++) o[head[j]] = vals[i][j];
    out.push(o);
  }
  return out;
}
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}
function normTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  var s = String(v);
  return s.length === 4 ? '0' + s : s;
}

function getSettings_() {
  var rows = rows_(cfgSheet_(), ['key','value']);
  var s = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  var NUMS = {slotMin:1, minNoticeH:1, bufferMin:1, maxDaysAhead:1, remind1H:1, remind2H:1};
  var BOOLS = {remind1On:1, remind2On:1, reqConfirm:1, checkCalendar:1};
  rows.forEach(function (r) {
    if (r.key === 'weekTemplate') { try { s.weekTemplate = JSON.parse(r.value); } catch (e) {} }
    else if (r.key === 'services') { try { s.services = JSON.parse(r.value); } catch (e) {} }
    else if (NUMS[r.key]) s[r.key] = Number(r.value);
    else if (BOOLS[r.key]) s[r.key] = String(r.value) === 'true';
    else if (r.key) s[r.key] = String(r.value);
  });
  // תאימות לאחור: אם אין שירותים מוגדרים — צור אחד מ-slotMin
  if (!s.services || !s.services.length) {
    s.services = [{ id: 's1', name: 'אימון אישי', dur: s.slotMin || 60, desc: '', free: false, on: true }];
  }
  return s;
}
function saveSettings_(s) {
  var sh = cfgSheet_();
  sh.clearContents();
  sh.appendRow(['key','value']);
  sh.appendRow(['slotMin', s.slotMin]);
  sh.appendRow(['minNoticeH', s.minNoticeH]);
  sh.appendRow(['bufferMin', s.bufferMin || 0]);
  sh.appendRow(['maxDaysAhead', s.maxDaysAhead || 60]);
  sh.appendRow(['phone', s.phone || '']);
  sh.appendRow(['meetLink', s.meetLink || '']);
  sh.appendRow(['twilioSid',   s.twilioSid   || '']);
  sh.appendRow(['twilioToken', s.twilioToken || '']);
  sh.appendRow(['twilioFrom',  s.twilioFrom  || '']);
  sh.appendRow(['remind1On', s.remind1On !== false]);
  sh.appendRow(['remind1H',  s.remind1H || 24]);
  sh.appendRow(['remind2On', s.remind2On !== false]);
  sh.appendRow(['remind2H',  s.remind2H || 3]);
  sh.appendRow(['reqConfirm', s.reqConfirm !== false]);
  sh.appendRow(['checkCalendar', s.checkCalendar !== false]);
  sh.appendRow(['services', JSON.stringify(s.services || [])]);
  sh.appendRow(['weekTemplate', JSON.stringify(s.weekTemplate)]);
}

function getAvail_() {
  var map = {};
  rows_(avSheet_(), ['date','start','end']).forEach(function (r) {
    if (r.date) map[normDate_(r.date)] = { start: normTime_(r.start), end: normTime_(r.end) };
  });
  return map;
}
function writeAvail_(map) {
  var sh = avSheet_();
  sh.clearContents();
  sh.appendRow(['date','start','end']);
  var dates = Object.keys(map).sort();
  if (dates.length) {
    var data = dates.map(function (d) { return [d, map[d].start, map[d].end]; });
    sh.getRange(2, 1, data.length, 3).setValues(data);
  }
}

function getBookings_() {
  return rows_(bkSheet_(), BK_HEAD).map(function (b) {
    b.date = normDate_(b.date);
    b.time = normTime_(b.time);
    return b;
  });
}
function setBookingFields_(rowNum, fields) {
  var sh = bkSheet_();
  Object.keys(fields).forEach(function (k) {
    var col = BK_HEAD.indexOf(k) + 1;
    if (col > 0) sh.getRange(rowNum, col).setValue(fields[k]);
  });
}

/* ---------------- time helpers ---------------- */
function ilDate_(dateStr, timeStr) {
  var guess = new Date(dateStr + 'T' + timeStr + ':00+02:00');
  var shown = Utilities.formatDate(guess, TZ, 'HH:mm');
  if (shown !== timeStr) {
    var want = parseInt(timeStr, 10), got = parseInt(shown, 10);
    var diff = want - got;
    if (diff > 12) diff -= 24;
    if (diff < -12) diff += 24;
    guess = new Date(guess.getTime() + diff * 3600000);
  }
  return guess;
}
function heDate_(dateStr) {
  var d = ilDate_(dateStr, '12:00');
  var dow = Number(Utilities.formatDate(d, TZ, 'u')) % 7; // u: Mon=1..Sun=7
  return 'יום ' + DAY_NAMES[dow] + ', ' + d.getDate() + ' ב' + MONTHS[Number(Utilities.formatDate(d, TZ, 'M')) - 1];
}

function tmin_(t) { var a = String(t).split(':'); return parseInt(a[0], 10) * 60 + parseInt(a[1] || '0', 10); }
function fmin_(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

/* אינטרוולים תפוסים בתאריך: פגישות קיימות (לפי המשך שלהן) + אירועים ביומן גוגל */
function busyFor_(dateStr, bookings, settings) {
  var out = [];
  bookings.forEach(function (b) {
    if (b.date === dateStr && (b.status === 'pending' || b.status === 'approved')) {
      var d = Number(b.dur) || settings.slotMin || 60;
      out.push({ t: b.time, d: d });
    }
  });
  if (settings.checkCalendar !== false) {
    try {
      var evs = CalendarApp.getDefaultCalendar().getEventsForDay(ilDate_(dateStr, '12:00'));
      evs.forEach(function (ev) {
        if (ev.isAllDayEvent()) return;
        var s = tmin_(Utilities.formatDate(ev.getStartTime(), TZ, 'HH:mm'));
        var e = tmin_(Utilities.formatDate(ev.getEndTime(), TZ, 'HH:mm'));
        if (e > s) out.push({ t: fmin_(s), d: e - s, cal: 1 });
      });
    } catch (e) {}
  }
  return out;
}

/* האם פנוי להתחיל פגישה באורך dur ב-time בתאריך נתון */
function freeAt_(dateStr, time, dur, avail, settings, bookings, busyCache) {
  var win = avail[dateStr];
  if (!win) return false;
  var st = tmin_(time), en = st + dur;
  if (st < tmin_(win.start) || en > tmin_(win.end)) return false;
  if (ilDate_(dateStr, time) <= new Date(Date.now() + settings.minNoticeH * 3600000)) return false;
  var buf = settings.bufferMin || 0;
  var blocks = busyCache || busyFor_(dateStr, bookings, settings);
  for (var i = 0; i < blocks.length; i++) {
    var bs = tmin_(blocks[i].t) - buf, be = tmin_(blocks[i].t) + blocks[i].d + buf;
    if (st < be && en > bs) return false;
  }
  return true;
}

/* רשימת שעות פנויות לתאריך עבור משך נתון */
function slotsForDur_(dateStr, dur, avail, settings, bookings, busyCache) {
  var win = avail[dateStr];
  if (!win) return [];
  var out = [];
  var step = dur + (settings.bufferMin || 0);
  for (var cur = tmin_(win.start); cur + dur <= tmin_(win.end); cur += step) {
    var t = fmin_(cur);
    if (freeAt_(dateStr, t, dur, avail, settings, bookings, busyCache)) out.push(t);
  }
  return out;
}

function sendSms_(to, msg) {
  var s = getSettings_();
  if (!s.twilioSid || !s.twilioToken || !s.twilioFrom) return;
  var phone = String(to).replace(/\D/g, '');
  if (phone.length === 10 && phone.charAt(0) === '0') phone = '972' + phone.slice(1);
  if (phone.charAt(0) !== '+') phone = '+' + phone;
  try {
    UrlFetchApp.fetch(
      'https://api.twilio.com/2010-04-01/Accounts/' + s.twilioSid + '/Messages.json',
      { method: 'post',
        headers: { Authorization: 'Basic ' + Utilities.base64Encode(s.twilioSid + ':' + s.twilioToken) },
        payload: { From: s.twilioFrom, To: phone, Body: msg },
        muteHttpExceptions: true }
    );
  } catch (e) {}
}

/* ---------------- web app entry ---------------- */
function doGet(e) {
  var a = (e.parameter && e.parameter.action) || '';
  if (a === 'state')  return json_(publicState_());
  if (a === 'approve' || a === 'reject') {
    if (e.parameter.key !== ADMIN_KEY) return html_('⛔', 'אין הרשאה');
    var res = (a === 'approve') ? approve_(e.parameter.id) : reject_(e.parameter.id);
    return html_(res.ok ? (a === 'approve' ? '✅' : '🚫') : '⚠️', res.msg);
  }
  if (a === 'confirm') {
    var c = confirmAttend_(e.parameter.id);
    return html_(c.ok ? '🎉' : '⚠️', c.msg, true);
  }
  return json_({ ok: true, service: 'aviel-booking' });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  var a = body.action;
  if (a === 'book')   return json_(book_(body));
  if (a === 'status') return json_(statusOf_(body.ids || []));
  if (a === 'cancelMine') return json_(cancelMine_(body));
  if (a === 'confirmMine') return json_(confirmAttend_(body.id));
  if (a === 'reschedule') return json_(reschedule_(body));
  if (a === 'admin') {
    if (body.key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });
    ensureTrigger_();
    if (body.op === 'list')    return json_(adminList_());
    if (body.op === 'approve') return json_(approve_(body.id));
    if (body.op === 'reject')  return json_(reject_(body.id));
    if (body.op === 'cancel')  return json_(cancel_(body.id, true));
    if (body.op === 'markDone')   return json_(markStatus_(body.id, 'completed'));
    if (body.op === 'markNoshow') return json_(markStatus_(body.id, 'noshow'));
    if (body.op === 'setAvail')      return json_(setAvail_(body));
    if (body.op === 'applyTemplate') return json_(applyTemplate_(body));
    if (body.op === 'saveSettings')  { saveSettings_(body.settings); return json_({ ok: true }); }
    if (body.op === 'saveSummary')   return json_(saveSummary_(body));
    if (body.op === 'listSummaries') return json_({ ok: true, summaries: getSummaries_() });
  }
  return json_({ ok: false, error: 'unknown action' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function html_(emoji, msg, clientView) {
  var link = clientView
    ? '<a href="' + APP_URL + '">לעמוד הזימונים ←</a>'
    : '<a href="' + APP_URL + '?admin=1">למערכת הניהול ←</a>';
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html lang="he" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{background:#0a0c12;color:#eef0ff;font-family:-apple-system,Arial;display:flex;align-items:center;justify-content:center;min-height:90vh;text-align:center}' +
    '.c{background:#161a26;border:1px solid #2a3050;border-radius:20px;padding:40px 30px;max-width:340px}' +
    '.e{font-size:3.5rem;margin-bottom:14px}h2{color:#00d68f;font-size:1.15rem;line-height:1.6}' +
    'a{display:inline-block;margin-top:22px;color:#00d68f;font-weight:700}</style></head>' +
    '<body><div class="c"><div class="e">' + emoji + '</div><h2>' + msg + '</h2>' +
    link + '</div></body></html>'
  );
}

/* ---------------- public actions ---------------- */
function publicState_() {
  var settings = getSettings_();
  var avail = getAvail_();
  var bookings = getBookings_();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var maxD = Utilities.formatDate(new Date(Date.now() + (settings.maxDaysAhead || 60) * 864e5), TZ, 'yyyy-MM-dd');
  var services = (settings.services || []).filter(function (s) { return s.on !== false; });
  var defDur = (services[0] && services[0].dur) || settings.slotMin || 60;

  var openAvail = {}, busy = {}, slots = {};
  Object.keys(avail).forEach(function (d) {
    if (d < today || d > maxD) return;
    openAvail[d] = avail[d];
    var blocks = busyFor_(d, bookings, settings);
    busy[d] = blocks;
    // slots תאימות לאחור: למשך השירות הראשון
    var s = slotsForDur_(d, defDur, avail, settings, bookings, blocks);
    if (s.length) slots[d] = s;
  });
  return {
    ok: true,
    avail: openAvail,
    busy: busy,
    services: services,
    slots: slots,            // legacy
    slotMin: defDur,
    bufferMin: settings.bufferMin || 0,
    minNoticeH: settings.minNoticeH || 0,
    maxDaysAhead: settings.maxDaysAhead || 60,
    phone: settings.phone
  };
}

function book_(b) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!b.name || !b.phone || !b.email || !b.date || !b.time) return { ok: false, error: 'missing fields' };
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.email)) return { ok: false, error: 'bad email' };
    var settings = getSettings_();
    // איתור השירות שנבחר (ברירת מחדל: הראשון הפעיל)
    var svcs = (settings.services || []).filter(function (s) { return s.on !== false; });
    var svc = null;
    for (var i = 0; i < svcs.length; i++) if (svcs[i].id === b.serviceId) svc = svcs[i];
    if (!svc) svc = svcs[0] || { id: 's1', name: 'אימון אישי', dur: settings.slotMin || 60 };
    var dur = Number(svc.dur) || settings.slotMin || 60;

    if (!freeAt_(b.date, b.time, dur, getAvail_(), settings, getBookings_())) return { ok: false, error: 'slot_taken' };

    var ph = String(b.phone).replace(/\D/g, '').slice(-7);
    var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var mine = getBookings_().filter(function (x) {
      return String(x.phone).replace(/\D/g, '').slice(-7) === ph && x.date >= today &&
        (x.status === 'pending' || x.status === 'approved');
    });
    if (mine.length >= 3) return { ok: false, error: 'too_many' };

    var id = 'bk' + Date.now() + Math.floor(Math.random() * 1000);
    bkSheet_().appendRow([id, b.date, b.time, b.name, b.phone, b.email, b.note || '', 'pending',
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), '', '', '', '', svc.id, svc.name, dur]);

    b.serviceName = svc.name; b.dur = dur;
    notifyAviel_(id, b);
    mailClient_({ name: b.name, phone: b.phone, email: b.email, date: b.date, time: b.time, serviceName: svc.name, dur: dur }, 'received');
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function statusOf_(ids) {
  var map = {};
  getBookings_().forEach(function (b) {
    if (ids.indexOf(b.id) !== -1) map[b.id] = { status: b.status, date: b.date, time: b.time, name: b.name, confirmed: b.confirmed === 'yes', serviceName: b.serviceName || '', dur: Number(b.dur) || 0 };
  });
  return { ok: true, bookings: map };
}

function cancelMine_(body) {
  var bk = findBooking_(body.id);
  if (!bk) return { ok: false, error: 'not found' };
  if (String(bk.phone).replace(/\D/g, '').slice(-7) !== String(body.phone || '').replace(/\D/g, '').slice(-7))
    return { ok: false, error: 'unauthorized' };
  return cancel_(body.id, false);
}

function reschedule_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var bk = findBooking_(body.id);
    if (!bk) return { ok: false, error: 'not found' };
    if (String(bk.phone).replace(/\D/g, '').slice(-7) !== String(body.phone || '').replace(/\D/g, '').slice(-7))
      return { ok: false, error: 'unauthorized' };
    if (bk.status !== 'pending' && bk.status !== 'approved') return { ok: false, error: 'bad status' };
    var settings = getSettings_();
    var dur = Number(bk.dur) || settings.slotMin || 60;
    var others = getBookings_().filter(function (x) { return x.id !== bk.id; });
    if (!freeAt_(body.date, body.time, dur, getAvail_(), settings, others)) return { ok: false, error: 'slot_taken' };
    if (bk.eventId) {
      try { CalendarApp.getDefaultCalendar().getEventById(bk.eventId).deleteEvent(); } catch (e) {}
    }
    setBookingFields_(bk._row, { date: body.date, time: body.time, status: 'pending', eventId: '', confirmed: '', r1: '', r2: '' });
    notifyAviel_(bk.id, { name: bk.name, phone: bk.phone, email: bk.email, date: body.date, time: body.time,
      serviceName: bk.serviceName, dur: bk.dur, note: (bk.note ? bk.note + ' · ' : '') + '🔄 שינוי מועד' });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- admin actions ---------------- */
function adminList_() {
  return { ok: true, bookings: getBookings_(), avail: getAvail_(), settings: getSettings_() };
}

function findBooking_(id) {
  var all = getBookings_();
  for (var i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
  return null;
}

function approve_(id) {
  var bk = findBooking_(id);
  if (!bk) return { ok: false, msg: 'הבקשה לא נמצאה' };
  if (bk.status === 'approved') return { ok: true, msg: 'הפגישה כבר אושרה — ' + bk.name + ', ' + heDate_(bk.date) + ' ' + bk.time };
  if (bk.status !== 'pending') return { ok: false, msg: 'הבקשה כבר טופלה (' + bk.status + ')' };

  var settings = getSettings_();
  var dur = Number(bk.dur) || settings.slotMin || 60;
  var start = ilDate_(bk.date, bk.time);
  var end = new Date(start.getTime() + dur * 60000);
  var svcLabel = bk.serviceName ? bk.serviceName : 'פגישה';
  var opts = {
    description: 'טלפון: ' + bk.phone + (bk.serviceName ? '\nסוג: ' + bk.serviceName : '') +
      (bk.note ? '\nהערה: ' + bk.note : '') +
      (settings.meetLink ? '\nקישור לפגישה: ' + settings.meetLink : '') + '\nנקבע דרך מערכת הזימונים',
    guests: bk.email, sendInvites: true
  };
  if (settings.meetLink) opts.location = settings.meetLink;
  var ev = CalendarApp.getDefaultCalendar().createEvent(
    '📅 ' + svcLabel + ' — ' + bk.name,
    start, end, opts
  );
  setBookingFields_(bk._row, { status: 'approved', eventId: ev.getId() });
  mailClient_(bk, 'approved');
  return { ok: true, msg: 'הפגישה אושרה ✓ ' + bk.name + ' קיבל זימון ליומן — ' + heDate_(bk.date) + ' בשעה ' + bk.time };
}

function reject_(id) {
  var bk = findBooking_(id);
  if (!bk) return { ok: false, msg: 'הבקשה לא נמצאה' };
  if (bk.status !== 'pending') return { ok: false, msg: 'הבקשה כבר טופלה (' + bk.status + ')' };
  setBookingFields_(bk._row, { status: 'rejected' });
  mailClient_(bk, 'rejected');
  return { ok: true, msg: 'הבקשה נדחתה ונשלחה הודעה ל-' + bk.name };
}

function cancel_(id, byAdmin) {
  var bk = findBooking_(id);
  if (!bk) return { ok: false, error: 'not found', msg: 'לא נמצא' };
  if (bk.eventId) {
    try { CalendarApp.getDefaultCalendar().getEventById(bk.eventId).deleteEvent(); } catch (e) {}
  }
  setBookingFields_(bk._row, { status: 'cancelled' });
  if (byAdmin) mailClient_(bk, 'cancelled');
  else MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '❌ ביטול פגישה — ' + bk.name + ' | ' + heDate_(bk.date) + ' ' + bk.time,
    htmlBody: mailShell_('המתאמן ביטל את הפגישה', '<b>' + esc_(bk.name) + '</b> ביטל את הפגישה של ' + heDate_(bk.date) + ' בשעה ' + bk.time + '.<br>השעה חזרה להיות פנויה במערכת.') });
  return { ok: true, msg: 'הפגישה בוטלה' };
}

function markStatus_(id, status) {
  var bk = findBooking_(id);
  if (!bk) return { ok: false, msg: 'לא נמצא' };
  setBookingFields_(bk._row, { status: status });
  return { ok: true, msg: status === 'completed' ? 'סומן כבוצע ✓' : 'סומן כלא-הגיע' };
}

function confirmAttend_(id) {
  var bk = findBooking_(id);
  if (!bk) return { ok: false, msg: 'הפגישה לא נמצאה' };
  if (bk.status === 'cancelled' || bk.status === 'rejected') return { ok: false, msg: 'הפגישה כבר בוטלה' };
  if (bk.confirmed === 'yes')
    return { ok: true, msg: 'כבר אישרת הגעה — נתראה ' + heDate_(bk.date) + ' בשעה ' + bk.time + ' 💪' };
  setBookingFields_(bk._row, { confirmed: 'yes' });
  var adminPhone = getSettings_().phone;
  if (adminPhone) sendSms_(adminPhone, '✅ ' + bk.name + ' אישר/ה הגעה — ' + heDate_(bk.date) + ' ' + bk.time);
  return { ok: true, msg: 'מעולה! אישרת הגעה לפגישה ב' + heDate_(bk.date) + ' בשעה ' + bk.time + ' 🤝' };
}

function setAvail_(body) {
  var map = getAvail_();
  (body.remove || []).forEach(function (d) { delete map[d]; });
  Object.keys(body.set || {}).forEach(function (d) { map[d] = body.set[d]; });
  writeAvail_(map);
  return { ok: true, avail: map };
}

function applyTemplate_(body) {
  var settings = getSettings_();
  var tpl = body.template || settings.weekTemplate;
  var ym = body.month; // 'YYYY-MM'
  var y = +ym.split('-')[0], m = +ym.split('-')[1];
  var daysInMonth = new Date(y, m, 0).getDate();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var map = getAvail_();
  for (var day = 1; day <= daysInMonth; day++) {
    var ds = ym + '-' + ('0' + day).slice(-2);
    if (ds < today) continue;
    var dow = new Date(y, m - 1, day).getDay();
    if (tpl[dow] && tpl[dow].on) map[ds] = { start: tpl[dow].start, end: tpl[dow].end };
    else delete map[ds];
  }
  writeAvail_(map);
  return { ok: true, avail: map };
}

/* ---------------- meeting summaries ---------------- */
function saveSummary_(body) {
  var s = body.summary || {};
  if (!s.bookingId) return { ok: false, error: 'missing bookingId' };
  var all = getSummaries_();
  var existing = null;
  for (var i = 0; i < all.length; i++) if (all[i].bookingId === s.bookingId) existing = all[i];
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  var id = existing ? existing.id : 'sm' + Date.now();
  var sentAt = existing ? existing.sentAt : '';
  if (body.send && s.clientEmail) {
    mailSummary_(s);
    sentAt = now;
  }
  var rowVals = [id, s.bookingId, s.date || '', s.clientName || '', s.clientEmail || '',
    s.measurements || '', s.homeworkClient || '', s.homeworkTrainer || '', s.changes || '',
    s.requests || '', s.updates || '', existing ? existing.createdAt : now, sentAt];
  if (existing) sumSheet_().getRange(existing._row, 1, 1, SUM_HEAD.length).setValues([rowVals]);
  else sumSheet_().appendRow(rowVals);
  return { ok: true, sent: !!(body.send && s.clientEmail) };
}

function mailSummary_(s) {
  function sec(emoji, title, txt) {
    if (!txt) return '';
    return '<div style="margin-top:16px"><div style="color:#00d68f;font-weight:bold;margin-bottom:4px">' +
      emoji + ' ' + title + '</div><div style="white-space:pre-line">' + esc_(txt) + '</div></div>';
  }
  var inner = 'היי ' + esc_(s.clientName) + ',<br>הנה סיכום הפגישה שלנו מ' + heDate_(s.date) + ':' +
    sec('📏', 'היקפים ומדדים', s.measurements) +
    sec('🏠', 'שיעורי בית שלך', s.homeworkClient) +
    sec('🤝', 'מה אני לוקח על עצמי', s.homeworkTrainer) +
    sec('🔁', 'שינויים בתוכנית', s.changes) +
    sec('🧪', 'בקשות ומשימות — בדיקות / תוספים', s.requests) +
    sec('📌', 'עדכונים נוספים', s.updates) +
    '<div style="margin-top:18px">נתראה בפגישה הבאה 💪<br><a href="' + APP_URL + '" style="color:#00d68f;font-weight:bold">לקביעת הפגישה הבאה ←</a></div>';
  try {
    MailApp.sendEmail({ to: s.clientEmail, subject: '📋 סיכום הפגישה שלנו — ' + heDate_(s.date),
      htmlBody: mailShell_('סיכום פגישה', inner) });
  } catch (e) {}
}

/* ---------------- emails ---------------- */
function esc_(s) {
  return String(s || '').replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function mailShell_(title, inner, buttons) {
  return '<div dir="rtl" style="background:#0a0c12;padding:30px 16px;font-family:Arial,sans-serif">' +
    '<div style="max-width:480px;margin:0 auto;background:#161a26;border:1px solid #2a3050;border-radius:18px;padding:28px">' +
    '<div style="font-size:1.6rem;margin-bottom:4px">🏋️</div>' +
    '<h2 style="color:#00d68f;margin:0 0 14px;font-size:1.2rem">' + title + '</h2>' +
    '<div style="color:#eef0ff;font-size:.95rem;line-height:1.8">' + inner + '</div>' +
    (buttons || '') +
    '<div style="color:#8892b0;font-size:.75rem;margin-top:22px;border-top:1px solid #2a3050;padding-top:12px">מערכת הזימונים של אביאל</div>' +
    '</div></div>';
}

function notifyAviel_(id, b) {
  var base = ScriptApp.getService().getUrl();
  var ok = base + '?action=approve&id=' + id + '&key=' + ADMIN_KEY;
  var no = base + '?action=reject&id=' + id + '&key=' + ADMIN_KEY;
  var btns =
    '<div style="margin-top:24px">' +
    '<a href="' + ok + '" style="display:inline-block;background:#00d68f;color:#06281c;font-weight:bold;padding:13px 30px;border-radius:10px;text-decoration:none;margin-left:10px">✓ אשר את הפגישה</a>' +
    '<a href="' + no + '" style="display:inline-block;background:#2a3050;color:#e45858;font-weight:bold;padding:13px 30px;border-radius:10px;text-decoration:none">✗ דחה</a>' +
    '</div>';
  var inner =
    '<table dir="rtl" style="color:#eef0ff;font-size:.95rem;line-height:2">' +
    '<tr><td style="color:#8892b0;padding-left:14px">מתאמן</td><td><b>' + esc_(b.name) + '</b></td></tr>' +
    (b.serviceName ? '<tr><td style="color:#8892b0">סוג</td><td><b>' + esc_(b.serviceName) + (b.dur ? ' · ' + b.dur + ' דק\'' : '') + '</b></td></tr>' : '') +
    '<tr><td style="color:#8892b0">מועד</td><td><b>' + heDate_(b.date) + ' · ' + b.time + '</b></td></tr>' +
    '<tr><td style="color:#8892b0">טלפון</td><td><a href="tel:' + esc_(b.phone) + '" style="color:#00d68f">' + esc_(b.phone) + '</a></td></tr>' +
    '<tr><td style="color:#8892b0">מייל</td><td>' + esc_(b.email) + '</td></tr>' +
    (b.note ? '<tr><td style="color:#8892b0">הערה</td><td>' + esc_(b.note) + '</td></tr>' : '') +
    '</table>' +
    '<div style="color:#8892b0;font-size:.8rem;margin-top:10px">באישור — נוצר אירוע ביומן שלך והמתאמן מקבל זימון אוטומטי למייל.</div>';
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: '🔔 בקשת פגישה חדשה — ' + b.name + ' | ' + heDate_(b.date) + ' ' + b.time,
    htmlBody: mailShell_('בקשת פגישה חדשה ממתינה לאישור שלך', inner, btns)
  });
  var adminPhone = getSettings_().phone;
  if (adminPhone) sendSms_(adminPhone, '🔔 בקשה חדשה: ' + b.name + ' | ' + heDate_(b.date) + ' ' + b.time + '. לאישור: ' + APP_URL + '?admin=1');
}

function dateBox_(bk, settings) {
  var dur = Number(bk.dur) || settings.slotMin || 60;
  return '<div style="background:rgba(0,214,143,.12);border:2px solid #00d68f;border-radius:14px;padding:16px 20px;text-align:center;margin:16px 0">' +
    (bk.serviceName ? '<div style="color:#eef0ff;font-size:.95rem;font-weight:bold;margin-bottom:6px">' + esc_(bk.serviceName) + '</div>' : '') +
    '<div style="color:#00d68f;font-size:.74rem;font-weight:bold;margin-bottom:5px;letter-spacing:.04em">מועד הפגישה</div>' +
    '<div style="color:#eef0ff;font-size:1.15rem;font-weight:bold">' + esc_(heDate_(bk.date)) + '</div>' +
    '<div style="color:#00d68f;font-size:1.05rem;font-weight:bold;margin-top:3px">' + bk.time + ' · ' + dur + ' דקות</div>' +
    '</div>';
}

function mailClient_(bk, kind) {
  var settings = getSettings_();
  var subj, title, inner, smsText;
  var dateBox = dateBox_(bk, settings);

  if (kind === 'received') {
    subj  = '📩 קיבלנו את הבקשה שלך — ' + heDate_(bk.date) + ' ' + bk.time;
    title = 'הבקשה אצל אביאל 🎉';
    inner = 'היי <b>' + esc_(bk.name) + '</b>, קיבלנו את הבקשה שלך!' + dateBox +
      '<div style="color:#8892b0;font-size:.85rem;line-height:1.9">' +
      '① אביאל יאשר בקרוב<br>' +
      '② ברגע האישור — זימון ליומן גוגל אצלך במייל<br>' +
      '③ לשינוי מועד — <a href="' + APP_URL + '" style="color:#00d68f">לחצ/י כאן</a>' +
      '</div><br>לא צריך לעשות כלום בינתיים 💪';
    smsText = 'היי ' + bk.name + ' 👋 קיבלנו את הבקשה שלך לפגישה ב' + heDate_(bk.date) + ' בשעה ' + bk.time + '. ברגע שאביאל יאשר תקבל/י עדכון. 💪';
  } else if (kind === 'approved') {
    subj  = '✅ הפגישה אושרה — ' + heDate_(bk.date) + ' ' + bk.time;
    title = 'הפגישה אושרה! 🤝';
    var lnk = settings.meetLink;
    inner = 'היי <b>' + esc_(bk.name) + '</b>, אביאל אישר את הפגישה 🎉' + dateBox +
      '<div style="color:#8892b0;font-size:.85rem;line-height:1.9;margin-bottom:14px">' +
      '📧 זימון ליומן גוגל נשלח אליך בנפרד — אשר/י אותו כדי שיופיע ביומן.' +
      '</div>' +
      (lnk ? '<div style="text-align:center;margin-bottom:16px"><a href="' + lnk + '" style="display:inline-block;background:#00d68f;color:#06281c;font-weight:bold;padding:13px 28px;border-radius:10px;text-decoration:none;font-size:1rem">🎥 הצטרפות לפגישה</a></div>' : '') +
      '<div style="color:#8892b0;font-size:.78rem">לשינוי/ביטול: <a href="' + APP_URL + '" style="color:#00d68f">כאן ←</a></div>';
    smsText = '✅ הפגישה אושרה! ' + heDate_(bk.date) + ' בשעה ' + bk.time + '.' + (lnk ? ' קישור: ' + lnk : ' זימון נשלח למייל 📧');
  } else if (kind === 'rejected') {
    subj  = 'לגבי בקשת הפגישה שלך';
    title = 'השעה לא מסתדרת הפעם';
    inner = 'היי <b>' + esc_(bk.name) + '</b>,<br><br>' +
      'השעה שביקשת — ' + heDate_(bk.date) + ' · ' + bk.time + ' — לא מסתדרת הפעם.<br><br>' +
      '<div style="text-align:center"><a href="' + APP_URL + '" style="display:inline-block;background:#00d68f;color:#06281c;font-weight:bold;padding:13px 28px;border-radius:10px;text-decoration:none">בחר/י שעה אחרת ←</a></div>';
    smsText = 'שלום ' + bk.name + ', בקשת הפגישה ל' + heDate_(bk.date) + ' לא אושרה. לקביעה חדשה: ' + APP_URL;
  } else {
    subj  = 'הפגישה בוטלה — ' + heDate_(bk.date) + ' ' + bk.time;
    title = 'הפגישה בוטלה';
    inner = 'היי <b>' + esc_(bk.name) + '</b>,<br><br>' +
      'הפגישה של ' + heDate_(bk.date) + ' בשעה ' + bk.time + ' בוטלה.<br><br>' +
      '<div style="text-align:center"><a href="' + APP_URL + '" style="display:inline-block;background:#00d68f;color:#06281c;font-weight:bold;padding:13px 28px;border-radius:10px;text-decoration:none">קביעת מועד חדש ←</a></div>';
    smsText = 'הפגישה ב' + heDate_(bk.date) + ' בשעה ' + bk.time + ' בוטלה. לקביעה חדשה: ' + APP_URL;
  }
  try { MailApp.sendEmail({ to: bk.email, subject: subj, htmlBody: mailShell_(title, inner) }); } catch (e) {}
  if (smsText) sendSms_(bk.phone, smsText);
}

/* ---------------- reminders ---------------- */
/* רץ כל שעה — שולח תזכורת ראשונה (מייל+SMS) ותזכורת אחרונה (SMS) לפי ההגדרות */
function sendReminders() {
  var settings = getSettings_();
  if (!settings.remind1On && !settings.remind2On) return;
  var base = ScriptApp.getService().getUrl();
  var now = Date.now();
  getBookings_().forEach(function (b) {
    if (b.status !== 'approved') return;
    var hrs = (ilDate_(b.date, b.time).getTime() - now) / 3600000;
    if (hrs <= 0) return;
    if (settings.remind1On && !b.r1 && hrs <= settings.remind1H && hrs > settings.remind2H) {
      sendReminderMsg_(b, settings, base, 1);
      setBookingFields_(b._row, { r1: 'sent' });
    } else if (settings.remind2On && !b.r2 && hrs <= settings.remind2H) {
      sendReminderMsg_(b, settings, base, 2);
      setBookingFields_(b._row, { r2: 'sent' });
    }
  });
}

function sendReminderMsg_(b, settings, base, which) {
  var confirmUrl = base + '?action=confirm&id=' + b.id;
  var when = heDate_(b.date) + ' בשעה ' + b.time;
  if (which === 1) {
    var inner = 'היי <b>' + esc_(b.name) + '</b>, תזכורת ידידותית 👋' + dateBox_(b, settings) +
      (settings.reqConfirm ? '<div style="text-align:center;margin-bottom:14px"><a href="' + confirmUrl + '" style="display:inline-block;background:#00d68f;color:#06281c;font-weight:bold;padding:13px 30px;border-radius:10px;text-decoration:none;font-size:1rem">✅ אני מאשר/ת הגעה</a></div>' : '') +
      (settings.meetLink ? '<div style="text-align:center;margin-bottom:14px"><a href="' + settings.meetLink + '" style="color:#00d68f;font-weight:bold">🎥 קישור להצטרפות לפגישה</a></div>' : '') +
      '<div style="color:#8892b0;font-size:.8rem;text-align:center">צריך/ה לשנות? <a href="' + APP_URL + '" style="color:#00d68f">שינוי או ביטול כאן ←</a></div>';
    try { MailApp.sendEmail({ to: b.email, subject: '⏰ תזכורת לפגישה — ' + when, htmlBody: mailShell_('תזכורת: פגישה מתקרבת ⏰', inner) }); } catch (e) {}
    sendSms_(b.phone, '⏰ תזכורת: פגישה עם אביאל ' + when + '.' + (settings.reqConfirm ? ' לאישור הגעה: ' + confirmUrl : ' 💪'));
  } else {
    sendSms_(b.phone, '⏰ עוד מעט נתראה! הפגישה היום בשעה ' + b.time + '.' + (settings.meetLink ? ' קישור: ' + settings.meetLink : ' 💪'));
  }
}

/* ---------------- daily digest ---------------- */
function ensureTrigger_() {
  var fns = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (fns.indexOf('dailyDigest') === -1)
    ScriptApp.newTrigger('dailyDigest').timeBased().atHour(6).everyDays(1).inTimezone(TZ).create();
  if (fns.indexOf('sendReminders') === -1)
    ScriptApp.newTrigger('sendReminders').timeBased().everyHours(1).create();
}
function dailyDigest() {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var list = getBookings_().filter(function (b) { return b.date === today && b.status === 'approved'; })
    .sort(function (a, b) { return a.time < b.time ? -1 : 1; });
  if (!list.length) return;
  var inner = '<b>' + list.length + ' פגישות היום:</b><br><br>' + list.map(function (b) {
    return '🕐 <b>' + b.time + '</b> — ' + esc_(b.name) + ' · <a href="tel:' + esc_(b.phone) + '" style="color:#00d68f">' + esc_(b.phone) + '</a>' + (b.note ? '<br><span style="color:#8892b0;font-size:.85rem">📝 ' + esc_(b.note) + '</span>' : '');
  }).join('<br><br>');
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '☀️ הלוז שלך להיום — ' + list.length + ' פגישות | ' + heDate_(today),
    htmlBody: mailShell_('בוקר טוב אלוף, זה הלוז של היום', inner) });
}
