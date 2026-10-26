/**
 * Paydirt Growth: free audit bookings
 *
 * Paste this into Extensions > Apps Script inside the "Paydirt Audit Requests" Google Sheet.
 * Free Leak Report requests (Paydirt Watchdog homepage): adds a row to the "Leak Requests" tab,
 * emails you, and emails the owner the read-only access steps.
 *
 * Audit call bookings (older form) still work too. For every booking it does three things:
 *   1. Adds a row to this Sheet
 *   2. Puts a 30-minute Google Meet on your calendar and sends the owner an invite with the agenda
 *   3. Emails you the details
 * It also tells the website which times are open, based on your real calendar.
 *
 * Needs the Google Calendar advanced service (for the Meet link):
 * in the Apps Script editor, click + next to Services, pick Google Calendar API, click Add.
 */

// ---- Settings: change these two lines if you need to ----
var NOTIFY_EMAIL = 'steven@paydirtgrowth.com';   // where new-booking emails go
var SITE_URL = 'https://paydirtgrowth.com';
var CALENDAR_ID = 'primary';                     // 'primary' is your main Google Calendar

var TZ = 'America/Denver';
var TIMES = ['8:00 AM', '9:30 AM', '11:00 AM', '1:00 PM', '2:30 PM', '4:00 PM'];
var CALL_MINUTES = 30;
var DAYS_AHEAD = 5;          // how many weekdays to offer
var MIN_NOTICE_HOURS = 3;    // no bookings sooner than this
var HEADERS = ['Received', 'Name', 'Mobile', 'Email', 'Business', 'Trade', 'Monthly ad spend', 'Call date', 'Call time', 'Source', 'Status', 'Meet link'];

/** Run this once from the editor to create the header row and approve permissions. */
function setup() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sh.setName('Bookings');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  CalendarApp.getCalendarById(CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  MailApp.getRemainingDailyQuota();
}

function cal_() { return CalendarApp.getCalendarById(CALENDAR_ID) || CalendarApp.getDefaultCalendar(); }

function slotStart_(dateStr, time) { return Utilities.parseDate(dateStr + ' ' + time, TZ, 'yyyy-MM-dd h:mm a'); }

function isFree_(start) {
  var end = new Date(start.getTime() + CALL_MINUTES * 60000);
  return cal_().getEvents(start, end).length === 0;
}

/** GET: the website asks for open times. */
function doGet() {
  var cache = CacheService.getScriptCache(), hit = cache.get('slots');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var out = [], d = new Date(), earliest = new Date(Date.now() + MIN_NOTICE_HOURS * 3600000);
  var guard = 0;
  while (out.length < DAYS_AHEAD && guard++ < 21) {
    d = new Date(d.getTime() + 24 * 3600000);
    var dow = Number(Utilities.formatDate(d, TZ, 'u')); // 1 = Mon ... 7 = Sun
    if (dow > 5) continue;
    var dateStr = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    var dayStart = slotStart_(dateStr, '12:00 AM'), dayEnd = new Date(dayStart.getTime() + 24 * 3600000);
    var events = cal_().getEvents(dayStart, dayEnd);
    var slots = TIMES.map(function (t) {
      var s = slotStart_(dateStr, t), e = new Date(s.getTime() + CALL_MINUTES * 60000);
      var busy = s < earliest || events.some(function (ev) { return ev.getStartTime() < e && ev.getEndTime() > s; });
      return { t: t, taken: busy };
    });
    out.push({ date: dateStr, dow: Utilities.formatDate(d, TZ, 'EEE'), day: Number(Utilities.formatDate(d, TZ, 'd')), mon: Utilities.formatDate(d, TZ, 'MMM'), slots: slots });
  }
  var body = JSON.stringify({ ok: true, days: out });
  cache.put('slots', body, 120); // open times load instantly for 2 minutes; bookings still double-check the real calendar
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/** POST: the website sends a booking. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var b = JSON.parse(e.postData.contents || '{}');
    var clean = function (v, n) { return String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, n || 200); };
    if (b.type === 'leak') return leak_(b, clean);
    var name = clean(b.name, 80), phone = clean(b.phone, 40), email = clean(b.email, 120), biz = clean(b.biz, 120);
    var trade = clean(b.trade, 60), spend = clean(b.spend, 60), date = clean(b.date, 10), time = clean(b.time, 10), src = clean(b.source, 300);
    if (!name || !phone || !biz || !/^\d{4}-\d{2}-\d{2}$/.test(date) || TIMES.indexOf(time) < 0) return json_({ ok: false, error: 'missing' });
    if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'email' });
    var start = slotStart_(date, time);
    if (start < new Date(Date.now() + MIN_NOTICE_HOURS * 3600000) || !isFree_(start)) return json_({ ok: false, error: 'taken' });

    var end = new Date(start.getTime() + CALL_MINUTES * 60000);
    var first = name.split(' ')[0];
    var agenda = [
      'Hi ' + first + ', thanks for booking your free Google audit with Paydirt Growth.',
      '',
      'In 30 minutes we will show you exactly how ' + biz + ' shows up on Google, where leads are slipping away, and what to fix first. Here is the plan:',
      '',
      '1. Your Google score (5 min)',
      'A score out of 100 across your map listing, Google Business Profile, ads, and reviews. Same scorecard we used to take our own turf company from a 45 to a 92.',
      '',
      '2. Where leads are leaking (10 min)',
      'The searches you are missing, the competitors outranking you in your service area, and any ad money going to people who were never going to hire you.',
      '',
      '3. Your reply speed (5 min)',
      'How fast a new lead hears back from you today, and how that compares with the business that wins the job.',
      '',
      '4. Your top three fixes (5 min)',
      'What to fix first and what it could mean in booked jobs. You keep the scorecard either way.',
      '',
      '5. Your questions (5 min)',
      'Ask anything. No pressure and no contract talk unless you want it.',
      '',
      'If you run Google Ads, have your login handy. Read-only access lets us show you your real numbers.',
      '',
      'Join with the Google Meet link in this invite.',
      'See what we do: ' + SITE_URL,
      '',
      'Paydirt Growth',
      'From the first click to the closed job.'
    ].join('\n');

    var ev = {
      summary: 'Free Google audit · ' + biz + ' × Paydirt Growth',
      description: agenda,
      start: { dateTime: start.toISOString(), timeZone: TZ },
      end: { dateTime: end.toISOString(), timeZone: TZ },
      attendees: email ? [{ email: email, displayName: name }] : [],
      conferenceData: { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } },
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 10 }] }
    };
    var made = Calendar.Events.insert(ev, CALENDAR_ID, { conferenceDataVersion: 1, sendUpdates: 'all' });
    var meet = made.hangoutLink || '';

    SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow([
      new Date(), name, phone, email, biz, trade, spend, date, time, src, 'Booked', meet
    ]);

    var when = Utilities.formatDate(start, TZ, 'EEE, MMM d') + ' at ' + time + ' Mountain';
    MailApp.sendEmail(NOTIFY_EMAIL, 'New audit booked: ' + biz + ' · ' + when,
      name + ' booked a free audit for ' + when + '.\n\nName: ' + name + '\nMobile: ' + phone + '\nEmail: ' + (email || 'not given') +
      '\nBusiness: ' + biz + '\nTrade: ' + (trade || 'not given') + '\nMonthly ad spend: ' + (spend || 'not given') +
      '\nMeet: ' + (meet || 'see calendar') + '\nSource: ' + (src || 'direct'));
    CacheService.getScriptCache().remove('slots');
    return json_({ ok: true, when: when });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

var LEAK_HEADERS = ['Received', 'Name', 'Email', 'Business', 'Trade', 'Monthly ad spend', 'Who runs ads', 'Wants Plus', 'Source', 'Status'];

/** POST with type 'leak': a free Leak Report request from the homepage. */
function leak_(b, clean) {
  var name = clean(b.name, 80), email = clean(b.email, 120), biz = clean(b.biz, 120);
  var trade = clean(b.trade, 60), spend = clean(b.spend, 60), who = clean(b.who, 60), plus = clean(b.plus, 5), src = clean(b.source, 300);
  if (!name || !biz || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'missing' });

  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Leak Requests');
  if (!sh) {
    sh = ss.insertSheet('Leak Requests');
    sh.getRange(1, 1, 1, LEAK_HEADERS.length).setValues([LEAK_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), name, email, biz, trade, spend, who, plus, src, 'Waiting for access']);

  MailApp.sendEmail(NOTIFY_EMAIL, 'New Leak Report request: ' + biz,
    name + ' asked for a free Leak Report.\n\nName: ' + name + '\nEmail: ' + email + '\nBusiness: ' + biz +
    '\nTrade: ' + (trade || 'not given') + '\nMonthly ad spend: ' + (spend || 'not given') + '\nWho runs ads: ' + (who || 'not given') +
    '\nWants Plus: ' + (plus || 'no') + '\nSource: ' + (src || 'direct') + '\n\nWatch audits@paydirtgrowth.com for their Google Ads invitation.');

  var first = name.split(' ')[0];
  var agency = who === 'An agency or freelancer';
  var steps = [
    'Hi ' + first + ',',
    '',
    'Thanks for asking for a free Leak Report for ' + biz + '. One step left: give Paydirt Watchdog read-only access to your Google Ads account. It takes about 2 minutes.',
    '',
    '1. Sign in at ads.google.com.',
    '2. Click Admin, then Access and security.',
    '3. Click the blue + button.',
    '4. Enter audits@paydirtgrowth.com and choose Read only.',
    '5. Click Send invitation.',
    '',
    agency ? 'If your agency manages the account, you can forward this email to them and ask them to do it.' : 'If someone else manages the account, you can forward this email to them.',
    '',
    'Read only means we can look but cannot change anything. You can remove our access anytime in the same place.',
    '',
    'Your Leak Report will arrive by email within 2 business days after access is accepted.',
    '',
    'Steven Black',
    'Paydirt Watchdog, made by Paydirt Growth',
    SITE_URL
  ].join('\n');
  MailApp.sendEmail(email, 'Your free Leak Report: one last step', steps, { name: 'Paydirt Watchdog', replyTo: NOTIFY_EMAIL });
  return json_({ ok: true });
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---- Reminder for owners who asked for a Leak Report but haven't given access yet ----
// Run setupReminders() once from the editor. It checks every morning at about 9 AM Mountain.
// When an owner's invitation comes in to audits@paydirtgrowth.com, change their Status in the
// "Leak Requests" tab to "Connected" so they don't get the reminder.
var REMINDER_AFTER_HOURS = 48;

function setupReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'remindStalled') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('remindStalled').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
}

function remindStalled() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leak Requests');
  if (!sh || sh.getLastRow() < 2) return;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = function (name) { return head.indexOf(name) + 1; };
  var cRem = col('Reminder sent');
  if (!cRem) { cRem = head.length + 1; sh.getRange(1, cRem).setValue('Reminder sent').setFontWeight('bold'); }
  var cRec = col('Received'), cName = col('Name'), cEmail = col('Email'), cBiz = col('Business'), cWho = col('Who runs ads'), cStatus = col('Status');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(cRem, sh.getLastColumn())).getValues();
  var cutoff = Date.now() - REMINDER_AFTER_HOURS * 3600000;
  rows.forEach(function (r, i) {
    var received = r[cRec - 1], status = String(r[cStatus - 1] || ''), sent = r[cRem - 1], email = String(r[cEmail - 1] || '');
    if (!(received instanceof Date) || received.getTime() > cutoff) return;
    if (sent || status !== 'Waiting for access' || !email) return;
    var first = String(r[cName - 1] || '').split(' ')[0] || 'there';
    var agency = r[cWho - 1] === 'An agency or freelancer';
    var body = [
      'Hi ' + first + ',',
      '',
      'Your free Leak Report for ' + (r[cBiz - 1] || 'your business') + ' is waiting on one step: read-only access to your Google Ads account.',
      '',
      '1. Sign in at ads.google.com and click Admin.',
      '2. Open Access and security, then click the blue + button.',
      '3. Enter audits@paydirtgrowth.com, choose Read only, and click Send invitation.',
      '',
      agency ? 'If your agency manages the account, forward this email and ask them to add it.' : 'If someone else manages the account, you can forward this email to them.',
      '',
      'Read only means we can look but cannot change anything. Your report arrives within 2 business days after access is accepted.',
      '',
      'Questions? Just reply.',
      '',
      'Steven Black',
      'Paydirt Watchdog, made by Paydirt Growth',
      SITE_URL
    ].join('\n');
    MailApp.sendEmail(email, 'Your Leak Report is waiting on one step', body, { name: 'Paydirt Watchdog', replyTo: NOTIFY_EMAIL });
    sh.getRange(i + 2, cRem).setValue(new Date());
  });
}
