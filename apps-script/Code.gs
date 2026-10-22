/**
 * Requests Desk — backend (Google Apps Script)
 * ------------------------------------------------------------
 * A STANDALONE Apps Script project (not the one behind the request form).
 * It reads/writes the "Employee Requests" sheet and gates every call behind
 * an email-code login with a 30-day session.
 *
 * Deploy: Deploy → New deployment → Web app
 *   Execute as: Me      Who has access: Anyone
 * ("Anyone" only means the URL answers; every action except login requires
 *  a valid signed session for an email listed as Active on the Users tab.)
 */

const SHEET_ID = '13oBx6Hq5-JrvR_8Jx0Ss5rbD_rKbiIqmTvutSmsu6ls';
const SESSION_DAYS = 30;
const CODE_MINUTES = 10;
const APP_NAME = 'Requests Desk';

// Queues the dashboard can see. "Store Supplies Orders" is intentionally absent.
const TABS = {
  shipping: 'Shipping Label Requests',
  cards:    'Business Card Requests',
  badges:   'ID Badge Requests',
  shirts:   'Shirt Requests',
  brinks:   'Brinks Login Requests',
  hire:     'Hire Requests'
};
const EXTRA_COLS = ['Handled By', 'Updated At', 'Team Notes'];
const STATUS_LABEL = { open: '', working: 'In Progress', issue: 'Needs Info', done: 'Done' };

const USERS_TAB = 'Users';
const LOG_TAB = 'Activity Log';

/* ---------------- one-time setup ---------------- */
// Run once from the editor (select setup → Run). Safe to re-run.
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let u = ss.getSheetByName(USERS_TAB);
  if (!u) {
    u = ss.insertSheet(USERS_TAB);
    u.appendRow(['Email', 'Name', 'Role', 'Active', 'Added']);
    u.appendRow(['jon.sanchez@nyconnectionli.com', 'Jonathan Sanchez', 'Admin', true, new Date()]);
    u.setFrozenRows(1);
    u.getRange('A1:E1').setFontWeight('bold');
    u.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Admin', 'Helper'], true).build());
    u.getRange('D2:D').insertCheckboxes();
  }
  let l = ss.getSheetByName(LOG_TAB);
  if (!l) {
    l = ss.insertSheet(LOG_TAB);
    l.appendRow(['Time', 'User', 'Queue', 'Row', 'Request', 'Action', 'Note']);
    l.setFrozenRows(1);
    l.getRange('A1:G1').setFontWeight('bold');
  }
  Object.values(TABS).forEach(name => { const sh = ss.getSheetByName(name); if (sh) ensureCols_(sh); });
  secret_();
  return 'Setup complete';
}

/* ---------------- HTTP ---------------- */
function doGet() { return json_({ ok: true, app: APP_NAME }); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    switch (body.action) {
      case 'requestCode': return json_(requestCode_(body.email));
      case 'verifyCode':  return json_(verifyCode_(body.email, body.code));
    }
    const me = auth_(body.token);
    if (!me) return json_({ ok: false, error: 'signed_out' });
    switch (body.action) {
      case 'me':        return json_({ ok: true, me });
      case 'data':      return json_(data_(me));
      case 'update':    return json_(update_(me, body));
      case 'bulk':      return json_(bulk_(me, body));
      case 'users':     return json_(users_(me));
      case 'saveUser':  return json_(saveUser_(me, body));
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}

/* ---------------- auth ---------------- */
function normEmail_(e) { return String(e || '').trim().toLowerCase(); }

function findUser_(email) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_TAB);
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (normEmail_(v[i][0]) === email) {
      const active = v[i][3] === true || String(v[i][3]).toUpperCase() === 'TRUE';
      if (!active) return null;
      return { email, name: String(v[i][1] || email.split('@')[0]), role: String(v[i][2] || 'Helper') === 'Admin' ? 'Admin' : 'Helper' };
    }
  }
  return null;
}

function requestCode_(rawEmail) {
  const email = normEmail_(rawEmail);
  if (!/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(email)) return { ok: false, error: 'bad_email' };
  const cache = CacheService.getScriptCache();
  if (cache.get('sent:' + email)) return { ok: true }; // throttle: one email per minute
  const user = findUser_(email);
  if (user) {
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache.put('code:' + email, JSON.stringify({ code, tries: 0 }), CODE_MINUTES * 60);
    cache.put('sent:' + email, '1', 60);
    MailApp.sendEmail({
      to: email,
      subject: APP_NAME + ' sign-in code: ' + code,
      htmlBody: '<div style="font-family:-apple-system,Helvetica,Arial,sans-serif;font-size:15px;color:#1c1c1e">' +
        '<p>Hi ' + user.name.split(' ')[0] + ',</p><p>Your ' + APP_NAME + ' sign-in code is</p>' +
        '<p style="font-size:30px;font-weight:700;letter-spacing:6px;margin:12px 0">' + code + '</p>' +
        '<p style="color:#6e6e73">It expires in ' + CODE_MINUTES + ' minutes. If you didn\'t try to sign in, ignore this email.</p></div>'
    });
  }
  // Same answer either way, so the page never reveals who has access.
  return { ok: true };
}

function verifyCode_(rawEmail, rawCode) {
  const email = normEmail_(rawEmail);
  const cache = CacheService.getScriptCache();
  const k = 'code:' + email;
  const rec = cache.get(k);
  if (!rec) return { ok: false, error: 'code_expired' };
  const r = JSON.parse(rec);
  if (r.tries >= 5) { cache.remove(k); return { ok: false, error: 'too_many_tries' }; }
  if (String(rawCode || '').replace(/\D/g, '') !== r.code) {
    r.tries++; cache.put(k, JSON.stringify(r), CODE_MINUTES * 60);
    return { ok: false, error: 'wrong_code', triesLeft: 5 - r.tries };
  }
  cache.remove(k);
  const user = findUser_(email);
  if (!user) return { ok: false, error: 'no_access' };
  const exp = Date.now() + SESSION_DAYS * 864e5;
  return { ok: true, token: sign_(email + '|' + exp), expires: exp, me: user };
}

function auth_(token) {
  if (!token || token.indexOf('.') < 0) return null;
  const [p, sig] = token.split('.');
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(p)).getDataAsString(); } catch (e) { return null; }
  if (sign_(payload) !== token) return null;
  const [email, exp] = payload.split('|');
  if (!exp || Date.now() > Number(exp)) return null;
  return findUser_(email); // removing someone from Users (or unticking Active) locks them out immediately
}

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SESSION_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SESSION_SECRET', s); }
  return s;
}
function sign_(payload) {
  const p = Utilities.base64EncodeWebSafe(payload);
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_()));
  return p + '.' + sig;
}

/* ---------------- data ---------------- */
function ensureCols_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());
  const need = ['Status'].concat(EXTRA_COLS).filter(h => head.indexOf(h) < 0);
  if (need.length) {
    sh.getRange(1, lastCol + 1, 1, need.length).setValues([need]).setFontWeight('bold');
    return head.concat(need);
  }
  return head;
}

function data_(me) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tabs = {};
  Object.keys(TABS).forEach(key => {
    const sh = ss.getSheetByName(TABS[key]);
    tabs[key] = sh ? sh.getDataRange().getDisplayValues() : [];
  });
  return { ok: true, me, tabs, at: Date.now() };
}

function applyUpdate_(ss, me, b) {
  const name = TABS[b.tab];
  if (!name) return { ok: false, error: 'bad_tab' };
  if (!(b.status in STATUS_LABEL) && b.status !== undefined && b.status !== null) return { ok: false, error: 'bad_status' };
  const sh = ss.getSheetByName(name);
  const head = ensureCols_(sh);
  const row = Number(b.row);
  if (!(row >= 2 && row <= sh.getLastRow())) return { ok: false, error: 'row_moved' };
  // Make sure the row still holds the same request before writing.
  const ts = sh.getRange(row, 1).getDisplayValue();
  if (String(ts).trim() !== String(b.ts || '').trim()) return { ok: false, error: 'row_moved' };
  const col = h => head.indexOf(h) + 1;
  const sets = [];
  if (b.status !== undefined && b.status !== null) sets.push([col('Status'), STATUS_LABEL[b.status]]);
  if (typeof b.note === 'string') sets.push([col('Team Notes'), b.note.slice(0, 2000)]);
  sets.push([col('Handled By'), me.name]);
  sets.push([col('Updated At'), Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy, h:mm:ss a')]);
  sets.forEach(([c, v]) => sh.getRange(row, c).setValue(v));
  const log = ss.getSheetByName(LOG_TAB);
  if (log) log.appendRow([new Date(), me.name + ' <' + me.email + '>', name, row, String(b.title || '').slice(0, 200),
    b.status !== undefined && b.status !== null ? (STATUS_LABEL[b.status] || 'Reopened') : 'Note', typeof b.note === 'string' ? b.note.slice(0, 500) : '']);
  return { ok: true };
}

function update_(me, b) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return applyUpdate_(SpreadsheetApp.openById(SHEET_ID), me, b); }
  finally { lock.releaseLock(); }
}

function bulk_(me, b) {
  const items = (b.items || []).slice(0, 100);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const results = items.map(it => applyUpdate_(ss, me, Object.assign({}, it, { status: b.status })));
    return { ok: true, results };
  } finally { lock.releaseLock(); }
}

/* ---------------- team access (Admins) ---------------- */
function users_(me) {
  if (me.role !== 'Admin') return { ok: false, error: 'admin_only' };
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_TAB);
  const v = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  return { ok: true, users: v.map(r => ({ email: normEmail_(r[0]), name: String(r[1] || ''), role: String(r[2] || 'Helper'),
    active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE' })) };
}

function saveUser_(me, b) {
  if (me.role !== 'Admin') return { ok: false, error: 'admin_only' };
  const email = normEmail_(b.email);
  if (!/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(email)) return { ok: false, error: 'bad_email' };
  if (email === me.email && (b.active === false || b.role !== 'Admin')) return { ok: false, error: 'cant_remove_self' };
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_TAB);
  const v = sh.getDataRange().getValues();
  const row = [email, String(b.name || '').slice(0, 80), b.role === 'Admin' ? 'Admin' : 'Helper', b.active !== false];
  let firstEmpty = -1;
  for (let i = 1; i < v.length; i++) {
    if (normEmail_(v[i][0]) === email) { sh.getRange(i + 1, 1, 1, 4).setValues([row]); return users_(me); }
    if (firstEmpty < 0 && !String(v[i][0]).trim()) firstEmpty = i + 1;
  }
  // The Active column has checkboxes down the sheet, so fill the first blank row instead of appending at the bottom.
  const target = firstEmpty > 0 ? firstEmpty : sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, 5).setValues([row.concat([new Date()])]);
  return users_(me);
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
