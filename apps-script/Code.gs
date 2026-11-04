/**
 * AIESEC Egypt Finance website — Google Apps Script connector (general version).
 *
 * This connector only does three things: login, reading the sheet, writing the sheet.
 * All the finance logic lives in the website, so new features do NOT need this code to change.
 *
 * Install / update:
 *   Extensions → Apps Script → replace everything in Code.gs → Save.
 *   Deploy → Manage deployments → ✏️ → Version: New version → Deploy. (First time: Deploy → New deployment →
 *   Web app, Execute as: Me, Who has access: Anyone.)
 *
 * People log in with a username + password: sheet menu Finance website → Add or change a user.
 * Remove someone by deleting their row in the "Users" tab. The website can never read or change the "Users" tab.
 */

const USERS = 'Users';
const SESSION_DAYS = 7;
const VERSION = 2;

// ---------------------------------------------------------------- web entry
function doGet() {
  return ContentService.createTextOutput('Finance website connector is running (version ' + VERSION + ').');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'login') out = { ok: true, data: login_(req.data || {}) };
    else if (req.action === 'logout') { logout_(req.token); out = { ok: true, data: {} }; }
    else {
      const user = checkSession_(req.token);
      out = { ok: true, data: route_(req.action, req.data || {}, user) };
    }
  } catch (err) {
    out = { ok: false, error: err.message, loginNeeded: !!err.loginNeeded };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, d, user) {
  if (action === 'meta') return meta_(user);
  if (action === 'read') return read_(d.ranges || []);
  if (action === 'write') return withLock_(() => write_(d.ops || []));
  throw new Error('Unknown action: ' + action);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { const r = fn(); SpreadsheetApp.flush(); return r; } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- general sheet access
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
const tz_ = () => ss_().getSpreadsheetTimeZone();

function guard_(sheet) {
  if (!sheet) throw new Error('Tab not found');
  if (sheet.getName() === USERS) throw new Error('The Users tab cannot be used from the website');
  return sheet;
}
function sheet_(name) { return guard_(ss_().getSheetByName(name)); }
function range_(a1) { const r = ss_().getRange(a1); guard_(r.getSheet()); return r; }

// Dates travel as {d: "yyyy-MM-dd"} both ways.
function out_(v) { return v instanceof Date ? { d: Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') } : v; }
function in_(v) {
  if (v && typeof v === 'object' && typeof v.d === 'string') {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v.d);
    if (!m) throw new Error('Bad date');
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }
  return v;
}
function rowsIn_(values, row) {
  return values.map(r => r.map(v => (typeof v === 'string' && row) ? v.replace(/\{r\}/g, row) : in_(v)));
}

function meta_(user) {
  return {
    user: user, version: VERSION, today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
    sheets: ss_().getSheets().filter(s => s.getName() !== USERS).map(s => ({ name: s.getName(), rows: s.getLastRow(), cols: s.getLastColumn(), maxRows: s.getMaxRows() })),
  };
}

function read_(ranges) {
  return ranges.map(a1 => range_(a1).getValues().map(r => r.map(out_)));
}

/**
 * Runs write operations in order, all under one lock.
 *  {op:'set', range, values}                       write values / formulas (strings starting with "=")
 *  {op:'putNextRow', sheet, from, to, scanFrom, scanTo, col, values}
 *        finds the first free row after the last used row between rows from..to (looking at columns scanFrom..scanTo)
 *        and writes `values` there starting at column `col`. "{r}" inside values becomes the row number.
 *  {op:'clear', range}
 *  {op:'format', range, numberFormat, bold, background, fontSize}
 *  {op:'addSheet', name}   {op:'insertColumns', sheet, before, count}
 *  {op:'freeze', sheet, rows}   {op:'colWidth', sheet, col, px}
 */
function write_(ops) {
  return ops.map(o => {
    switch (o.op) {
      case 'set': { const r = range_(o.range); r.setValues(rowsIn_(o.values, r.getRow())); return {}; }
      case 'putNextRow': {
        const sh = sheet_(o.sheet);
        const n = o.to - o.from + 1;
        const scan = sh.getRange(o.from, o.scanFrom, n, o.scanTo - o.scanFrom + 1).getValues();
        let last = -1;
        for (let i = n - 1; i >= 0; i--) if (scan[i].some(v => v !== '' && v !== null)) { last = i; break; }
        const row = o.from + last + 1;
        if (row > o.to) throw new Error('No free row left in this part of "' + o.sheet + '". Add rows in the sheet first.');
        const vals = rowsIn_([o.values], row);
        sh.getRange(row, o.col, 1, vals[0].length).setValues(vals);
        return { row: row };
      }
      case 'clear': range_(o.range).clearContent(); return {};
      case 'format': {
        const r = range_(o.range);
        if (o.numberFormat) r.setNumberFormat(o.numberFormat);
        if (o.bold != null) r.setFontWeight(o.bold ? 'bold' : 'normal');
        if (o.background !== undefined) r.setBackground(o.background);
        if (o.fontSize) r.setFontSize(o.fontSize);
        return {};
      }
      case 'addSheet': {
        if (o.name === USERS) throw new Error('Name not allowed');
        if (ss_().getSheetByName(o.name)) throw new Error('A tab with this name already exists');
        ss_().insertSheet(o.name); return {};
      }
      case 'insertColumns': sheet_(o.sheet).insertColumnsBefore(o.before, o.count); return {};
      case 'freeze': sheet_(o.sheet).setFrozenRows(o.rows); return {};
      case 'colWidth': sheet_(o.sheet).setColumnWidth(o.col, o.px); return {};
    }
    throw new Error('Unknown write: ' + o.op);
  });
}

// ---------------------------------------------------------------- login
// "Users" tab: A Username | B Password (stored scrambled) | C Added on
function usersSheet_() {
  let sh = ss_().getSheetByName(USERS);
  if (!sh) {
    sh = ss_().insertSheet(USERS);
    sh.getRange('A1:C1').setValues([['Username', 'Password (scrambled — set it from the menu)', 'Added on']]).setFontWeight('bold').setBackground('#e8f1fd');
    sh.setFrozenRows(1); sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 380);
  }
  return sh;
}

function findUser_(username) {
  const u = String(username || '').trim().toLowerCase();
  if (!u) return null;
  const sh = usersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === u) return { row: i + 2, name: String(rows[i][0]).trim(), hash: String(rows[i][1]) };
  }
  return null;
}

function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function makeHash_(password) { const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16); return 'h$' + salt + '$' + hash_(salt, password); }
function checkHash_(stored, password) {
  const p = String(stored).split('$');
  if (p.length !== 3 || p[0] !== 'h') return false;
  return hash_(p[1], password) === p[2];
}

function loginError_(msg) { const e = new Error(msg); e.loginNeeded = true; return e; }

function login_(d) {
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + String(d.username || '').trim().toLowerCase();
  const fails = Number(cache.get(key) || 0);
  if (fails >= 5) throw loginError_('Too many wrong tries. Wait 10 minutes and try again.');
  const u = findUser_(d.username);
  if (!u || !checkHash_(u.hash, String(d.password || ''))) {
    cache.put(key, String(fails + 1), 600);
    throw loginError_('Wrong username or password');
  }
  cache.remove(key);
  cleanSessions_();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const session = { user: u.name, pw: u.hash.slice(-12), exp: Date.now() + SESSION_DAYS * 864e5 };
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify(session));
  return { token: token, user: u.name, days: SESSION_DAYS };
}

function checkSession_(token) {
  if (!token) throw loginError_('Please log in');
  const raw = PropertiesService.getScriptProperties().getProperty('s_' + token);
  if (!raw) throw loginError_('Please log in again');
  const s = JSON.parse(raw);
  const u = findUser_(s.user);
  // Logged out if: expired, user removed from the Users tab, or password changed.
  if (Date.now() > s.exp || !u || u.hash.slice(-12) !== s.pw) {
    PropertiesService.getScriptProperties().deleteProperty('s_' + token);
    throw loginError_('Please log in again');
  }
  return u.name;
}

function logout_(token) { if (token) PropertiesService.getScriptProperties().deleteProperty('s_' + token); }

function cleanSessions_() {
  const p = PropertiesService.getScriptProperties();
  const all = p.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('s_') !== 0) return;
    try { if (Date.now() > JSON.parse(all[k]).exp) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
  });
}

// ---------------------------------------------------------------- sheet menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finance website')
    .addItem('Add or change a user', 'addUser')
    .addItem('Log everyone out', 'logoutEveryone')
    .addToUi();
}

function addUser() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Add or change a user', 'Username (e.g. abdelrahman):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const name = r1.getResponseText().trim();
  if (!name) return ui.alert('Username cannot be empty.');
  const r2 = ui.prompt('Password for ' + name, 'At least 6 characters:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const pw = r2.getResponseText();
  if (pw.length < 6) return ui.alert('The password must have at least 6 characters.');
  const sh = usersSheet_();
  const u = findUser_(name);
  if (u) sh.getRange(u.row, 2).setValue(makeHash_(pw));
  else sh.appendRow([name, makeHash_(pw), new Date()]);
  ui.alert((u ? 'Password changed for ' : 'User added: ') + name + '\n\nThey can now log in to the website.');
}

function logoutEveryone() {
  const p = PropertiesService.getScriptProperties();
  Object.keys(p.getProperties()).forEach(k => { if (k.indexOf('s_') === 0) p.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('Everyone is logged out of the website.');
}
