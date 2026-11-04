/**
 * AIESEC Egypt Finance website — Google Apps Script connector.
 *
 * Paste this whole file into the "finance" Google Sheet:
 *   Extensions → Apps Script → replace everything in Code.gs → Save.
 * Then: Deploy → New deployment → type "Web app",
 *   Execute as: Me,  Who has access: Anyone  → Deploy → copy the Web app URL.
 *   (Updating later: Deploy → Manage deployments → ✏️ → Version: New version → Deploy. The URL stays the same.)
 *
 * People log in to the website with a username + password.
 * Add people from the sheet menu: Finance website → Add or change a user.
 * Remove someone by deleting their row in the "Users" tab.
 */

const MAIN = '[MAIN] Cashflow';
const IS = '[NEW] Income Statement';
const EVENT_PREFIX = '[EVENT] ';
const USERS = 'Users';
const SESSION_DAYS = 7;

// Columns in [MAIN] Cashflow (1 = A)
const C = { date: 4, paidEGP: 5, paidUSD: 7, paidEUR: 8, recEGP: 9, recUSD: 10, recEUR: 11,
  desc: 12, pettyBank: 13, incomeCat: 14, expenseCat: 15, receipt: 16, channel: 17, notes: 18 };

// ---------------------------------------------------------------- web entry
function doGet() {
  return ContentService.createTextOutput('Finance website connector is running.');
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
  switch (action) {
    case 'init': return Object.assign(init_(), { user: user });
    case 'cashflow': return cashflow_(d.month);
    case 'addCashflow': return withLock_(() => addCashflow_(d, user));
    case 'listEvents': return listEvents_();
    case 'createEvent': return withLock_(() => createEvent_(d));
    case 'getEvent': return getEvent_(d.name);
    case 'addEventEntry': return withLock_(() => addEventEntry_(d));
    case 'addEventLine': return withLock_(() => addEventLine_(d));
  }
  throw new Error('Unknown action: ' + action);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
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

// ---------------------------------------------------------------- helpers
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
const sheet_ = name => { const s = ss_().getSheetByName(name); if (!s) throw new Error('Tab not found: ' + name); return s; };
const tz_ = () => ss_().getSpreadsheetTimeZone();
const ymd_ = d => d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : '';
const num_ = v => (typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''))) || 0;
const parseDate_ = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); if (!m) throw new Error('Choose a date'); return new Date(+m[1], +m[2] - 1, +m[3]); };

function categories_() {
  const v = sheet_(IS).getRange(1, 3, 140, 2).getValues(); // C = code, D = name
  const income = [], expense = [];
  v.forEach(([code, name]) => {
    name = String(name).trim();
    if (!name.startsWith('MC |')) return;
    if (/-RV-/.test(code)) income.push(name); else if (/-CO-/.test(code)) expense.push(name);
  });
  return { income, expense };
}

/** Month sections of [MAIN] Cashflow: each starts with a date in column L. */
function blocks_() {
  const sh = sheet_(MAIN);
  const last = sh.getLastRow();
  const vals = sh.getRange(1, 12, last, 2).getValues(); // L, M
  const starts = [];
  vals.forEach((r, i) => { if (r[0] instanceof Date) starts.push({ row: i + 1, month: r[0].getMonth() }); });
  return starts.map((b, i) => {
    let usdRow = 0, eurRow = 0;
    for (let r = b.row; r < b.row + 8 && r <= last; r++) {
      const label = String(vals[r - 1][0]).toLowerCase();
      if (label.indexOf('usd') >= 0 && label.indexOf('rate') >= 0) usdRow = r;
      if (label.indexOf('euro') >= 0 && label.indexOf('rate') >= 0) eurRow = r;
    }
    const first = (eurRow || b.row + 4) + 1;
    const end = i + 1 < starts.length ? starts[i + 1].row - 1 : first + 145;
    return { month: b.month, first, end, usdRow, eurRow };
  });
}

function blockFor_(date) {
  const b = blocks_().filter(x => x.month === date.getMonth())[0];
  if (!b) throw new Error('No month section in ' + MAIN + ' for ' + Utilities.formatDate(date, tz_(), 'MMMM'));
  return b;
}

function rates_(b) {
  const sh = sheet_(MAIN);
  const usd = b.usdRow ? num_(sh.getRange(b.usdRow, 13).getValue()) : 0;
  const eur = b.eurRow ? num_(sh.getRange(b.eurRow, 13).getValue()) : 0;
  return { EGP: 1, USD: usd, EUR: eur };
}

/** Last row inside a month section that has anything in D..R (0 if none). */
function lastUsedRow_(b) {
  const vals = sheet_(MAIN).getRange(b.first, C.date, b.end - b.first + 1, C.notes - C.date + 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i].some(v => v !== '' && v !== null)) return b.first + i;
  return 0;
}

// ---------------------------------------------------------------- main cash flow
function init_() {
  const cats = categories_();
  return { incomeCategories: cats.income, expenseCategories: cats.expense, today: ymd_(new Date()) };
}

function cashflow_(month) {
  const date = month ? parseDate_(month + '-01') : new Date();
  const b = blockFor_(date);
  const sh = sheet_(MAIN);
  const used = lastUsedRow_(b);
  const lines = [];
  if (used) {
    const vals = sh.getRange(b.first, 1, used - b.first + 1, C.notes).getValues();
    vals.forEach((r, i) => {
      const paid = { EGP: num_(r[C.paidEGP - 1]), USD: num_(r[C.paidUSD - 1]), EUR: num_(r[C.paidEUR - 1]) };
      const rec = { EGP: num_(r[C.recEGP - 1]), USD: num_(r[C.recUSD - 1]), EUR: num_(r[C.recEUR - 1]) };
      ['EGP', 'USD', 'EUR'].forEach(cur => {
        if (paid[cur]) lines.push({ row: b.first + i, date: ymd_(r[C.date - 1]), type: 'Expense', currency: cur, amount: paid[cur], description: String(r[C.desc - 1]), category: String(r[C.expenseCat - 1]) });
        if (rec[cur]) lines.push({ row: b.first + i, date: ymd_(r[C.date - 1]), type: 'Revenue', currency: cur, amount: rec[cur], description: String(r[C.desc - 1]), category: String(r[C.incomeCat - 1]) });
      });
    });
  }
  const balRow = (used || b.first - 1) + 1;
  const bal = sh.getRange(balRow, 1, 1, 3).getValues()[0];
  return { month: Utilities.formatDate(date, tz_(), 'yyyy-MM'), balance: { EGP: num_(bal[0]), USD: num_(bal[1]), EUR: num_(bal[2]) }, lines: lines.reverse() };
}

/** Writes one line into the right month of [MAIN] Cashflow. */
function writeMain_(d) {
  const date = parseDate_(d.date);
  const amount = num_(d.amount);
  if (!(amount > 0)) throw new Error('Amount must be more than 0');
  if (['EGP', 'USD', 'EUR'].indexOf(d.currency) < 0) throw new Error('Choose a currency');
  if (d.type !== 'Revenue' && d.type !== 'Expense') throw new Error('Choose Revenue or Expense');
  const cats = categories_();
  if ((d.type === 'Revenue' ? cats.income : cats.expense).indexOf(d.category) < 0) throw new Error('Choose a category');

  const b = blockFor_(date);
  const used = lastUsedRow_(b);
  const row = used ? used + 1 : b.first;
  if (row > b.end) throw new Error('The ' + Utilities.formatDate(date, tz_(), 'MMMM') + ' section of the sheet is full. Add rows there first.');

  const sh = sheet_(MAIN);
  const colMap = d.type === 'Expense' ? { EGP: C.paidEGP, USD: C.paidUSD, EUR: C.paidEUR } : { EGP: C.recEGP, USD: C.recUSD, EUR: C.recEUR };
  sh.getRange(row, C.date).setValue(date);
  sh.getRange(row, colMap[d.currency]).setValue(amount);
  sh.getRange(row, C.desc).setValue(d.description || '');
  sh.getRange(row, d.type === 'Revenue' ? C.incomeCat : C.expenseCat).setValue(d.category);
  sh.getRange(row, C.notes).setValue(d.note || 'Added from website');
  return { row, rates: rates_(b) };
}

function addCashflow_(d, user) {
  if (!String(d.description || '').trim()) throw new Error('Write a description');
  const r = writeMain_(Object.assign({}, d, { note: 'Added from website by ' + user }));
  return { row: r.row };
}

// ---------------------------------------------------------------- events
// Each event has its OWN cash flow. Nothing is written to [MAIN] Cashflow.
// Event tab layout:
//   A1 title, row 2 net summary
//   Budget (row 4 headers, rows 5+):         A Type | B Item | C Expected (EGP) | D Actual (EGP) | E Difference
//   Event cash flow (row 4 headers, rows 5+): G Date | H Type | I Item | J Currency | K Amount | L EGP equivalent | M Description
const EV_FIRST = 5;

function eventSheet_(name) {
  const sh = ss_().getSheetByName(EVENT_PREFIX + name);
  if (!sh) throw new Error('Event not found: ' + name);
  return sh;
}

function cleanName_(name) {
  name = String(name || '').replace(/[\[\]\*\?\/\\:]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!name) throw new Error('Write the event name');
  if (name.length > 80) throw new Error('Event name is too long');
  return name;
}

function checkLines_(lines) {
  const seen = {};
  lines.forEach(l => {
    l.item = String(l.item || '').trim();
    if (!l.item) throw new Error('Every line needs a name');
    if (l.type !== 'Revenue' && l.type !== 'Expense') throw new Error('Bad line type');
    const key = l.type + '|' + l.item.toLowerCase();
    if (seen[key]) throw new Error('"' + l.item + '" is listed twice');
    seen[key] = true;
    l.expected = num_(l.expected);
    if (l.expected < 0) throw new Error('Expected amount cannot be negative');
  });
}

function budgetRow_(sh, r, l) {
  sh.getRange(r, 1, 1, 3).setValues([[l.type, l.item, l.expected]]);
  sh.getRange(r, 4).setFormula(`=SUMIFS($L$${EV_FIRST}:$L,$H$${EV_FIRST}:$H,A${r},$I$${EV_FIRST}:$I,B${r})`);
  sh.getRange(r, 5).setFormula(`=D${r}-C${r}`);
  sh.getRange(r, 3, 1, 3).setNumberFormat('#,##0.00');
}

function createEvent_(d) {
  const name = cleanName_(d.name);
  if (ss_().getSheetByName(EVENT_PREFIX + name)) throw new Error('An event with this name already exists');
  const lines = d.lines || [];
  if (!lines.length) throw new Error('Add at least one revenue or expense line');
  checkLines_(lines);

  const sh = ss_().insertSheet(EVENT_PREFIX + name);
  sh.getRange('A1').setValue('Event: ' + name).setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Net (revenue − expenses)');
  const net = col => `=SUMIF($A$${EV_FIRST}:$A,"Revenue",${col}$${EV_FIRST}:${col})-SUMIF($A$${EV_FIRST}:$A,"Expense",${col}$${EV_FIRST}:${col})`;
  sh.getRange('C2').setFormula(net('C'));
  sh.getRange('D2').setFormula(net('D'));
  sh.getRange('E2').setFormula('=D2-C2');
  sh.getRange('C2:E2').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('A3').setValue('BUDGET (EGP)').setFontWeight('bold');
  sh.getRange('G3').setValue('EVENT CASH FLOW').setFontWeight('bold');
  sh.getRange('A4:E4').setValues([['Type', 'Item', 'Expected (EGP)', 'Actual (EGP)', 'Difference']]);
  sh.getRange('G4:M4').setValues([['Date', 'Type', 'Item', 'Currency', 'Amount', 'EGP equivalent', 'Description']]);
  sh.getRange('A4:E4').setFontWeight('bold').setBackground('#e8f1fd');
  sh.getRange('G4:M4').setFontWeight('bold').setBackground('#e8f1fd');
  const sorted = lines.filter(l => l.type === 'Revenue').concat(lines.filter(l => l.type === 'Expense'));
  sorted.forEach((l, i) => budgetRow_(sh, EV_FIRST + i, l));
  sh.setFrozenRows(4);
  sh.setColumnWidth(2, 180); sh.setColumnWidth(9, 180); sh.setColumnWidth(13, 240);
  return { name };
}

function addEventLine_(d) {
  const sh = eventSheet_(d.name);
  const ev = getEvent_(d.name);
  const l = { type: d.type, item: d.item, expected: d.expected };
  checkLines_(ev.lines.map(x => ({ type: x.type, item: x.item, expected: x.expected })).concat([l]));
  budgetRow_(sh, EV_FIRST + ev.lines.length, l);
  return {};
}

function listEvents_() {
  return ss_().getSheets().map(s => s.getName()).filter(n => n.indexOf(EVENT_PREFIX) === 0).map(n => {
    const v = ss_().getSheetByName(n).getRange('C2:D2').getValues()[0];
    return { name: n.slice(EVENT_PREFIX.length), expectedNet: num_(v[0]), actualNet: num_(v[1]) };
  });
}

function getEvent_(name) {
  const sh = eventSheet_(name);
  const last = Math.max(sh.getLastRow(), EV_FIRST);
  const budget = sh.getRange(EV_FIRST, 1, last - EV_FIRST + 1, 3).getValues().filter(r => r[1] !== '');
  const flow = sh.getRange(EV_FIRST, 7, last - EV_FIRST + 1, 7).getValues().filter(r => r[2] !== '');
  const entries = flow.map(r => ({ date: ymd_(r[0]), type: r[1], item: r[2], currency: r[3], amount: num_(r[4]), egp: num_(r[5]), description: String(r[6]) }));
  const lines = budget.map(r => {
    const actual = entries.filter(e => e.type === r[0] && e.item === r[1]).reduce((s, e) => s + e.egp, 0);
    return { type: r[0], item: r[1], expected: num_(r[2]), actual };
  });
  return { name, lines, entries: entries.reverse() };
}

function addEventEntry_(d) {
  const sh = eventSheet_(d.name);
  const ev = getEvent_(d.name);
  const line = ev.lines.filter(l => l.type === d.type && l.item === d.item)[0];
  if (!line) throw new Error('Choose a budget line');
  const date = parseDate_(d.date);
  const amount = num_(d.amount);
  if (!(amount > 0)) throw new Error('Amount must be more than 0');
  if (['EGP', 'USD', 'EUR'].indexOf(d.currency) < 0) throw new Error('Choose a currency');
  // The exchange rate is only READ from that month's rates in [MAIN] Cashflow.
  const rate = d.currency === 'EGP' ? 1 : rates_(blockFor_(date))[d.currency];
  if (!rate) throw new Error('No exchange rate for ' + d.currency + ' in the sheet');
  const row = EV_FIRST + ev.entries.length;
  sh.getRange(row, 7, 1, 7).setValues([[date, d.type, line.item, d.currency, amount, amount * rate, String(d.description || '').trim()]]);
  sh.getRange(row, 11, 1, 2).setNumberFormat('#,##0.00');
  return { row };
}
