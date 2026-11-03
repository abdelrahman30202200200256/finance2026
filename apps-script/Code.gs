/**
 * AIESEC Egypt Finance website — Google Apps Script connector.
 *
 * Paste this whole file into the "finance" Google Sheet:
 *   Extensions → Apps Script → replace everything in Code.gs → Save.
 * Then: Deploy → New deployment → type "Web app",
 *   Execute as: Me,  Who has access: Anyone  → Deploy → copy the Web app URL.
 * Reload the sheet; a menu "Finance website" appears → "Show API key".
 * Paste the URL and the key into the website once.
 */

const MAIN = '[MAIN] Cashflow';
const IS = '[NEW] Income Statement';
const EVENT_PREFIX = '[EVENT] ';

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
    if (req.key !== apiKey_()) throw new Error('Wrong API key');
    out = { ok: true, data: route_(req.action, req.data || {}) };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, d) {
  switch (action) {
    case 'init': return init_();
    case 'cashflow': return cashflow_(d.month);
    case 'addCashflow': return withLock_(() => addCashflow_(d));
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

// ---------------------------------------------------------------- API key
function apiKey_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('API_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('API_KEY', k); }
  return k;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Finance website').addItem('Show API key', 'showApiKey').addToUi();
}

function showApiKey() {
  const k = apiKey_();
  try { SpreadsheetApp.getUi().alert('API key for the finance website:\n\n' + k); } catch (e) { Logger.log(k); }
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

function addCashflow_(d) {
  if (!String(d.description || '').trim()) throw new Error('Write a description');
  const r = writeMain_(d);
  return { row: r.row };
}

// ---------------------------------------------------------------- events
// Event tab layout:
//   A1 title, A2 net summary
//   Budget (row 4 headers, rows 5+): A Type | B Item | C GFB category | D Expected (EGP) | E Actual (EGP) | F Difference
//   Event cash flow (row 4 headers, rows 5+): H Date | I Type | J Item | K GFB category | L Currency | M Amount | N EGP equivalent | O Description | P Main cash flow row
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
  const cats = categories_();
  const seen = {};
  lines.forEach(l => {
    l.item = String(l.item || '').trim();
    if (!l.item) throw new Error('Every line needs a name');
    if (l.type !== 'Revenue' && l.type !== 'Expense') throw new Error('Bad line type');
    if ((l.type === 'Revenue' ? cats.income : cats.expense).indexOf(l.gfb) < 0) throw new Error('Choose a GFB category for "' + l.item + '"');
    const key = l.type + '|' + l.item.toLowerCase();
    if (seen[key]) throw new Error('"' + l.item + '" is listed twice');
    seen[key] = true;
    l.expected = num_(l.expected);
    if (l.expected < 0) throw new Error('Expected amount cannot be negative');
  });
}

function budgetRow_(sh, r, l) {
  sh.getRange(r, 1, 1, 4).setValues([[l.type, l.item, l.gfb, l.expected]]);
  sh.getRange(r, 5).setFormula(`=SUMIFS($N$${EV_FIRST}:$N,$I$${EV_FIRST}:$I,A${r},$J$${EV_FIRST}:$J,B${r})`);
  sh.getRange(r, 6).setFormula(`=E${r}-D${r}`);
  sh.getRange(r, 4, 1, 3).setNumberFormat('#,##0.00');
}

function createEvent_(d) {
  const name = cleanName_(d.name);
  if (ss_().getSheetByName(EVENT_PREFIX + name)) throw new Error('An event with this name already exists');
  const lines = d.lines || [];
  if (!lines.length) throw new Error('Add at least one revenue or expense line');
  checkLines_(lines);

  const sh = ss_().insertSheet(EVENT_PREFIX + name);
  sh.getRange('A1').setValue('Event: ' + name).setFontSize(14).setFontWeight('bold');
  sh.getRange('A2:F2').setValues([['Net (revenue − expenses)', '', '', '', '', '']]);
  const net = col => `=SUMIF($A$${EV_FIRST}:$A,"Revenue",${col}$${EV_FIRST}:${col})-SUMIF($A$${EV_FIRST}:$A,"Expense",${col}$${EV_FIRST}:${col})`;
  sh.getRange('D2').setFormula(net('D'));
  sh.getRange('E2').setFormula(net('E'));
  sh.getRange('F2').setFormula('=E2-D2');
  sh.getRange('D2:F2').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('A3').setValue('BUDGET (EGP)').setFontWeight('bold');
  sh.getRange('H3').setValue('EVENT CASH FLOW').setFontWeight('bold');
  sh.getRange('A4:F4').setValues([['Type', 'Item', 'GFB category', 'Expected (EGP)', 'Actual (EGP)', 'Difference']]);
  sh.getRange('H4:P4').setValues([['Date', 'Type', 'Item', 'GFB category', 'Currency', 'Amount', 'EGP equivalent', 'Description', 'Main cash flow row']]);
  sh.getRange('A4:P4').setFontWeight('bold').setBackground('#e8f1fd');
  sh.getRange('G4').setBackground(null);
  const sorted = lines.filter(l => l.type === 'Revenue').concat(lines.filter(l => l.type === 'Expense'));
  sorted.forEach((l, i) => budgetRow_(sh, EV_FIRST + i, l));
  sh.setFrozenRows(4);
  sh.setColumnWidth(3, 320); sh.setColumnWidth(11, 320); sh.setColumnWidth(2, 160); sh.setColumnWidth(10, 160);
  return { name };
}

function addEventLine_(d) {
  const sh = eventSheet_(d.name);
  const ev = getEvent_(d.name);
  const l = { type: d.type, item: d.item, gfb: d.gfb, expected: d.expected };
  checkLines_(ev.lines.map(x => ({ type: x.type, item: x.item, gfb: x.gfb, expected: x.expected })).concat([l]));
  budgetRow_(sh, EV_FIRST + ev.lines.length, l);
  return {};
}

function listEvents_() {
  return ss_().getSheets().map(s => s.getName()).filter(n => n.indexOf(EVENT_PREFIX) === 0).map(n => {
    const v = ss_().getSheetByName(n).getRange('D2:E2').getValues()[0];
    return { name: n.slice(EVENT_PREFIX.length), expectedNet: num_(v[0]), actualNet: num_(v[1]) };
  });
}

function getEvent_(name) {
  const sh = eventSheet_(name);
  const last = Math.max(sh.getLastRow(), EV_FIRST);
  const budget = sh.getRange(EV_FIRST, 1, last - EV_FIRST + 1, 4).getValues().filter(r => r[1] !== '');
  const flow = sh.getRange(EV_FIRST, 8, last - EV_FIRST + 1, 9).getValues().filter(r => r[2] !== '');
  const entries = flow.map(r => ({ date: ymd_(r[0]), type: r[1], item: r[2], gfb: r[3], currency: r[4], amount: num_(r[5]), egp: num_(r[6]), description: String(r[7]), mainRow: r[8] }));
  const lines = budget.map(r => {
    const actual = entries.filter(e => e.type === r[0] && e.item === r[1]).reduce((s, e) => s + e.egp, 0);
    return { type: r[0], item: r[1], gfb: r[2], expected: num_(r[3]), actual };
  });
  return { name, lines, entries: entries.reverse() };
}

function addEventEntry_(d) {
  const sh = eventSheet_(d.name);
  const ev = getEvent_(d.name);
  const line = ev.lines.filter(l => l.type === d.type && l.item === d.item)[0];
  if (!line) throw new Error('Choose a budget line');
  const desc = String(d.description || '').trim();
  const main = writeMain_({ date: d.date, type: d.type, currency: d.currency, amount: d.amount, category: line.gfb,
    description: d.name + ' | ' + line.item + (desc ? ' | ' + desc : ''), note: 'Event: ' + d.name });
  const rate = main.rates[d.currency];
  if (!rate) throw new Error('No exchange rate for ' + d.currency + ' in the sheet');
  const row = EV_FIRST + ev.entries.length;
  sh.getRange(row, 8, 1, 9).setValues([[parseDate_(d.date), d.type, line.item, line.gfb, d.currency, num_(d.amount), num_(d.amount) * rate, desc, main.row]]);
  sh.getRange(row, 13, 1, 2).setNumberFormat('#,##0.00');
  return { mainRow: main.row };
}
