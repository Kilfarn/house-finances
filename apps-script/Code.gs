/**
 * House Finances phone app.
 * Reads and writes the "House Finances" Google Sheet. Nothing is stored anywhere else.
 * Deploy as a web app: Execute as "User accessing the web app", access "Anyone with Google account".
 * Only people the sheet is shared with can see any data.
 */

const SHEET_ID = '1ZDwiLdQjUttZy6F-SfLDiy0AIlyUMUAuLq8PuHieP3w';
const TODO_FIRST_ROW = 4;
const GIFT_FIRST_ROW = 15;
const SAVINGS_FIRST_ROW = 5;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('House Finances')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    // Lets the installable phone app (kilfarn.github.io/house-finances/) show this page full screen.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function num_(v) {
  return typeof v === 'number' && isFinite(v) ? v : null;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** Everything the app needs, in one call. */
function getAll() {
  const ss = ss_();
  return {
    month: getMonth_(ss),
    todos: getTodos_(ss),
    savings: getSavings_(ss),
    christmas: getChristmas_(ss),
    statements: getToSort_(ss),
    loadedAt: new Date().toISOString(),
  };
}

/* ---------------- This month (Monthly overview tab) ---------------- */

function getMonth_(ss) {
  const sh = ss.getSheetByName('Monthly overview');
  const rng = sh.getRange('A1:H60');
  const v = rng.getValues();
  const d = rng.getDisplayValues();

  let months = [];
  const rule = sh.getRange('C4').getDataValidation();
  if (rule) {
    const c = rule.getCriteriaValues();
    if (c && Array.isArray(c[0])) months = c[0].map(String);
  }

  const out = { label: d[3][2], months: months, headline: [], average: [], sections: [], biggest: [] };
  const first = v[3][4];
  out.ym = first instanceof Date ? Utilities.formatDate(first, ss.getSpreadsheetTimeZone(), 'yyyy-MM') : '';

  // Headline row: header row has "Money in", the next two rows are "This month" and "Average month".
  for (let r = 0; r < v.length; r++) {
    if (d[r][2] === 'Money in') {
      for (let c = 2; c <= 6; c++) {
        out.headline.push({ name: d[r][c], value: num_(v[r + 1][c]), text: d[r + 1][c] });
        out.average.push({ name: d[r][c], value: num_(v[r + 2][c]), text: d[r + 2][c] });
      }
      break;
    }
  }

  // Categories: from the "Category" header down to "Biggest purchases".
  let r = d.findIndex(row => row[1] === 'Category') + 1;
  let section = null;
  for (; r > 0 && r < d.length; r++) {
    const name = d[r][1];
    if (!name) continue;
    if (/^Biggest/.test(name)) break;
    const empty = d[r].slice(2, 8).every(x => x === '');
    if (empty && !/^Total/.test(name)) {
      section = { name: name, items: [], total: null };
      out.sections.push(section);
      continue;
    }
    if (!section) continue;
    const item = {
      name: name,
      thisMonth: num_(v[r][2]) || 0,
      average: num_(v[r][3]) || 0,
      budget: num_(v[r][5]),
      target: num_(v[r][6]),
    };
    if (/^Total/.test(name)) section.total = item;
    else section.items.push(item);
  }

  // Biggest purchases
  const start = d.findIndex((row, i) => i > 0 && row[1] === 'Date' && row[2] === 'What');
  if (start > 0) {
    for (let i = start + 1; i < d.length && d[i][1]; i++) {
      out.biggest.push({ date: d[i][1], what: d[i][2], amount: num_(v[i][3]), category: d[i][4] });
    }
  }
  return out;
}

/** Changes the month picked on the Monthly overview tab (cell C4) and returns the new figures. */
function setMonth(label) {
  const ss = ss_();
  const sh = ss.getSheetByName('Monthly overview');
  withLock_(() => {
    sh.getRange('C4').setValue(label);
    SpreadsheetApp.flush();
  });
  return getMonth_(ss);
}

/* ---------------- To-dos (To do tab) ---------------- */

function readTodoRows_(sh) {
  const last = Math.max(sh.getLastRow(), TODO_FIRST_ROW);
  const n = last - TODO_FIRST_ROW + 1;
  return { n: n, rows: sh.getRange(TODO_FIRST_ROW, 2, n, 4).getValues() };
}

function getTodos_(ss) {
  const sh = ss.getSheetByName('To do');
  const { rows } = readTodoRows_(sh);
  const out = [];
  rows.forEach((r, i) => {
    if (String(r[1]).trim() === '') return;
    out.push({ row: TODO_FIRST_ROW + i, done: r[0] === true, task: String(r[1]), who: String(r[2]), notes: String(r[3]) });
  });
  return out;
}

/** Open items stay at the top, ticked items move to the bottom. Order is otherwise kept. */
function sortTodos_(sh) {
  const { n, rows } = readTodoRows_(sh);
  const tasks = rows.filter(r => String(r[1]).trim() !== '');
  const open = tasks.filter(r => r[0] !== true);
  const done = tasks.filter(r => r[0] === true);
  const sorted = open.concat(done).map(r => [r[0] === true, r[1], r[2], r[3]]);
  while (sorted.length < n) sorted.push([false, '', '', '']);
  sh.getRange(TODO_FIRST_ROW, 2, n, 4).setValues(sorted);
}

function setTodoDone(row, task, done) {
  const ss = ss_();
  const sh = ss.getSheetByName('To do');
  withLock_(() => {
    if (String(sh.getRange(row, 3).getValue()) !== task) {
      throw new Error('The to-do list changed while you were looking at it. Pull to refresh and try again.');
    }
    sh.getRange(row, 2).setValue(!!done);
    sortTodos_(sh);
  });
  return getTodos_(ss);
}

function addTodo(task, who, notes) {
  task = String(task || '').trim();
  if (!task) throw new Error('Type a task first.');
  const ss = ss_();
  const sh = ss.getSheetByName('To do');
  withLock_(() => {
    const { rows } = readTodoRows_(sh);
    let i = rows.findIndex(r => String(r[1]).trim() === '');
    if (i < 0) i = rows.length;
    const row = TODO_FIRST_ROW + i;
    sh.getRange(row, 2, 1, 4).setValues([[false, task, who || '', notes || '']]);
    sh.getRange(row, 2).insertCheckboxes();
    sortTodos_(sh);
  });
  return getTodos_(ss);
}

/* ---------------- Savings (Savings balances tab) ---------------- */

function getSavings_(ss) {
  const sh = ss.getSheetByName('Savings balances');
  const last = sh.getLastRow();
  if (last < SAVINGS_FIRST_ROW) return [];
  const rng = sh.getRange(SAVINGS_FIRST_ROW, 1, last - SAVINGS_FIRST_ROW + 1, 5);
  const v = rng.getValues();
  const d = rng.getDisplayValues();
  const out = [];
  v.forEach((r, i) => {
    if (!d[i][1]) return;
    out.push({
      row: SAVINGS_FIRST_ROW + i,
      month: d[i][0],
      account: d[i][1],
      balance: num_(r[2]),
      paidIn: num_(r[3]),
      notes: d[i][4],
    });
  });
  return out;
}

/**
 * entries: [{row, account, balance, paidIn, paidInChanged}]
 * Balance is always written. Paid in is only written when it was changed, so the
 * default formula (planned amount from the Monthly Summary) is kept otherwise.
 */
function saveSavings(month, entries) {
  const ss = ss_();
  const sh = ss.getSheetByName('Savings balances');
  withLock_(() => {
    entries.forEach(e => {
      const check = sh.getRange(e.row, 1, 1, 2).getDisplayValues()[0];
      if (check[0] !== month || check[1] !== e.account) {
        throw new Error('The Savings balances tab changed. Pull to refresh and try again.');
      }
    });
    entries.forEach(e => {
      const bal = e.balance === '' || e.balance === null ? '' : Number(e.balance);
      sh.getRange(e.row, 3).setValue(bal);
      if (e.paidInChanged) sh.getRange(e.row, 4).setValue(e.paidIn === '' || e.paidIn === null ? 0 : Number(e.paidIn));
    });
  });
  return getSavings_(ss);
}

/* ---------------- Christmas (Christmas tab) ---------------- */

function getChristmas_(ss) {
  const sh = ss.getSheetByName('Christmas');
  const summary = sh.getRange('B5:C10').getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ name: r[0], text: r[1] }));
  const summaryVals = sh.getRange('B5:C10').getValues();
  summary.forEach((s, i) => (s.value = num_(summaryVals[i][1])));

  const groups = sh.getRange('K5:L12').getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ name: r[0], suggested: r[1] }));

  let paidWith = [];
  const rule = sh.getRange('H' + GIFT_FIRST_ROW).getDataValidation();
  if (rule) {
    const c = rule.getCriteriaValues();
    if (c && Array.isArray(c[0])) paidWith = c[0].map(String);
  }

  const last = Math.max(sh.getLastRow(), GIFT_FIRST_ROW);
  const rng = sh.getRange(GIFT_FIRST_ROW, 2, last - GIFT_FIRST_ROW + 1, 8);
  const v = rng.getValues();
  const d = rng.getDisplayValues();
  const people = [];
  v.forEach((r, i) => {
    if (!d[i][0]) return;
    people.push({
      row: GIFT_FIRST_ROW + i,
      who: d[i][0],
      group: d[i][1],
      budget: num_(r[2]),
      idea: d[i][3],
      bought: r[4] === true,
      spent: num_(r[5]),
      paidWith: d[i][6],
      notes: d[i][7],
    });
  });
  return { summary: summary, groups: groups, paidWith: paidWith, people: people };
}

/** f: {group, idea, bought, spent, paidWith, notes}. The Budget column is a formula and is left alone. */
function saveGift(row, who, f) {
  const ss = ss_();
  const sh = ss.getSheetByName('Christmas');
  withLock_(() => {
    if (sh.getRange(row, 2).getDisplayValue() !== who) {
      throw new Error('The Christmas list changed. Pull to refresh and try again.');
    }
    sh.getRange(row, 3).setValue(f.group || '');
    sh.getRange(row, 5).setValue(f.idea || '');
    sh.getRange(row, 6).setValue(!!f.bought);
    sh.getRange(row, 7).setValue(f.spent === '' || f.spent === null || f.spent === undefined ? '' : Number(f.spent));
    sh.getRange(row, 8).setValue(f.paidWith || '');
    sh.getRange(row, 9).setValue(f.notes || '');
  });
  return getChristmas_(ss);
}

function addGift(who, group, idea) {
  who = String(who || '').trim();
  if (!who) throw new Error('Type a name first.');
  const ss = ss_();
  const sh = ss.getSheetByName('Christmas');
  withLock_(() => {
    const last = Math.max(sh.getLastRow(), GIFT_FIRST_ROW);
    const names = sh.getRange(GIFT_FIRST_ROW, 2, last - GIFT_FIRST_ROW + 1, 1).getDisplayValues();
    let i = names.findIndex(r => !r[0]);
    if (i < 0) i = names.length;
    const row = GIFT_FIRST_ROW + i;
    sh.getRange(row, 2).setValue(who);
    sh.getRange(row, 3).setValue(group || '');
    sh.getRange(row, 5).setValue(idea || '');
    if (i === names.length) {
      // Past the prepared rows: copy the budget formula and checkbox down from the row above.
      sh.getRange(row - 1, 4).copyTo(sh.getRange(row, 4));
      sh.getRange(row, 6).insertCheckboxes();
    }
  });
  return getChristmas_(ss);
}

/* ---------------- Statements: upload, import and "To sort" ---------------- */

const FOLDERS = {
  barclays: { name: 'Barclays joint account', id: '1QPwkL2pi7uQZ_VCYhIaJoyqH6SMGmQqr' },
  mns: { name: 'M&S card', id: '1AZRkoHNdHT6ICF4wua_6NazT8Gd4B78i' },
  tesco: { name: 'Tesco card', id: '1OrcGLaHhcp3qwtXb2gcsxAMMzfihcaIK' },
  paypal: { name: 'PayPal', id: '1ZMZ0hNU_Odtn9868s39156yDXzKxhNnV' },
};
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const ST = { label: 0, date: 1, amount: 2, type: 3, desc: 4, ref: 5, cat: 6, what: 7, trip: 10, merchant: 11 };

function letters_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z]/g, '');
}
function isAmazon_(desc) {
  return /^(AMAZON|AMZN)/i.test(String(desc).trim());
}
function titleCase_(s) {
  return String(s).replace(/\s{2,}.*$/, '').replace(/[0-9#*_]+/g, ' ').replace(/\s+/g, ' ').trim()
    .toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase());
}
/** Barclays statements close on the 21st, so 22 Oct to 21 Nov is "Oct-Nov". */
function periodOf_(y, m, d) {
  let start = d >= 22 ? m : m - 1;
  if (start < 1) start = 12;
  const end = start === 12 ? 1 : start + 1;
  return { label: MONTHS[start - 1] + '-' + MONTHS[end - 1], start: MONTHS[start - 1], end: MONTHS[end - 1] };
}

function statementsInfo_(ss) {
  const sh = ss.getSheetByName('Statements');
  const tz = ss.getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  const v = sh.getRange(1, 1, last, 12).getValues();
  let totalRow = 0;
  for (let r = v.length - 1; r > 0; r--) if (v[r][1] === 'Total shown') { totalRow = r + 1; break; }
  let lastData = 1;
  for (let r = (totalRow || last + 1) - 2; r > 0; r--) if (v[r][1] instanceof Date) { lastData = r + 1; break; }
  const rows = [];
  for (let r = 1; r < lastData; r++) {
    const row = v[r];
    if (!(row[1] instanceof Date)) continue;
    rows.push({ row: r + 1, v: row, ymd: Utilities.formatDate(row[1], tz, 'yyyy-MM-dd') });
  }
  return { sh: sh, tz: tz, rows: rows, lastData: lastData, totalRow: totalRow };
}

function categoryList_(sh) {
  const rule = sh.getRange('G2').getDataValidation();
  if (rule) {
    const c = rule.getCriteriaValues();
    if (c && Array.isArray(c[0])) return c[0].map(String);
  }
  return [];
}

function getToSort_(ss) {
  const info = statementsInfo_(ss);
  const trips = {};
  const items = [];
  info.rows.forEach(r => {
    if (r.v[ST.trip]) trips[r.v[ST.trip]] = true;
    if (String(r.v[ST.cat]).trim() === '') {
      items.push({
        row: r.row, date: r.ymd, amount: num_(r.v[ST.amount]), desc: String(r.v[ST.desc]).trim(),
        ref: String(r.v[ST.ref]).trim(), label: String(r.v[ST.label]), amazon: isAmazon_(r.v[ST.desc]),
      });
    }
  });
  return {
    toSort: items,
    categories: categoryList_(info.sh),
    trips: Object.keys(trips).sort(),
    folders: Object.keys(FOLDERS).map(k => ({ key: k, name: FOLDERS[k].name })),
  };
}

/** Reads a Barclays CSV export: Number,Date,Account,Amount,Subcategory,Memo (memo = payee TAB reference TAB extra). */
function parseBarclays_(text) {
  const lines = String(text).replace(/\r/g, '').split('\n');
  if (!/^\s*Number,Date,Account,Amount,Subcategory,Memo/i.test(lines[0] || '')) {
    throw new Error("This doesn't look like a Barclays CSV export (the first line should start Number,Date,Account).");
  }
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].replace(/^[\t ]+/, '').replace(/[\t ]+$/, '');
    if (!line) continue;
    const parts = line.split(',');
    if (parts.length < 6) continue;
    const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(parts[1].trim());
    const amount = Number(parts[3]);
    if (!m || !isFinite(amount)) continue;
    const memo = parts.slice(5).join(',').split('\t');
    out.push({
      d: +m[1], m: +m[2], y: +m[3],
      ymd: m[3] + '-' + m[2] + '-' + m[1], dmy: parts[1].trim(),
      amount: amount, type: parts[4].trim(),
      desc: (memo[0] || '').replace(/\s+$/, ''), ref: (memo[1] || '').trim(),
      order: i,
    });
  }
  if (!out.length) throw new Error('No transactions found in this file.');
  return out;
}

function dupKey_(ymd, amount, desc) {
  return ymd + '|' + Math.round(amount * 100) + '|' + letters_(desc).slice(0, 8);
}

/** Works out what importing this CSV would do. With commit, also writes it. */
function planBarclays_(ss, text, complete) {
  const info = statementsInfo_(ss);
  const tx = parseBarclays_(text);
  tx.sort((a, b) => (a.ymd < b.ymd ? -1 : a.ymd > b.ymd ? 1 : a.order - b.order));
  const first = tx[0], last = tx[tx.length - 1];

  // Existing rows in the same date range, counted so genuine repeats (two £12 drinks) are kept.
  const have = {};
  info.rows.forEach(r => {
    if (r.ymd < first.ymd || r.ymd > last.ymd) return;
    const k = dupKey_(r.ymd, Number(r.v[ST.amount]), r.v[ST.desc]);
    have[k] = (have[k] || 0) + 1;
  });

  // Learn categories from history, newest wins. Money in and money out are learnt separately.
  const exact = {}, prefix = {};
  info.rows.forEach(r => {
    const cat = String(r.v[ST.cat]).trim();
    if (!cat || isAmazon_(r.v[ST.desc]) || r.v[ST.trip]) return;
    const dir = Number(r.v[ST.amount]) >= 0 ? '+' : '-';
    const L = letters_(r.v[ST.desc]);
    const rec = { cat: cat, merchant: String(r.v[ST.merchant] || '') };
    exact[dir + L] = rec;
    if (L.length >= 6) prefix[dir + L.slice(0, 8)] = rec;
  });

  const newRows = [];
  let skipped = 0;
  tx.forEach(t => {
    const k = dupKey_(t.ymd, t.amount, t.desc);
    if (have[k]) { have[k]--; skipped++; return; }
    const p = periodOf_(t.y, t.m, t.d);
    const dir = t.amount >= 0 ? '+' : '-';
    const L = letters_(t.desc);
    let match = null, how = 'new';
    if (!isAmazon_(t.desc)) {
      if (exact[dir + L]) { match = exact[dir + L]; how = 'known'; }
      else if (L.length >= 6 && prefix[dir + L.slice(0, 8)]) { match = prefix[dir + L.slice(0, 8)]; how = 'guessed'; }
    }
    newRows.push({
      t: t, label: p.label, how: how,
      cat: match ? match.cat : '',
      merchant: match && match.merchant ? match.merchant : (isAmazon_(t.desc) ? 'Amazon' : titleCase_(t.desc)),
    });
  });

  const lp = periodOf_(last.y, last.m, last.d);
  const fp = periodOf_(first.y, first.m, first.d);
  // A guess at whether this is the full statement: it reaches the 19th or later of the closing month.
  const looksComplete = last.d >= 19 && last.d <= 21;
  return {
    info: info, newRows: newRows, skipped: skipped,
    summary: {
      from: first.dmy, to: last.dmy, total: tx.length, skipped: skipped, adding: newRows.length,
      known: newRows.filter(r => r.how === 'known').length,
      guessed: newRows.filter(r => r.how === 'guessed').length,
      toSort: newRows.filter(r => !r.cat).length,
      period: lp.label, periods: fp.label === lp.label ? [lp.label] : [fp.label, lp.label],
      looksComplete: looksComplete,
      fileName: lp.start + ' - ' + lp.end + (complete ? '' : ' Incomplete') + '.csv',
    },
  };
}

function b64Text_(b64) {
  return Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();
}

/** Step 1 for a Barclays CSV: say what would happen, change nothing. */
function previewBarclays(b64) {
  const plan = planBarclays_(ss_(), b64Text_(b64), true);
  return plan.summary;
}

function saveFile_(kind, name, mime, b64) {
  const f = FOLDERS[kind];
  if (!f) throw new Error('Pick which account this statement is for.');
  const folder = DriveApp.getFolderById(f.id);
  let finalName = name, n = 2;
  while (folder.getFilesByName(finalName).hasNext()) {
    finalName = name.replace(/(\.[a-z0-9]+)?$/i, ' (' + n++ + ')$1');
  }
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', finalName));
  return { name: finalName, url: file.getUrl(), folder: f.name };
}

/** Card statements and other files: just file them in the right Drive folder. */
function fileStatement(kind, name, mime, b64) {
  return saveFile_(kind, name, mime, b64);
}

/** Step 2 for a Barclays CSV: file it in Drive and add the new rows to the Statements tab. */
function importBarclays(b64, complete) {
  const ss = ss_();
  const text = b64Text_(b64);
  let result;
  withLock_(() => {
    const plan = planBarclays_(ss, text, !!complete);
    const info = plan.info, sh = info.sh;
    const n = plan.newRows.length;

    if (n) {
      const values = plan.newRows.map(r => {
        const row = new Array(12).fill('');
        row[ST.label] = complete ? r.label : r.label + ' (incomplete)';
        row[ST.date] = r.t.dmy;
        row[ST.amount] = r.t.amount;
        row[ST.type] = r.t.type;
        row[ST.desc] = r.t.desc;
        row[ST.ref] = r.t.ref;
        row[ST.cat] = r.cat;
        row[ST.merchant] = r.merchant;
        return row;
      });
      sh.insertRowsAfter(info.lastData, n);
      sh.getRange(info.lastData + 1, 1, n, 12).setValues(values);
      // Keep the "Total shown" formulas covering every row.
      const newLast = info.lastData + n;
      const cells = ['J1'];
      if (info.totalRow) cells.push('C' + (info.totalRow + n), 'D' + (info.totalRow + n));
      cells.forEach(a => {
        const c = sh.getRange(a), f = c.getFormula();
        if (f) c.setFormula(f.replace(/C2:C\d+/g, 'C2:C' + newLast));
      });
    }

    // A full statement replaces the "(incomplete)" label on rows already loaded for those periods.
    let relabelled = 0;
    if (complete) {
      plan.summary.periods.forEach(p => {
        const inc = p + ' (incomplete)';
        info.rows.forEach(r => {
          if (r.v[ST.label] === inc) { sh.getRange(r.row, 1).setValue(p); relabelled++; }
        });
      });
    }
    SpreadsheetApp.flush();
    const file = saveFile_('barclays', plan.summary.fileName, 'text/csv', b64);
    result = Object.assign({}, plan.summary, {
      relabelled: relabelled, file: file,
      added: plan.newRows.map(r => ({ date: r.t.ymd, amount: r.t.amount, desc: r.t.desc, cat: r.cat, how: r.how })),
    });
  });
  result.statements = getToSort_(ss);
  return result;
}

/** Sets the category (and optionally the trip) on one Statements row. */
function setCategory(row, desc, category, trip) {
  const ss = ss_();
  const sh = ss.getSheetByName('Statements');
  withLock_(() => {
    if (String(sh.getRange(row, ST.desc + 1).getValue()).trim() !== String(desc).trim()) {
      throw new Error('The Statements tab changed. Pull to refresh and try again.');
    }
    sh.getRange(row, ST.cat + 1).setValue(category || '');
    if (trip !== undefined && trip !== null) sh.getRange(row, ST.trip + 1).setValue(trip);
  });
  return getToSort_(ss);
}

/* ---------------- Transactions behind a category ---------------- */

/**
 * Every bank and card transaction in one category for one month (yyyy-MM), from the
 * "All transactions" tab, matching how the Spending and Monthly overview tabs add them up.
 */
function getCategoryTransactions(category, ym) {
  const ss = ss_();
  const sh = ss.getSheetByName('All transactions');
  const tz = ss.getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, 9).getValues();
  const key = category + '|' + ym;
  const out = [];
  v.forEach(r => {
    if (r[8] !== key || !(r[1] instanceof Date)) return;
    out.push({
      from: String(r[0]), date: Utilities.formatDate(r[1], tz, 'yyyy-MM-dd'), amount: num_(r[2]),
      desc: String(r[3]).replace(/\s{2,}/g, ' ').trim(), what: String(r[5] || ''), trip: String(r[6] || ''),
      merchant: String(r[7] || ''),
    });
  });
  out.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  return out;
}
