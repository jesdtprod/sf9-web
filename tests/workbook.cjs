const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const X = require('../vendor/xlsx.full.min.js');
const E = require('../js/core.js');
const context = { window: {} };
vm.runInNewContext(fs.readFileSync(require.resolve('../js/config.js'), 'utf8'), context);
const cfg = context.window.ECARD_CONFIG;
const sheet = rows => X.utils.aoa_to_sheet(rows);
const base = () => ({ Sheets: {
  Students: sheet([[], [], [], [null, 'TEST-1', 'TEST ONLY', '09/11/2009', 'FEMALE', '2026-2027', 'K to 12', 'GRADE 11', 'TEST SECTION']]),
  'CARD FRONT': sheet([[], [], ['TEST ONLY ADVISER', null, 'TEST ONLY PRINCIPAL'], ['Class Adviser', null, 'School Principal']]),
  'Core Values': sheet([[], [], [], [null, 'TEST-1', 'TEST ONLY', 'SO', 'AO', 'RO']]),
  Attendance: sheet([[null, null, null, 'JUNE'], [], [], [null, 'TEST-1', 'TEST ONLY', 20, 19, 1]])
} });
const wb = base();
wb.Sheets.Core_Subjects = sheet([[null, null, null, null, 'Core Test'], [], [null, 1, 'TEST-1', 'TEST ONLY', 90, 0, 80]]);
wb.Sheets.Elective_Subjects = sheet([[null, null, null, null, 'Elective One', 'Elective Two', null, 'Elective Three', 'Elective Four', null, 'Elective Five'], [], [null, 1, 'TEST-1', 'TEST ONLY', 86, 87, null, 88, 89, null, 91]]);
wb.Sheets.Other_Subjects = sheet([[null, null, null, null, 'Conduct', null, null, null, 'Average'], [], [null, 1, 'TEST-1', 'TEST ONLY', 74, 75, 76, null, 85, 87, 89]]);
let model = E.parse(X, wb, cfg), card = E.buildCard(model, model.students[0], cfg);
assert.equal(model.mode, 'terms-g11');
assert.equal(card.student.sex, 'FEMALE');
assert.equal(card.adviser, 'TEST ONLY ADVISER');
assert.equal(card.principal, 'TEST ONLY PRINCIPAL');
assert.deepEqual(card.terms.subjects[0].grades, [90, null, 80]);
assert.equal(card.terms.subjects[0].final, 85);
assert.deepEqual(card.terms.subjects.slice(1, 6).map(s => s.grades), [[86, null, null], [87, null, null], [null, 88, null], [null, 89, null], [null, null, 91]]);
assert.deepEqual(card.terms.averages, [85, 87, 89, 87]);
assert.equal(card.terms.subjects.at(-1).remark, 'PASSED');
assert.deepEqual(card.values[0].q, ['SO', 'AO', 'RO']);
assert.equal(card.months[0].absent, 1);
const missing = E.buildCard(model, { name: 'MISSING', lrn: 'MISSING' }, cfg);
assert.ok(missing.terms.subjects.every(s => s.final === null && s.remark === ''));
// One elective per term, in both compact and spaced workbook templates.
for (const columns of [[4, 5, 6], [4, 7, 10], [4, 5, 7]]) {
  const one = base();
  one.Sheets.Core_Subjects = wb.Sheets.Core_Subjects;
  one.Sheets.Other_Subjects = wb.Sheets.Other_Subjects;
  const headers = [], grades = [null, 1, 'TEST-1', 'TEST ONLY'];
  columns.forEach((col, index) => {
    headers[col] = 'Biology ' + (index + 1);
    grades[col] = 87 - index;
  });
  one.Sheets.Elective_Subjects = sheet([headers, [], grades]);
  const parsed = E.parse(X, one, cfg);
  const electives = E.buildCard(parsed, parsed.students[0], cfg).terms.subjects.filter(s => s.group === 'Elective Subjects');
  assert.equal(electives.length, 3);
  assert.deepEqual(electives.map(s => s.grades), [[87, null, null], [null, 86, null], [null, null, 85]]);
  assert.deepEqual(electives.map(s => s.activeTerms), [[true, false, false], [false, true, false], [false, false, true]]);
  assert.deepEqual(electives.map(s => s.final), [87, 86, 85]);
}
const terms = base();
for (const [i, name] of ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL'].entries()) {
  terms.Sheets[name] = sheet([[null, null, null, null, 'Math', 'Average', 'Remarks'], [], [null, 1, 'TEST-1', 'TEST ONLY', 80 + i, 80 + i, 'PASSED']]);
}
model = E.parse(X, terms, cfg);
assert.equal(model.mode, 'terms');
assert.deepEqual(E.buildCard(model, model.students[0], cfg).terms.subjects[0].grades, [80, 81, 82]);
const quarters = base();
for (const name of ['1st Sem', '2nd Sem']) quarters.Sheets[name] = sheet([[null, null, null, null, null, null, 'Math'], [], [null, 1, 'TEST-1', 'TEST ONLY', null, null, 80, 90]]);
model = E.parse(X, quarters, cfg);
assert.equal(model.mode, 'quarters');
assert.equal(E.buildCard(model, model.students[0], cfg).sem1.groups[0].subjects[0].final, 85);
assert.equal(E.buildCard(model, model.students[0], cfg).principal, 'TEST ONLY PRINCIPAL');

// Optional private workbook validation: never copy its data into the repository.
if (process.argv[2]) {
  const actual = X.read(fs.readFileSync(process.argv[2]), { type: 'buffer', cellDates: true });
  model = E.parse(X, actual, cfg);
  assert.equal(model.mode, 'terms-g11');
  const back = actual.Sheets['CARD BACK'];
  let checked = 0;
  for (const student of model.students) {
    card = E.buildCard(model, student, cfg);
    for (let r = 6; r <= 61; r++) {
      const name = E.norm(back['A' + r]?.v);
      const subject = /^average$/i.test(name) ? { grades: card.terms.averages.slice(0, 3) } : card.terms.subjects.find(s => s.name === name);
      if (!subject) continue;
      for (const [i, col] of ['B', 'C', 'D'].entries()) {
        const formula = back[col + r]?.f;
        if (!formula) { assert.equal(subject.grades[i], null, name + ':' + col + r); continue; }
        const match = formula.match(/XLOOKUP\([^,]+,([^!]+)!\$D\$3:[^,]+,[^!]+!([A-Z]+)3:/);
        assert.ok(match, 'Recognize workbook lookup');
        const source = actual.Sheets[match[1]];
        const range = X.utils.decode_range(source['!ref']);
        range.s = { r: 0, c: 0 };
        const rows = X.utils.sheet_to_json(source, { header: 1, range, defval: null });
        const index = rows.findIndex(row => E.key(row[3]) === E.key(student.name));
        const value = index < 0 ? null : source[match[2] + (index + 1)]?.v;
        assert.equal(subject.grades[i], value === 0 || value == null || value === '' ? null : value);
        checked++;
      }
    }
  }
  const selected = model.students.find(s => E.key(s.name) === E.key(back.H71.v));
  card = E.buildCard(model, selected, cfg);
  for (let r = 6; r <= 61; r++) {
    const subject = card.terms.subjects.find(s => s.name === E.norm(back['A' + r]?.v));
    if (subject && back['E' + r]?.v != null) assert.ok(Math.abs(subject.final - back['E' + r].v) < 1e-8);
  }
  console.log(`Private workbook: ${model.students.length} students, ${checked} source grades matched; cached final ratings matched.`);
}
console.log('G11 mapping, blank grades, electives, averages, attendance, values, G7-10 and G12 regression checks passed.');

// Exercise loading, navigation, term masking and print handlers without a browser.
(async () => {
  const elements = new Map();
  const get = id => {
    if (id === 'toastContainer') return null;
    if (!elements.has(id)) elements.set(id, {
      value: id === 'selFormat' ? 'quarters' : '', selectedIndex: 0, hidden: false,
      dataset: {}, style: { setProperty() {} }, classList: { toggle() {} },
      clientWidth: 1200, clientHeight: 900, scrollHeight: 1800, scrollTop: 0,
      addEventListener(name, fn) { this[name] = fn; },
      setAttribute() {}, focus() {}, prepend() {}, querySelector() { return null; }, querySelectorAll() { return []; },
      set innerHTML(html) { this.html = html; if (id.startsWith('sel')) this.value = html.match(/value="([^"]*)"/)?.[1] || ''; },
      get innerHTML() { return this.html || ''; }
    });
    return elements.get(id);
  };
  let printed = 0;
  const cache = new Map();
  const indexedDB = { open() {
    const request = {};
    request.result = { transaction() {
      const transaction = { objectStore() { return {
        get(key) { const query = {}; queueMicrotask(() => { query.result = cache.get(key); query.onsuccess(); }); return query; },
        put(value, key) { cache.set(key, value); queueMicrotask(() => transaction.oncomplete()); }
      }; } };
      return transaction;
    } };
    queueMicrotask(() => request.onsuccess());
    return request;
  } };
  const browser = { ECARD_CONFIG: cfg, ECard: E, ECardDropdowns: { sync() {} },
    addEventListener() {}, print() { printed++; } };
  const names = new Map();
  const localStorage = { getItem: key => names.get(key) || null, setItem: (key, value) => names.set(key, value) };
  const sandbox = { window: browser, localStorage, document: { documentElement: { classList: { remove() {} } }, getElementById: get, addEventListener() {}, createElement() { return {}; } },
    indexedDB, XLSX: X, requestAnimationFrame() {}, setTimeout(fn) { fn(); }, ResizeObserver: class { observe() {} } };
  vm.runInNewContext(fs.readFileSync(require.resolve('../js/app.js'), 'utf8'), sandbox);
  get('chooseNdp').onclick();
  assert.equal(get('schoolWorkspace').hidden, false);
  assert.equal(get('schoolLanding').hidden, true);
  assert.equal(get('schoolFavicon').href, 'assets/ndp.png');
  const bytes = X.write({ ...wb, SheetNames: Object.keys(wb.Sheets) }, { type: 'buffer', bookType: 'xlsx' });
  await get('fileFallback').onchange({ target: { files: [{ name: 'TEST ONLY.xlsx', arrayBuffer: async () => bytes }], value: '' } });
  assert.equal(get('selFormat').value, 'terms-g11');
  assert.equal(cache.get('workbook:ndp').name, 'TEST ONLY.xlsx');
  assert.match(get('status').textContent, /loaded/i);
  assert.match(get('cards').innerHTML, /Notre Dame of Parang, Inc./);
  assert.match(get('cards').innerHTML, /Mabolo St\., Poblacion I, Parang, Maguindanao del Norte/);
  assert.match(get('cards').innerHTML, /Government Recognition No. 139 s., 1967/);
  assert.match(get('cards').innerHTML, /assets\/ndp.png/);
  assert.match(get('cards').innerHTML, /TEST ONLY PRINCIPAL/);
  assert.match(get('cards').innerHTML, /TEST ONLY ADVISER/);
  assert.match(get('cards').innerHTML, /SENIOR HIGH SCHOOL/);
  assert.match(get('cards').innerHTML, /^<div class="card-spread">/);
  assert.ok(get('cards').innerHTML.indexOf('page front') < get('cards').innerHTML.indexOf('page back'));
  assert.match(get('cards').innerHTML, /SF 9-SHS/);
  assert.match(get('cards').innerHTML, /Core Subjects/);
  assert.match(get('cards').innerHTML, /Elective Subjects/);
  assert.doesNotMatch(get('cards').innerHTML, /coming soon|NaN|undefined/);
  assert.equal(get('cardViewField').hidden, false);
  assert.equal(get('btnPrint').disabled, false);
  for (const view of ['t1', 't2', 't3', 'whole']) {
    get('selCardView').value = view;
    get('selCardView').onchange();
    if (view === 't2') assert.match(get('cards').innerHTML, /term-2-values-only/);
    if (view === 't3') assert.match(get('cards').innerHTML, /term-3-values-only/);
    get('btnPrint').onclick();
  }
  get('btnPrintAll').onclick();
  assert.equal(printed, 5);
  get('btnSignatories').onclick();
  assert.equal(get('signatureSettings').hidden, false);
  get('principalName').value = 'EDITED NDP PRINCIPAL';
  get('adviserName').value = 'EDITED NDP ADVISER';
  get('principalName').oninput();
  assert.match(get('cards').innerHTML, /EDITED NDP PRINCIPAL/);
  assert.match(get('cards').innerHTML, /EDITED NDP ADVISER/);
  get('btnPrintAll').onclick();
  assert.match(get('cards').innerHTML, /EDITED NDP PRINCIPAL/);
  get('btnCloseSignatories').onclick();
  assert.equal(get('signatureSettings').hidden, true);
  get('selFormat').value = 'quarters';
  get('selFormat').change();
  assert.equal(get('btnPrint').disabled, true);
  assert.match(get('cards').innerHTML, /Open a matching workbook/);
  assert.match(get('cards').innerHTML, /LOADED WORKBOOK/);
  assert.match(get('cards').innerHTML, /SELECTED FORMAT/);
  get('btnReturnFormat').onclick();
  assert.equal(get('selFormat').value, 'terms-g11');
  assert.equal(get('btnPrint').disabled, false);
  assert.match(get('cards').innerHTML, /card-spread/);
  get('btnChangeSchool').onclick();
  assert.equal(get('schoolWorkspace').hidden, true);
  assert.equal(get('schoolLanding').hidden, false);
  get('chooseNds').onclick();
  assert.equal(get('principalName').value, '');
  assert.equal(get('schoolFavicon').href, 'assets/nds.png');
  assert.doesNotMatch(get('cards').innerHTML, /TEST ONLY/);
  await get('fileFallback').onchange({ target: { files: [{ name: 'TEST ONLY NDS.xlsx', arrayBuffer: async () => bytes }], value: '' } });
  assert.equal(cache.get('workbook:nds').name, 'TEST ONLY NDS.xlsx');
  assert.equal(cache.get('workbook:ndp').name, 'TEST ONLY.xlsx');
  assert.match(get('cards').innerHTML, /Notre Dame of Sarmiento, Inc./);
  assert.match(get('cards').innerHTML, /Government Recognition No. 40 s., 1971/);
  assert.match(get('cards').innerHTML, /assets\/nds.png/);
  assert.doesNotMatch(get('cards').innerHTML, /Notre Dame of Parang/);
  assert.doesNotMatch(get('cards').innerHTML, /EDITED NDP/);
  get('adviserName').value = '';
  get('adviserName').oninput();
  assert.doesNotMatch(get('cards').innerHTML, /TEST ONLY ADVISER/);
  get('chooseNdp').onclick();
  assert.equal(get('principalName').value, 'EDITED NDP PRINCIPAL');
  assert.equal(get('adviserName').value, 'EDITED NDP ADVISER');
  const withoutFront = { ...wb, Sheets: { ...wb.Sheets } };
  delete withoutFront.Sheets['CARD FRONT'];
  const noFrontBytes = X.write({ ...withoutFront, SheetNames: Object.keys(withoutFront.Sheets) }, { type: 'buffer', bookType: 'xlsx' });
  await get('fileFallback').onchange({ target: { files: [{ name: 'TEST ONLY NO FRONT.xlsx', arrayBuffer: async () => noFrontBytes }], value: '' } });
  assert.match(get('cards').innerHTML, /EDITED NDP PRINCIPAL/);
  assert.match(get('cards').innerHTML, /EDITED NDP ADVISER/);
  console.log('Editable signatories, saved school isolation, clearing, printing and workbook without CARD FRONT passed.');
  cache.delete('workbook:nds');
  get('chooseNds').onclick();
  await new Promise(resolve => setImmediate(resolve));
  assert.match(get('cards').innerHTML, /workbook-welcome/);
  assert.match(get('cards').innerHTML, /Notre Dame of Sarmiento/);
  let openClicks = 0;
  get('btnOpen').click = () => openClicks++;
  get('btnEmptyOpen').onclick();
  assert.equal(openClicks, 1);
  console.log('Empty workbook state and Open Excel File action passed.');
  const termBytes = X.write({ ...terms, SheetNames: Object.keys(terms.Sheets) }, { type: 'buffer', bookType: 'xlsx' });
  await get('fileFallback').onchange({ target: { files: [{ name: 'TEST ONLY G7-G10.xlsx', arrayBuffer: async () => termBytes }], value: '' } });
  const mathRow = get('cards').innerHTML.match(/<tr[^>]*>(?:(?!<\/tr>)[\s\S])*?term-learning-area[^>]*><span class="back-text">Math<\/span>[\s\S]*?<\/tr>/);
  assert.ok(mathRow, 'G7-G10 subject row renders');
  assert.match(mathRow[0], />83<\/span>/, 'Final sheet grade is visible');
  assert.match(mathRow[0], />PASSED<\/span>/, 'Subject remark is visible');
  console.log('School selection, branding, favicon, switching and workbook signatures passed.');
  console.log('App HTML, format selection, term views, print handlers and incompatible format checks passed (no visual browser check).');
})().catch(error => { console.error(error); process.exitCode = 1; });

// Excel number formats survive separately from numeric calculation values.
const formatted = base();
for (const name of ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL']) {
  formatted.Sheets[name] = sheet([[null, null, null, null, 'MAPEH', 'Music and Arts', 'PE and Health', 'Average'], [], [null, 1, 'TEST-1', 'TEST ONLY', 90, 90.5, 91, 90.50]]);
  formatted.Sheets[name].F3.z = '0.0';
  formatted.Sheets[name].H3.z = '0.00';
}
const fm = E.parse(X, formatted, cfg);
const fc = E.buildCard(fm, fm.students[0], cfg);
assert.equal(fc.terms.subjects[1].display[0], '90.5');
assert.equal(fc.terms.subjects[2].display[3], '91');
assert.equal(fc.terms.averageDisplay[3], '90.50');
for (const name of ['Music', 'Arts', 'PE', 'P.E.', 'Health', 'Music and Arts', 'PE and Health']) assert.ok(E.isComponent(name), name);
assert.ok(!E.isComponent('MAPEH'));
assert.ok(!E.isComponent('Mathematics'));

// Remarks follow rounded final ratings, but use unrounded averages in every format.
for (const [value, finalRemark, avgRemark] of [[74.4, 'FAILED', 'FAILED'], [74.5, 'PASSED', 'FAILED'], [74.99, 'PASSED', 'FAILED'], [75, 'PASSED', 'PASSED']]) {
  const boundary = base();
  for (const name of ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL']) {
    boundary.Sheets[name] = sheet([[null, null, null, null, 'Music and Arts', 'Average', 'Remarks'], [], [null, 1, 'TEST-1', 'TEST ONLY', value, value, 'INCORRECT SOURCE REMARK']]);
  }
  let m = E.parse(X, boundary, cfg), c = E.buildCard(m, m.students[0], cfg);
  assert.equal(c.terms.subjects[0].remark, finalRemark);
  assert.equal(c.terms.remark, avgRemark);
  for (const name of ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL']) delete boundary.Sheets[name];
  for (const name of ['Core_Subjects', 'Elective_Subjects', 'Other_Subjects']) {
    boundary.Sheets[name] = sheet([[null, null, null, null, name === 'Other_Subjects' ? 'Average' : 'Test Subject'], [], [null, 1, 'TEST-1', 'TEST ONLY', value, value, value]]);
  }
  m = E.parse(X, boundary, cfg); c = E.buildCard(m, m.students[0], cfg);
  assert.equal(c.terms.subjects[0].remark, finalRemark);
  assert.equal(c.terms.remark, avgRemark);
  for (const name of ['Core_Subjects', 'Elective_Subjects', 'Other_Subjects']) delete boundary.Sheets[name];
  for (const name of ['1st Sem', '2nd Sem']) {
    boundary.Sheets[name] = sheet([[null, null, null, null, null, null, 'Test Subject'], [], [null, 1, 'TEST-1', 'TEST ONLY', null, null, value, value]]);
  }
  m = E.parse(X, boundary, cfg); c = E.buildCard(m, m.students[0], cfg);
  assert.equal(c.sem1.groups[0].subjects[0].remark, finalRemark);
  assert.equal(c.sem1.remark, avgRemark);
}
console.log('Final-rating rounding and unrounded average pass boundaries passed for all formats.');
