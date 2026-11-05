/* Reads the Excel workbook and computes everything the card needs.
   No DOM in here, so it can be tested on its own. */
(function (root) {
  'use strict';

  const norm = s => (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim();   // kills the trailing-space problem
  const key = s => norm(s).toUpperCase();
  const num = v => {
    if (v == null || v === '') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).trim());
    return Number.isFinite(n) ? n : null;
  };
  const round2 = n => Math.round(n * 100) / 100;

  function sheetRows(XLSX, wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error('The Excel file has no sheet named "' + name + '".');
    const r = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    r.s = { r: 0, c: 0 };                       // absolute positions, row 1 = index 0
    return XLSX.utils.sheet_to_json(ws, { header: 1, range: r, defval: null, raw: true, blankrows: true });
  }

  function parseDate(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400000));
    const m = norm(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
    const textDate = norm(v).match(/^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})$/);
    if (!textDate) return null;
    const parsed = new Date(textDate[1] + ' ' + textDate[2] + ', ' + textDate[3]);
    return isNaN(parsed) ? null : parsed;
  }

  function readStudents(rows, mode) {
    const out = [];
    for (let r = 3; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = norm(row[2]);
      if (!name) continue;
      if (mode !== 'quarters') {
        out.push({ lrn: norm(row[1]), name, birth: parseDate(row[3]), sex: norm(row[4]),
          schoolYear: norm(row[5]), curriculum: norm(row[6]), grade: norm(row[7]), section: norm(row[8]), strand: '' });
        continue;
      }
      out.push({
        lrn: norm(row[1]), name, birth: parseDate(row[4]) || parseDate(row[3]),
        sex: norm(row[5]), schoolYear: norm(row[6]), grade: norm(row[7]),
        section: norm(row[8]), curriculum: norm(row[9]), strand: norm(row[10])
      });
    }
    return out;
  }

  function indexRows(rows, startRow, nameCol, lrnCol) {
    const byName = new Map(), byLrn = new Map();
    for (let r = startRow; r < rows.length; r++) {
      const row = rows[r] || [];
      const n = key(row[nameCol]);
      if (n && !byName.has(n)) byName.set(n, row);
      const l = norm(row[lrnCol]);
      if (l && !byLrn.has(l)) byLrn.set(l, row);
    }
    return { byName, byLrn };
  }

  function readSem(rows) {
    const subjects = [];
    // New workbooks put the category above each subject; older ones put
    // the subject directly in the first row above the quarter headings.
    const grouped = norm((rows[1] || [])[6]) && !/^(1ST|2ND|3RD|4TH|Q[1-4])\b/i.test(norm(rows[1][6]));
    for (let c = 6; c < (rows[0] || []).length; c += 4) {
      const nm = norm((rows[grouped ? 1 : 0] || [])[c]);
      if (!nm) continue;
      subjects.push({ name: nm, col: c, group: grouped ? norm(rows[0][c]) : '' });
    }
    return Object.assign({ subjects }, indexRows(rows, 2, 3, 2));
  }

  function readTermValues(wb, cfg) {
    const card = wb.Sheets['CARD BACK'], values = [];
    if (card) for (let r = 1; r <= 69; r++) {
      const title = norm(card['H' + r] && card['H' + r].v);
      const match = title.match(/^([A-Z])\s*[-?]\s*(.+)$/i);
      const statement = norm(card['I' + r] && card['I' + r].v);
      if (match && statement) values.push({ code: match[1].toUpperCase(), name: match[2], statement });
    }
    return values.length === 9 ? values : cfg.values;
  }

  function parse(XLSX, wb, cfg) {
    const S = cfg.sheets;
    const g11Sheets = ['Core_Subjects', 'Elective_Subjects', 'Other_Subjects'];
    const mode = g11Sheets.every(name => !!wb.Sheets[name]) ? 'terms-g11' :
      ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL'].every(name => !!wb.Sheets[name]) ? 'terms' : 'quarters';
    const students = readStudents(sheetRows(XLSX, wb, S.students), mode);
    const vRows = sheetRows(XLSX, wb, S.values);
    const aRows = sheetRows(XLSX, wb, S.attendance);
    const months = [];
    const periodCount = mode === 'quarters' ? 4 : 3;
    for (let c = 3; c < (aRows[0] || []).length; c += periodCount) {
      const m = norm(aRows[0][c]);
      if (!m || m.toUpperCase() === 'TOTAL') break;
      months.push({ name: m, col: c });
    }
    const shared = { mode, students, values: indexRows(vRows, 3, 2, 1),
      attendance: Object.assign({ months }, indexRows(aRows, 3, 2, 1)) };
    if (mode === 'terms-g11') {
      const gradeSheets = g11Sheets.map((name, i) => {
        const rows = sheetRows(XLSX, wb, name), headers = rows[0] || [];
        const subjects = [];
        for (let col = 4; col < headers.length; col++) {
          const name = norm(headers[col]);
          if (!name) continue;
          // Electives occupy one term only; the other sheets use three columns per subject.
          const term = i === 1 ? Math.floor((col - 4) / 3) : null;
          subjects.push({ name, cols: term == null ? [col, col + 1, col + 2] :
            [0, 1, 2].map(k => k === term ? col : null) });
        }
        if (i === 0) {
          const combined = subjects.findIndex(s => /Effective Communication\s*\//i.test(s.name));
          if (combined > 0) subjects.unshift(subjects.splice(combined, 1)[0]);
        }
        return Object.assign({ title: ['Core Subjects', 'Elective Subjects', 'Other Subjects'][i], subjects }, indexRows(rows, 2, 3, 2));
      });
      const front = wb.Sheets['CARD FRONT'];
      return Object.assign(shared, { gradeSheets, termValues: readTermValues(wb, cfg),
        adviser: norm(front && front.Q29 && front.Q29.v), principal: norm(front && front.N34 && front.N34.v) });
    }
    if (mode === 'terms') {
      const termSheets = ['TERM 1', 'TERM 2', 'TERM 3', 'FINAL'].map(name => {
        const rows = sheetRows(XLSX, wb, name);
        return Object.assign({ headers: rows[0] || [] }, indexRows(rows, 2, 3, 2));
      });
      const headers = termSheets[0].headers;
      const averageCol = headers.findIndex((header, i) => i >= 4 && /^(general\s+)?average$/i.test(norm(header)));
      const subjects = headers.slice(4, averageCol < 0 ? headers.length : averageCol).map((name, i) => ({
        name: norm(name), col: 4 + i
      })).filter(subject => subject.name && !/^rank$/i.test(subject.name));
      const front = wb.Sheets['CARD FRONT'];
      return Object.assign(shared, { termSheets, subjects, averageCol, termValues: readTermValues(wb, cfg), adviser: norm(front && front.Q29 && front.Q29.v),
        principal: norm(front && front.N34 && front.N34.v) });
    }
    return Object.assign(shared, {
      sem1: readSem(sheetRows(XLSX, wb, S.sem1)),
      sem2: readSem(sheetRows(XLSX, wb, S.sem2))
    });
  }

  const find = (idx, st) => idx.byName.get(key(st.name)) || idx.byLrn.get(st.lrn) || null;

  function buildSem(sem, st, cfg, semNo) {
    const row = find(sem, st);
    const excl = new Set(cfg.excludeFromAverage.map(key));
    const subjects = sem.subjects.map(s => {
      const q1 = row ? num(row[s.col]) : null, q2 = row ? num(row[s.col + 1]) : null;
      const ok = cfg.requireBothQuarters ? (q1 != null && q2 != null) : (q1 != null || q2 != null);
      const vals = [q1, q2].filter(v => v != null);
      const final = ok ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
      return { name: s.name, group: s.group, q1, q2, final, remark: remark(final, cfg), inAverage: !excl.has(key(s.name)) };
    });
    const fin = subjects.filter(s => s.inAverage && s.final != null).map(s => s.final);
    const genAvg = fin.length ? fin.reduce((a, b) => a + b, 0) / fin.length : null;
    // Keep the standard category order, while taking membership from the file.
    const layout = cfg.layout[semNo];
    const groups = layout.map(([title]) => ({ title, subjects: [] }));
    const categoryKey = value => key(value).replace(/[^A-Z0-9]/g, '');
    let boundary = 0, fallback = 0;
    subjects.forEach((subject, index) => {
      while (fallback < layout.length - 1 && index >= boundary + layout[fallback][1]) boundary += layout[fallback++][1];
      let group = subject.group && groups.find(g => categoryKey(g.title) === categoryKey(subject.group));
      if (!group && subject.group) {
        group = { title: subject.group, subjects: [] };
        groups.push(group);
      }
      (group || groups[fallback]).subjects.push(subject);
    });
    return { groups, genAvg, remark: remark(genAvg, cfg), found: !!row };
  }

  function remark(n, cfg) { return n == null ? '' : (n > cfg.passMark - 0.1 ? 'PASSED' : 'FAILED'); }

  function buildCard(model, st, cfg) {
    const threeTerms = model.mode !== 'quarters';
    const vrow = find(model.values, st);
    const values = (threeTerms ? model.termValues : cfg.values).map((v, i) => ({
      ...v, q: Array.from({ length: threeTerms ? 3 : 4 }, (_, k) => vrow ? norm(vrow[3 + i * (threeTerms ? 3 : 4) + k]) : '')
    }));
    const arow = find(model.attendance, st);
    const months = model.attendance.months.map(m => {
      const g = k => (arow ? num(arow[m.col + k]) : null);
      const days = g(0), present = g(1);
      return { name: m.name, days, present, absent: days == null || present == null ? null : days - present,
        tardy: g(2), mass: threeTerms ? null : g(3) };
    });
    const sum = k => { const v = months.map(m => m[k]).filter(x => x != null); return v.length ? v.reduce((a, b) => a + b, 0) : null; };
    const terms = model.mode === 'terms-g11' ? (() => {
      const subjects = model.gradeSheets.flatMap(sheet => {
        const row = find(sheet, st);
        return sheet.subjects.map(subject => {
          const grades = subject.cols.map(col => { const n = row && col != null ? num(row[col]) : null; return n === 0 ? null : n; });
          const filled = grades.filter(n => n != null);
          const final = filled.length ? filled.reduce((a, b) => a + b, 0) / filled.length : null;
          return { name: subject.name, group: sheet.title, grades, final,
            remark: final == null ? '' : final > 74.4 ? 'PASSED' : 'FAILED' };
        });
      });
      const average = subjects.find(s => /^average$/i.test(s.name));
      return { subjects: subjects.filter(s => s !== average),
        averages: average ? [...average.grades, average.final] : [null, null, null, null], remark: average ? average.remark : '' };
    })() : model.mode === 'terms' ? (() => {
      const rows = model.termSheets.map(sheet => find(sheet, st));
      const finalHeaders = model.termSheets[3].headers;
      const finalAverageCol = finalHeaders.findIndex(header => /^(general\s+)?average$/i.test(norm(header)));
      const finalRemarkCol = finalHeaders.findIndex(header => /^remarks?$/i.test(norm(header)));
      const subjects = model.subjects.map(subject => {
        const grades = rows.map(row => { const n = row ? num(row[subject.col]) : null; return n === 0 ? null : n; });
        const final = grades[3];
        return { name: subject.name, grades: grades.slice(0, 3), final, remark: final == null ? '' : (final > 74.4 ? 'PASSED' : 'FAILED') };
      });
      const averages = rows.map((row, i) => { const col = i === 3 ? finalAverageCol : model.averageCol; const n = row && col >= 0 ? num(row[col]) : null; return n === 0 ? null : n; });
      const finalRemark = finalRemarkCol >= 0 ? (rows[3] ? norm(rows[3][finalRemarkCol]) : '') : (averages[3] == null ? '' : (averages[3] > 74.4 ? 'PASSED' : 'FAILED'));
      return { subjects, averages, remark: finalRemark };
    })() : null;
    return {
      mode: model.mode, terms, adviser: model.adviser, principal: model.principal,
      student: st,
      age: st.birth ? new Date().getFullYear() - st.birth.getFullYear() : '',
      sem1: terms ? null : buildSem(model.sem1, st, cfg, 1),
      sem2: terms ? null : buildSem(model.sem2, st, cfg, 2),
      values, months,
      total: { days: sum('days'), present: sum('present'), absent: sum('absent'), tardy: sum('tardy'), mass: sum('mass') }
    };
  }

  const fmt = n => (n == null ? '' : String(round2(n)));

  root.ECard = { parse, buildCard, fmt, norm, key };
  if (typeof module !== 'undefined') module.exports = root.ECard;
})(typeof window !== 'undefined' ? window : globalThis);
