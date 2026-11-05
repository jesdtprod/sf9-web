(function () {
  'use strict';
  const cfg = window.ECARD_CONFIG, E = window.ECard;
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = E.fmt;
  const isTerms = card => card.mode !== 'quarters';
  const fmtFinal = n => n == null ? '' : Number(fmt(n)).toFixed(2);

  let model = null, handle = null, visible = [];   // students in the loaded workbook
  let schoolId = null;
  let signatories = null;
  let loadGeneration = 0, saveQueue = Promise.resolve();

  /* ---------- status & centralized toast ---------- */
  function toast(msg, isErr) {
    const c = $('toastContainer');
    if (!c) return;
    const t = document.createElement('div');
    t.className = 'toast' + (isErr ? ' err' : '');
    t.innerHTML = '<span class="toast-dot"></span><span>' + esc(msg) + '</span>';
    c.innerHTML = '';
    c.appendChild(t);
    requestAnimationFrame(() => t.classList.add('show'));
    setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 220); }, 3400);
  }
  function status(msg, isErr) {
    const el = $('status');
    el.textContent = msg;
    el.title = msg;
    el.classList.toggle('err', !!isErr);
    toast(msg, isErr);
  }
  function showSkeleton() {
    $('cards').innerHTML = '<div class="skeleton-page no-print"><div><div class="skel-box" style="height:24px;width:45%"></div><div class="skel-box" style="height:140px"></div><div class="skel-box" style="height:180px"></div></div><div><div class="skel-box" style="height:64px"></div><div class="skel-box" style="height:160px"></div><div class="skel-box" style="height:140px"></div></div></div>';
  }
  const formatName = mode => mode === 'terms-g11' ? '3 Semesters - G11' : mode === 'terms' ? '3 Semesters - G7-10' : '4 Quarters - G12';
  function showWorkbookPrompt(mismatch = false) {
    $('cards').innerHTML = '<section class="workbook-welcome no-print" aria-labelledby="workbookWelcomeTitle">' +
      '<div class="workbook-school"><img src="' + esc(cfg.school.logo) + '" alt=""><div><span>REPORT CARD WORKSPACE</span><strong>' + esc(cfg.school.name) + '</strong></div></div>' +
      '<div class="workbook-welcome-body"><div class="workbook-file-icon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5"/></svg></div>' +
      '<h2 id="workbookWelcomeTitle">' + (mismatch ? 'Open a matching workbook' : 'Ready to prepare your report cards?') + '</h2><p>' +
      (mismatch ? 'This workbook uses a different report card format.<br class="workbook-break"> Open a matching file or return to your loaded cards.' : 'Open your Excel workbook to view student records,<br class="workbook-break"> preview both sides, and print your cards.') + '</p>' +
      (mismatch ? '<div class="workbook-formats"><div><span>LOADED WORKBOOK</span><strong>' + esc(formatName(model.mode)) + '</strong></div><div><span>SELECTED FORMAT</span><strong>' + esc(formatName(selectedFormat())) + '</strong></div></div>' : '') +
      '<div class="workbook-actions"><button id="btnEmptyOpen" class="btn primary"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4m-4 4 4-4 4 4M4 16v4h16v-4"/></svg>Open Excel File</button>' +
      (mismatch ? '<button id="btnReturnFormat" class="btn">Return to Loaded Cards</button>' : '') + '</div>' +
      '<span class="workbook-file-types">Supports .xlsx and .xlsm workbooks</span></div>' +
      '<div class="workbook-privacy"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="5" y="10" width="14" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></svg>Your workbook stays on this device.</div></section>';
    $('btnEmptyOpen').onclick = () => $('btnOpen').click();
    if (mismatch) $('btnReturnFormat').onclick = () => { $('selFormat').value = model.mode; renderOne(); };
  }
  function setBtnLoading(btn, loading) {
    if (!btn) return;
    btn.disabled = loading;
    const ico = btn.querySelector('.btn-icon');
    const sp = btn.querySelector('.spinner');
    if (loading) {
      if (ico) ico.style.display = 'none';
      if (!sp) { const s = document.createElement('span'); s.className = 'spinner'; btn.prepend(s); }
    } else {
      if (sp) sp.remove();
      if (ico) ico.style.display = '';
    }
  }

  /* ---------- remember the chosen file between visits (Chrome / Edge) ---------- */
  const DB = { open() { return new Promise((res, rej) => { const r = indexedDB.open('ecard', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
    async get(k) { try { const db = await this.open(); return await new Promise(res => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(null); }); } catch (e) { return null; } },
    async set(k, v) { try { const db = await this.open(); return await new Promise(res => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => res(true); t.onerror = () => res(false); t.onabort = () => res(false); }); } catch (e) { return false; } } };

  /* ---------- loading ---------- */
  function loadBuffer(buf, label, generation) {
    if (generation !== loadGeneration) return false;
    try {
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const nextModel = E.parse(XLSX, wb, cfg);
      if (generation !== loadGeneration) return false;
      model = nextModel;
      if (!signatories) {
        signatories = { principal: model.principal || '', adviser: model.adviser || '' };
        fillSignatories();
      }
      $('selFormat').value = model.mode;
      fillStudents();
      updateNavigation();
      status('Loaded "' + label + '" - ' + model.students.length + ' students.');
      return true;
    } catch (e) {
      if (generation === loadGeneration) status(e.message || String(e), true);
      return false;
    }
  }

  function rememberWorkbook(buffer, name, chosenHandle) {
    // Keep the bytes and their own handle together, in the same database record.
    const cacheKey = 'workbook:' + schoolId;
    const save = async () => {
      let saved = await DB.set(cacheKey, { buffer, name, handle: chosenHandle || null });
      if (!saved && chosenHandle) saved = await DB.set(cacheKey, { buffer, name, handle: null });
      return saved;
    };
    saveQueue = saveQueue.then(save, save);
    return saveQueue;
  }

  async function loadFile(file, btn, generation, chosenHandle) {
    if (generation !== loadGeneration) return false;
    const button = btn || $('btnOpen');
    setBtnLoading(button, true);
    showSkeleton();
    try {
      const buffer = await file.arrayBuffer();
      if (generation !== loadGeneration) return false;
      if (!loadBuffer(buffer, file.name, generation)) {
        if (generation === loadGeneration) { if (model) renderOne(); else showWorkbookPrompt(); }
        return false;
      }
      handle = chosenHandle || null;
      $('btnReload').hidden = false;
      const saved = await rememberWorkbook(buffer, file.name, handle);
      if (!saved && generation === loadGeneration) status('Workbook loaded, but the browser could not save it for the next visit.', true);
      return true;
    } catch (e) {
      if (generation === loadGeneration) {
        status('Could not read the file: ' + (e.message || e), true);
        if (model) renderOne(); else showWorkbookPrompt();
      }
      return false;
    } finally {
      setBtnLoading(button, false);
    }
  }

  async function readHandle(chosenHandle, generation) {
    if (!chosenHandle) return;
    try {
      let perm = await chosenHandle.queryPermission({ mode: 'read' });
      if (generation !== loadGeneration) return;
      if (perm !== 'granted') perm = await chosenHandle.requestPermission({ mode: 'read' });
      if (generation !== loadGeneration) return;
      if (perm !== 'granted') return status('Permission to read the file was not given.', true);
      const file = await chosenHandle.getFile();
      await loadFile(file, !$('btnReload').hidden ? $('btnReload') : $('btnOpen'), generation, chosenHandle);
    } catch (e) {
      if (generation === loadGeneration) status('Could not read the file: ' + (e.message || e) + ' - save the workbook in Excel, then press Reload.', true);
    }
  }

  $('btnOpen').onclick = async () => {
    const generation = ++loadGeneration;
    if (window.showOpenFilePicker) {
      try {
        const [chosenHandle] = await window.showOpenFilePicker({ types: [{ description: 'Excel workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }] });
        if (generation === loadGeneration) await readHandle(chosenHandle, generation);
        return;
      } catch (e) {
        if (e.name === 'AbortError' || generation !== loadGeneration) return;
      }
    }
    $('fileFallback').click();
  };
  $('fileFallback').onchange = async ev => {
    const file = ev.target.files[0]; if (!file) return;
    const generation = ++loadGeneration;
    await loadFile(file, $('btnOpen'), generation, null);
    ev.target.value = '';
  };
  $('btnReload').onclick = () => {
    const generation = ++loadGeneration;
    if (handle) readHandle(handle, generation); else $('fileFallback').click();
  };

  async function restore() {
    const generation = loadGeneration;
    showSkeleton();
    const [saved, legacy, legacyHandle] = await Promise.all([DB.get('workbook:' + schoolId), schoolId === 'nds' ? DB.get('workbook') : null, schoolId === 'nds' ? DB.get('handle') : null]);
    const cached = saved || legacy;
    if (model || generation !== loadGeneration) return;
    if (!cached) {
      showWorkbookPrompt();
      return;
    }
    // Legacy records stored their handle separately; newly saved records bind it to the bytes.
    const savedHandle = Object.prototype.hasOwnProperty.call(cached, 'handle') ? cached.handle : legacyHandle;
    if (!loadBuffer(cached.buffer, cached.name, generation)) {
      showWorkbookPrompt();
      return;
    }
    handle = savedHandle || null;
    $('btnReload').hidden = false;
    status('Restored "' + cached.name + '" - ' + model.students.length + ' students. Saved local copy; ' + (savedHandle ? 'use Reload' : 'use Reload and select the file again') + ' after editing Excel.');
    if (!savedHandle) return;
    try {
      const permission = await savedHandle.queryPermission({ mode: 'read' });
      if (generation !== loadGeneration || handle !== savedHandle) return;
      if (permission === 'granted') {
        const file = await savedHandle.getFile();
        await loadFile(file, $('btnReload'), generation, savedHandle);
      }
    } catch (e) {
      if (generation === loadGeneration) status('Showing the saved workbook. The original file is unavailable; open it again to update.');
    }
  }

  /* ---------- section / student pickers ---------- */
  const secLabel = s => [s.grade, s.section].filter(Boolean).join(' \u2013 ') || '(no section)';
  function fillStudents() {
    visible = model.students;
    $('selStudent').innerHTML = visible.map((student, i) => '<option value="' + i + '">' + esc(student.name) + '</option>').join('');
    renderOne();
  }
  $('selStudent').onchange = renderOne;
  $('selCardView').onchange = renderOne;
  const step = d => { const sel = $('selStudent'); const i = Math.min(Math.max(sel.selectedIndex + d, 0), visible.length - 1); sel.selectedIndex = i; renderOne(); };
  $('btnPrev').onclick = () => step(-1);
  $('btnNext').onclick = () => step(1);
  document.addEventListener('keydown', e => { if (e.target.closest('input, textarea, select, [role=combobox], [role=listbox], .dropdown-panel')) return; if (e.key === 'ArrowLeft') step(-1); if (e.key === 'ArrowRight') step(1); });

  const selectedFormat = () => $('selFormat').value;
  $('selFormat').addEventListener('change', renderOne);

  function updateNavigation() {
    const format = selectedFormat();
    const view = $('selCardView');
    const viewMode = format !== 'quarters' ? 'terms' : 'quarters';
    if (view.dataset.mode !== viewMode) {
      view.innerHTML = viewMode === 'terms'
        ? '<option value="whole">Whole Card</option><option value="t1">Term 1</option><option value="t2">Term 2</option><option value="t3">Term 3</option>'
        : '<option value="whole">Whole Card</option><option value="q1">1st Quarter</option><option value="q2">2nd Quarter</option><option value="q3">3rd Quarter</option><option value="q4">4th Quarter</option>';
      view.dataset.mode = viewMode;
    }
    $('cardViewField').hidden = false;
    const incompatible = !!model && selectedFormat() !== model.mode;
    $('btnPrint').disabled = !model || incompatible;
    $('btnPrintAll').disabled = !model || incompatible;
    const index = $('selStudent').selectedIndex;
    $('btnPrev').disabled = !model || incompatible || index <= 0;
    $('btnNext').disabled = !model || incompatible || index < 0 || index >= visible.length - 1;
  }
  updateNavigation();

  function renderOne() {
    updateNavigation();
    window.ECardDropdowns.sync();
    if (!model || !visible.length) return;
    if (selectedFormat() !== model.mode) {
      showWorkbookPrompt(true);
      return;
    }
    const st = visible[+$('selStudent').value || 0];
    $('cards').innerHTML = cardHtml(E.buildCard(model, st, cfg));
    $('cards').scrollTop = 0;
  }
  $('btnPrint').onclick = () => { if (!model || selectedFormat() !== model.mode) return; renderOne(); window.print(); };
  $('btnPrintAll').onclick = () => {
    if (!model || !visible.length || selectedFormat() !== model.mode) return;
    const selected = visible[+$('selStudent').value || 0];
    const section = secLabel(selected);
    $('cards').innerHTML = visible.filter(student => secLabel(student) === section).map(student => cardHtml(E.buildCard(model, student, cfg))).join('');
    window.print();
  };
  window.addEventListener('afterprint', () => { if (model) renderOne(); });

  /* ---------- card templates ---------- */
  const cell = (v, cls) => '<td class="' + (cls || 'c') + '">' + esc(v) + '</td>';
  const remarkCell = r => '<td class="c ' + (r === 'FAILED' ? 'fail' : '') + '">' + esc(r) + '</td>';
  const field = (label, val) => '<div class="f' + (label === 'Section:' ? ' section-field' : '') + '"><b>' + label + '</b><span>' + esc(val) + '</span></div>';

  function attendanceTable(c) {
    const m = c.months;
    const zero = v => (v == null || (v === 0 && !c.showAttendanceZero) ? '' : v);
    const row = (label, k, blank0) => '<tr><th scope="row">' + label + '</th>' +
      m.map(x => cell(blank0 ? zero(x[k]) : (x[k] == null ? '' : x[k]))).join('') +
      cell(blank0 ? zero(c.total[k]) : (c.total[k] == null ? '' : c.total[k])) + '</tr>';
    return '<table class="front-attendance"><colgroup><col class="attendance-label"><col span="' + (m.length + 1) + '"></colgroup><thead><tr><th></th>' +
      m.map(x => '<th scope="col"><span>' + esc(x.name.toUpperCase()) + '</span></th>').join('') + '<th scope="col"><span>TOTAL</span></th></tr></thead><tbody>' +
      row('No. of Days', 'days', true) + row('No. of Days<br>Present', 'present', true) +
      (isTerms(c) ? row('No. of Days<br>Absent', 'absent', false) + '<tr class="term-attendance-spacer" aria-hidden="true"><td colspan="' + (m.length + 2) + '"></td></tr>' + row('Tardy', 'tardy', false) : row('Tardy', 'tardy', false) + row('Mass<br>Attendance', 'mass', true)) + '</tbody></table>';
  }

  function frontPage(c) {
    const s = c.student, sc = Object.assign({}, cfg.school, { adviser: c.adviser || '', principal: c.principal || '' }, isTerms(c) ? { subtitle: c.mode === 'terms-g11' ? 'Senior High School' : 'Junior High School', title: "LEARNER'S PROGRESS REPORT CARD", adviser: c.adviser || '', principal: c.principal || '' } : {});
    const blank = '<span class="write-line"></span>';
    return '<section class="page front' + (isTerms(c) ? ' term-front' : '') + '">' +
      '<div class="col front-records">' +
        '<h3 class="front-heading attendance-heading">ATTENDANCE RECORD</h3>' + attendanceTable(c) +
        '<h3 class="front-heading parent-heading">PARENT&#39;S / GUARDIAN&#39;S SIGNATURE</h3>' +
        '<div class="quarter-signatures">' + (isTerms(c) ? ['1st', '2nd', '3rd'] : ['1st', '2nd', '3rd', '4th']).map(q => '<div><span>' + q + (isTerms(c) ? ' Term' : ' Quarter') + '</span>' + blank + '</div>').join('') + '</div>' +
        '<h3 class="front-heading">CERTIFICATE OF TRANSFER</h3>' +
        '<div class="transfer-row"><span>Admitted to Grade:</span>' + blank + '<span>Section:</span>' + blank + '</div>' +
        '<div class="transfer-row"><span>' + (isTerms(c) ? 'Eligible for Admission to Grade:' : 'Eligible for Transfer to') + '</span>' + blank + '</div>' +
        '<div class="transfer-approval">Approved by:<div class="approval-signatures"><div>' + blank + '<b>School Principal</b></div><div>' + blank + '<b>Class Adviser</b></div></div></div>' +
        '<h3 class="front-heading cancel-heading">CANCELLATION OF ELIGIBILITY TO TRANSFER</h3>' +
        '<div class="cancellation-fields"><div class="transfer-row"><span>Admitted In:</span>' + blank + '</div><div class="transfer-row"><span>Date:</span>' + blank + '</div></div>' +
        '<div class="cancellation-approval"><span>Approved by:</span><div>' + blank + '<b>School Principal</b></div></div>' +
      '</div>' +
      '<div class="col front-cover">' +
        '<div class="front-school"><div class="school-emblem"><span>' + (c.mode === 'terms' ? 'SF 9-JHS' : 'SF 9-SHS') + '</span><img src="assets/deped.png" alt="School seal"></div><div class="school-name"><h2>' + esc(sc.name) + '</h2><div>' + esc(sc.address) + '</div>' + '<div class="school-recognition">' + esc(sc.recognition) + '</div>' + '<h3>' + esc(sc.subtitle.toUpperCase()) + '</h3></div><img class="school-seal" src="' + esc(sc.logo) + '" alt="' + esc(sc.name) + ' seal"></div>' +
        '<h1 class="front-card-title">' + esc(sc.title) + (isTerms(c) && s.schoolYear ? '<span class="term-school-year">School Year ' + esc(s.schoolYear) + '</span>' : '') + '</h1>' +
        '<div class="info front-info">' +
          (isTerms(c)
            ? '<div class="row term-name-row">' + field('Name:', s.name) + '</div>' +
              '<div class="row">' + field('Age:', c.age) + field('Sex:', s.sex) + '</div>' +
              '<div class="row">' + field('Grade:', s.grade) + field('Section:', s.section) + '</div>' +
              '<div class="row lrn-row">' + field('LRN:', s.lrn) + '</div>' +
              '<div class="row">' + field('Curriculum:', s.curriculum) + '</div>'
            : '<div class="row full-name-row">' + field('Name:', s.name) + '</div>' +
              '<div class="row">' + field('Age:', c.age) + field('Sex:', s.sex) + '</div>' +
              '<div class="row">' + field('Grade:', s.grade) + field('Section:', s.section) + '</div>' +
              '<div class="row">' + field('School Year:', s.schoolYear) + field('Strand:', s.strand) + '</div>' +
              '<div class="row lrn-row">' + field('LRN:', s.lrn) + field('Curriculum:', s.curriculum) + '</div>') +
        '</div>' +
        '<div class="front-parent-message"><p>Dear Parent,</p><p>This Report Card shows the ability and the progress of your child has made in the different learning areas as well as his/her core values.</p><p>The school welcomes you if you desire to know more about the progress of your child.</p></div>' +
        '<div class="cover-signatures"><div class="cover-adviser"><span>' + esc(sc.adviser) + '</span><b>Class Adviser</b></div><div class="cover-principal"><span>' + esc(sc.principal) + '</span><b>School Principal</b></div></div>' +
      '</div></section>';
  }

  function semTable(c, sem, label) {
    const first = label.startsWith('First');
    const gradeCell = (value, cls) => '<td class="' + cls + '"><span class="back-text">' + esc(value) + '</span></td>';
    let h = '<table class="back-grades"><colgroup><col class="grade-subject"><col class="grade-quarter"><col class="grade-quarter"><col class="grade-final"><col class="grade-remark"></colgroup><thead><tr><th rowspan="2">Subjects</th><th colspan="2">QUARTERS</th><th rowspan="2">' + label + '<br>Final Grade</th><th rowspan="2">Remarks</th></tr>' +
            '<tr><th>' + (first ? '1' : '3') + '</th><th>' + (first ? '2' : '4') + '</th></tr></thead><tbody>';
    sem.groups.forEach(g => {
      if (!g.subjects.length) return;
      h += '<tr class="grp"><td colspan="5">' + esc(g.title) + '</td></tr>';
      g.subjects.forEach(s => { h += '<tr>' + gradeCell(s.name, 'l') + gradeCell(fmt(s.q1), 'c') + gradeCell(fmt(s.q2), 'c') + gradeCell(fmtFinal(s.final), 'c') + gradeCell(s.remark, 'c ' + (s.remark === 'FAILED' ? 'fail' : '')) + '</tr>'; });
    });
    h += '<tr class="back-average"><td></td><th colspan="2">General Average</th>' + gradeCell(fmtFinal(sem.genAvg), 'c') + gradeCell(sem.remark, 'c ' + (sem.remark === 'FAILED' ? 'fail' : '')) + '</tr></tbody></table>';
    return h;
  }

  function valuesTable(c) {
    const count = isTerms(c) ? 3 : 4;
    let h = '<table class="back-values"><colgroup><col class="values-name"><col class="values-statement"><col span="' + count + '" class="values-quarter"></colgroup><thead><tr><th rowspan="2">Core Values</th><th rowspan="2">Behavior Statements</th><th colspan="' + count + '">' + (count === 3 ? 'TERM' : 'Quarter') + '</th></tr><tr>' + Array.from({length: count}, (_, i) => '<th>' + (i + 1) + '</th>').join('') + '</tr></thead><tbody>';
    c.values.forEach(v => { h += '<tr><td class="l"><b class="back-text">' + esc(v.code) + ' &ndash; ' + esc(v.name) + '</b></td><td class="l vtext"><span class="back-text">' + esc(v.statement) + '</span></td>' + v.q.map(x => '<td class="c"><span class="back-text">' + esc(x) + '</span></td>').join('') + '</tr>'; });
    return h + '</tbody></table>';
  }

  function termBackPage(c) {
    const gradeCell = (value, cls) => '<td class="' + cls + '"><span class="back-text">' + esc(value) + '</span></td>';
    const gradeRow = (name, grades, final, remarks, average, activeTerms) => '<tr' + (average ? ' class="term-average"' : '') + '>' +
      gradeCell(name, 'l ' + (/^(music|arts|p\.?e\.?|health)$/i.test(name) ? 'term-component' : 'term-learning-area')) + grades.map((n, i) => gradeCell(average ? fmtFinal(n) : (n == null ? '' : String(n)), 'c' + (activeTerms && !activeTerms[i] ? ' unused-term' : ''))).join('') +
      gradeCell(average || c.mode === 'terms-g11' ? fmtFinal(final) : '', 'c') + gradeCell(average || c.mode === 'terms-g11' ? remarks : '', 'c ' + (remarks === 'FAILED' ? 'fail' : '')) + '</tr>';
    let group = '';
    const rows = c.terms.subjects.map(subject => {
      const heading = subject.group && subject.group !== group ? '<tr class="grp"><td colspan="6">' + esc(subject.group) + '</td></tr>' : '';
      group = subject.group;
      return heading + gradeRow(subject.name, subject.grades, subject.final, subject.remark, false, subject.activeTerms);
    }).join('') +
      gradeRow('Average', c.terms.averages.slice(0, 3), c.terms.averages[3], c.terms.remark, true);
    const markings = cfg.markings.map(m => '<tr><td>' + esc(m[0]) + '</td><td>' + esc(m[1]) + '</td></tr>').join('');
    const scale = [['Advancing', '90-100', 'Passed'], ['Benchmarking', '85-89', 'Passed'], ['Connecting', '80-84', 'Passed'], ['Developing', '75-79', 'Passed'], ['Emerging', 'Below 75', 'Failed']]
      .map(r => '<tr><td>' + r.map(esc).join('</td><td>') + '</td></tr>').join('');
    return '<section class="page back terms' + (c.mode === 'terms-g11' ? ' g11' : '') + '">' +
      '<div class="col"><h3 class="sec">REPORT ON LEARNING PROGRESS AND ACHIEVEMENT</h3>' +
      '<table class="back-grades term-grades"><colgroup><col class="term-subject"><col span="3" class="term-period"><col class="term-final"><col class="term-remark"></colgroup><thead><tr><th rowspan="2">Learning Areas</th><th colspan="3">TERM</th><th rowspan="2">Final Rating</th><th rowspan="2">Remarks</th></tr><tr><th>1</th><th>2</th><th>3</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<div class="back-legends"><table class="scale-guide"><thead><tr><th>Descriptors</th><th>Grade Scale</th><th>Remarks</th></tr></thead><tbody>' + scale + '</tbody></table></div></div>' +
      '<div class="col"><h3 class="sec">REPORT ON LEARNER&#39;S OBSERVED VALUES</h3>' + valuesTable(c) +
      '<div class="back-legends"><div class="term-values-legend-title">Observed Values</div><table class="markings-guide"><thead><tr><th>Marking</th><th>Non-numerical Rating</th></tr></thead><tbody>' + markings + '</tbody></table></div></div></section>';
  }

  function backPage(c) {
    if (isTerms(c)) return termBackPage(c);
    const mk = cfg.markings.map(m => '<tr><td>' + esc(m[0]) + '</td><td>' + esc(m[1]) + '</td></tr>').join('');
    const sc = cfg.scale.map(r => '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>').join('');
    return '<section class="page back">' +
      '<div class="col"><h3 class="sec">ACADEMIC TRACK</h3>' + semTable(c, c.sem1, 'First Semester') + '<div class="back-semester-gap"></div>' + semTable(c, c.sem2, 'Second Semester') + '</div>' +
      '<div class="col"><h3 class="sec">REPORT ON LEARNER&#39;S OBSERVED VALUES</h3>' + valuesTable(c) +
        '<div class="back-legends"><table class="markings-guide"><thead><tr><th>Marking</th><th>Non-numerical Rating</th></tr></thead><tbody>' + mk + '</tbody></table>' +
        '<table class="scale-guide"><thead><tr><th>Descriptions</th><th>Grading Scale</th><th>Remarks</th></tr></thead><tbody>' + sc + '</tbody></table></div></div>' +
      '</section>';
  }

  function quarterCard(card, quarter) {
    const allowedMonths = { 1: ['JUN', 'JUL', 'AUG'], 2: ['SEP', 'OCT'], 3: ['NOV', 'DEC', 'JAN'], 4: ['FEB', 'MAR'] }[quarter];
    const visibleMonth = month => allowedMonths.includes(String(month.name).trim().toUpperCase().slice(0, 3));
    const months = card.months.map(month => visibleMonth(month) ? month : {
      ...month, days: null, present: null, absent: null, tardy: null, mass: null
    });
    const maskSemester = (semester, showSelectedQuarter) => ({
      ...semester,
      groups: semester.groups.map(group => ({
        ...group,
        subjects: group.subjects.map(subject => ({
          ...subject, q1: showSelectedQuarter && (quarter === 1 || quarter === 3) ? subject.q1 : null,
          q2: showSelectedQuarter && (quarter === 2 || quarter === 4) ? subject.q2 : null,
          final: showSelectedQuarter && (quarter === 2 || quarter === 4) ? subject.final : null,
          remark: showSelectedQuarter && (quarter === 2 || quarter === 4) ? subject.remark : ''
        }))
      })),
      genAvg: showSelectedQuarter && (quarter === 2 || quarter === 4) ? semester.genAvg : null,
      remark: showSelectedQuarter && (quarter === 2 || quarter === 4) ? semester.remark : ''
    });
    return {
      ...card, months, showAttendanceZero: true,
      total: quarter === 4 ? card.total : { days: null, present: null, absent: null, tardy: null, mass: null },
      sem1: maskSemester(card.sem1, quarter <= 2), sem2: maskSemester(card.sem2, quarter >= 3),
      values: card.values.map(value => ({ ...value, q: value.q.map((mark, index) => index === quarter - 1 ? mark : '') }))
    };
  }

  function termCard(card, term) {
    const allowed = { 1: ['JUN', 'JUL', 'AUG'], 2: ['SEP', 'OCT', 'NOV', 'DEC'], 3: ['JAN', 'FEB', 'MAR'] }[term];
    const maskGrades = grades => grades.map((grade, i) => i === term - 1 ? grade : null);
    return {
      ...card,
      months: card.months.map(month => allowed.includes(String(month.name).trim().toUpperCase().slice(0, 3)) ? month :
        { ...month, days: null, present: null, absent: null, tardy: null, mass: null }),
      total: term === 3 ? card.total : { days: null, present: null, absent: null, tardy: null, mass: null },
      showAttendanceZero: true,
      terms: { ...card.terms,
        subjects: card.terms.subjects.map(subject => ({ ...subject, grades: maskGrades(subject.grades),
          final: term === 3 ? subject.final : null, remark: term === 3 ? subject.remark : '' })),
        averages: [...maskGrades(card.terms.averages.slice(0, 3)), term === 3 ? card.terms.averages[3] : null],
        remark: term === 3 ? card.terms.remark : '' },
      values: card.values.map(value => ({ ...value, q: value.q.map((mark, index) => index === term - 1 ? mark : '') }))
    };
  }

  function cardHtml(card) {
    card = { ...card, ...signatories };
    const selectedView = $('selCardView').value;
    const quarter = { q1: 1, q2: 2, q3: 3, q4: 4 }[selectedView];
    const term = { t1: 1, t2: 2, t3: 3 }[selectedView];
    const c = card.mode === 'quarters' && quarter ? quarterCard(card, quarter) : isTerms(card) && term ? termCard(card, term) : card;
    const viewClass = c.mode === 'quarters' && (quarter >= 2 && quarter <= 4)
      ? 'page-view ' + selectedView + '-values-only' : isTerms(c) && term >= 2 ? 'page-view term-' + term + '-values-only' : 'page-view';
    return '<div class="card-spread"><div class="' + viewClass + '">' + frontPage(c) + '</div><div class="' + viewClass + '">' + backPage(c) + '</div></div>';
  }

  function fillSignatories() {
    $('principalName').value = signatories ? signatories.principal : '';
    $('adviserName').value = signatories ? signatories.adviser : '';
  }
  function showSignatories(show) {
    $('signatureSettings').hidden = !show;
    $('btnSignatories').setAttribute('aria-expanded', String(show));
    if (show) $('principalName').focus();
    else $('btnSignatories').focus();
  }
  $('btnSignatories').onclick = () => showSignatories($('signatureSettings').hidden);
  if ($('btnDismissSignatories')) $('btnDismissSignatories').onclick = () => showSignatories(false);
  $('signatureSettings').addEventListener('click', e => { if (e.target === $('signatureSettings')) showSignatories(false); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('signatureSettings').hidden) showSignatories(false); });
  const editSignatories = () => {
    signatories = { principal: $('principalName').value.trim(), adviser: $('adviserName').value.trim() };
    try { localStorage.setItem('ecard:signatories:' + schoolId, JSON.stringify(signatories)); }
    catch (e) { status('Names updated for this visit, but browser storage is unavailable.', true); }
    renderOne();
  };
  $('principalName').oninput = editSignatories;
  $('adviserName').oninput = editSignatories;
  $('btnCloseSignatories').onclick = () => {
    const btn = $('btnCloseSignatories');
    setBtnLoading(btn, true);
    editSignatories();
    setTimeout(() => {
      setBtnLoading(btn, false);
      showSignatories(false);
      toast('Signatories saved for ' + (cfg.school.name || 'this school') + '.');
    }, 220);
  };

  function chooseSchool(id) {
    ++loadGeneration;
    schoolId = id;
    signatories = null;
    try {
      const saved = JSON.parse(localStorage.getItem('ecard:signatories:' + id));
      if (saved && typeof saved.principal === 'string' && typeof saved.adviser === 'string') signatories = saved;
    } catch (e) { /* Names remain editable when browser storage is unavailable. */ }
    fillSignatories();
    $('signatureSettings').hidden = true;
    $('btnSignatories').setAttribute('aria-expanded', 'false');
    Object.assign(cfg.school, cfg.templates[id]);
    model = null; handle = null; visible = [];
    $('selStudent').innerHTML = '';
    $('btnReload').hidden = true;
    $('status').textContent = 'No file loaded. Open your E-CARD workbook to get started.';
    $('status').classList.toggle('err', false);
    $('workspaceLogo').src = cfg.school.logo;
    $('workspaceLogo').alt = cfg.school.name + ' seal';
    $('workspaceSchool').textContent = cfg.school.name;
    $('schoolFavicon').href = cfg.school.logo;
    document.title = 'E-Card | ' + cfg.school.name;
    $('schoolLanding').hidden = true;
    $('schoolWorkspace').hidden = false;
    updateNavigation();
    fitPreview();
    restore();
  }
  $('chooseNdp').onclick = () => chooseSchool('ndp');
  $('chooseNds').onclick = () => chooseSchool('nds');
  $('btnChangeSchool').onclick = () => {
    ++loadGeneration;
    model = null; handle = null; visible = [];
    $('schoolWorkspace').hidden = true;
    $('schoolLanding').hidden = false;
    $('schoolFavicon').href = 'assets/deped.png';
    document.title = 'E-Card DepEd SF9';
    $('chooseNdp').focus();
  };

  // Fit the preview to the space below the toolbar; paper dimensions stay fixed.
  const preview = $('cards');
  let userZoom = null; // null = fit whole page; otherwise a scale (1 = 100%)
  const zoomRange = $('zoomRange'), zoomValue = $('zoomValue');
  function fitScale() {
    return Math.min(1, Math.max(1, preview.clientWidth - 48) / (1056 * 2), Math.max(1, preview.clientHeight - 32) / 816);
  }
  function fitPreview() {
    const scale = userZoom === null ? fitScale() : userZoom;
    const pct = Math.round(scale * 100);
    preview.style.setProperty('--preview-scale', scale);
    preview.classList.toggle('zoomed', userZoom !== null);
    zoomRange.value = pct;
    const fillPct = Math.max(0, Math.min(100, ((pct - 25) / (300 - 25)) * 100));
    zoomRange.style.setProperty('--range-pct', fillPct + '%');
    zoomValue.textContent = userZoom === null ? 'Fit' : pct + '%';
    $('zoomFit').classList.toggle('active', userZoom === null);
    $('zoom100').classList.toggle('active', userZoom !== null && pct === 100);
  }
  function setZoom(z) {
    const keep = preview.scrollTop / Math.max(1, preview.scrollHeight);
    userZoom = z === null ? null : Math.min(3, Math.max(0.25, z));
    fitPreview();
    if (userZoom === null) {
      const n = preview.querySelectorAll('.card-spread').length;
      preview.scrollTop = Math.round(keep * n) * preview.clientHeight;
    } else preview.scrollTop = keep * preview.scrollHeight;
  }
  const curScale = () => userZoom === null ? fitScale() : userZoom;
  $('zoomIn').addEventListener('click', () => setZoom(curScale() + 0.1));
  $('zoomOut').addEventListener('click', () => setZoom(curScale() - 0.1));
  $('zoomFit').addEventListener('click', () => setZoom(null));
  $('zoom100').addEventListener('click', () => setZoom(1));
  zoomRange.addEventListener('input', () => setZoom(zoomRange.value / 100));
  let wheelTotal = 0, lastWheel = 0, wheelLockedUntil = 0;
  preview.addEventListener('wheel', event => {
    if (event.ctrlKey) { event.preventDefault(); setZoom(curScale() + (event.deltaY < 0 ? 0.1 : -0.1)); return; }
    if (userZoom !== null) return; // zoomed: normal scrolling
    if (!preview.querySelector('.card-spread') || Math.abs(event.deltaX) > Math.abs(event.deltaY)) return;
    event.preventDefault();
    const now = performance.now();
    if (now < wheelLockedUntil) return;
    if (now - lastWheel > 180 || Math.sign(wheelTotal) !== Math.sign(event.deltaY)) wheelTotal = 0;
    lastWheel = now;
    wheelTotal += event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? preview.clientHeight : 1);
    if (Math.abs(wheelTotal) < 35) return;
    const current = Math.round(preview.scrollTop / preview.clientHeight);
    const next = Math.max(0, Math.min(preview.querySelectorAll('.card-spread').length - 1, current + Math.sign(wheelTotal)));
    wheelTotal = 0;
    wheelLockedUntil = now + 600;
    preview.scrollTo({ top: next * preview.clientHeight, behavior: 'smooth' });
  }, { passive: false });
  new ResizeObserver(fitPreview).observe(preview);
  fitPreview();
})();
