/* global JSZip, PDFLib */
(() => {
  'use strict';

  const EXPECTED_PERIODS = ['Advisory', 'A', 'B', 'C', 'D', 'E', 'F', 'G'];
  const state = { students: [], issues: [], sourceName: '' };

  const el = (id) => document.getElementById(id);
  const fileInput = el('fileInput');
  const dropZone = el('dropZone');
  const fileStatus = el('fileStatus');
  const workspace = el('workspace');
  const eventTitle = el('eventTitle');
  const studentSelect = el('studentSelect');
  const preview = el('preview');
  const stats = el('stats');
  const warningPanel = el('warningPanel');
  const downloadNote = el('downloadNote');

  function clean(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
  }

  function escapeHtml(value) {
    return clean(value).replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
  }

  function columnIndex(ref) {
    const letters = (ref.match(/[A-Z]+/i) || ['A'])[0].toUpperCase();
    let index = 0;
    for (const char of letters) index = index * 26 + char.charCodeAt(0) - 64;
    return index - 1;
  }

  function xmlText(node) {
    return node ? Array.from(node.getElementsByTagName('t')).map((n) => n.textContent || '').join('') : '';
  }

  async function parseXlsx(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const requireEntry = (path) => {
      const entry = zip.file(path);
      if (!entry) throw new Error(`The workbook is missing ${path}.`);
      return entry;
    };
    const parser = new DOMParser();
    const parseXml = async (path) => parser.parseFromString(await requireEntry(path).async('string'), 'application/xml');

    const sharedStrings = [];
    if (zip.file('xl/sharedStrings.xml')) {
      const sharedDoc = await parseXml('xl/sharedStrings.xml');
      for (const si of sharedDoc.getElementsByTagName('si')) sharedStrings.push(xmlText(si));
    }

    const workbookDoc = await parseXml('xl/workbook.xml');
    const relsDoc = await parseXml('xl/_rels/workbook.xml.rels');
    const firstSheet = workbookDoc.getElementsByTagName('sheet')[0];
    if (!firstSheet) throw new Error('No worksheet was found.');
    const relId = firstSheet.getAttribute('r:id') || firstSheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    const relationship = Array.from(relsDoc.getElementsByTagName('Relationship')).find((r) => r.getAttribute('Id') === relId);
    if (!relationship) throw new Error('The first worksheet could not be located.');
    let target = relationship.getAttribute('Target').replace(/^\//, '');
    if (!target.startsWith('xl/')) target = `xl/${target.replace(/^\.\//, '')}`;
    const sheetDoc = await parseXml(target);

    const rows = [];
    for (const rowNode of sheetDoc.getElementsByTagName('row')) {
      const row = [];
      for (const cell of rowNode.getElementsByTagName('c')) {
        const idx = columnIndex(cell.getAttribute('r') || 'A1');
        const type = cell.getAttribute('t');
        const valueNode = cell.getElementsByTagName('v')[0];
        let value = '';
        if (type === 'inlineStr') value = xmlText(cell.getElementsByTagName('is')[0]);
        else if (type === 's') value = sharedStrings[Number(valueNode?.textContent || 0)] || '';
        else if (type === 'b') value = valueNode?.textContent === '1' ? 'TRUE' : 'FALSE';
        else value = valueNode?.textContent || '';
        row[idx] = value;
      }
      rows.push(row);
    }
    return rows;
  }

  function identifyColumns(rows) {
    const aliases = {
      student: ['student', 'student name', 'name'],
      course: ['class', 'class/course', 'course', 'course name'],
      teacher: ['teacher', 'instructor'],
      location: ['location', 'room', 'classroom'],
      period: ['period', 'block'],
      time: ['time', 'meeting time']
    };
    for (let r = 0; r < Math.min(rows.length, 25); r++) {
      const normalized = rows[r].map((v) => clean(v).toLowerCase());
      const columns = {};
      for (const [key, names] of Object.entries(aliases)) {
        columns[key] = normalized.findIndex((v) => names.includes(v));
      }
      if (Object.values(columns).every((v) => v >= 0)) return { headerRow: r, columns };
    }
    throw new Error('Required columns were not found. Expected Student, Class, Teacher, Location, Period, and Time.');
  }

  function normalizePeriod(value) {
    const raw = clean(value);
    if (/^advisory$/i.test(raw)) return 'Advisory';
    const letter = raw.toUpperCase();
    return /^[A-G]$/.test(letter) ? letter : raw;
  }

  function cleanCourse(course, period) {
    const value = clean(course);
    if (/^[A-G]$/.test(period)) return value.replace(new RegExp(`\\s*\\(${period}\\)\\s*$`, 'i'), '');
    return value.replace(/\s*\(Advisory\)\s*$/i, '');
  }

  function buildStudents(rows) {
    const { headerRow, columns } = identifyColumns(rows);
    const grouped = new Map();
    for (const row of rows.slice(headerRow + 1)) {
      const name = clean(row[columns.student]);
      if (!name) continue;
      const period = normalizePeriod(row[columns.period]);
      const record = {
        name,
        course: cleanCourse(row[columns.course], period),
        teacher: clean(row[columns.teacher]),
        location: clean(row[columns.location]),
        period,
        time: clean(row[columns.time])
      };
      if (!grouped.has(name)) grouped.set(name, []);
      grouped.get(name).push(record);
    }

    const students = Array.from(grouped, ([name, records]) => {
      const byPeriod = new Map();
      for (const record of records) {
        if (!byPeriod.has(record.period)) byPeriod.set(record.period, []);
        byPeriod.get(record.period).push(record);
      }
      const rowsOrdered = EXPECTED_PERIODS.map((period) => byPeriod.get(period)?.[0] || ({ name, period, time: '', course: '', teacher: '', location: '', missing: true }));
      const missing = EXPECTED_PERIODS.filter((p) => !byPeriod.has(p));
      const duplicate = Array.from(byPeriod).filter(([, values]) => values.length > 1).map(([p]) => p);
      const unexpected = Array.from(byPeriod.keys()).filter((p) => !EXPECTED_PERIODS.includes(p));
      return { name, records, rows: rowsOrdered, missing, duplicate, unexpected, valid: !missing.length && !duplicate.length && !unexpected.length };
    });
    return students.sort((a, b) => a.name.localeCompare(b.name));
  }

  function renderWorkspace() {
    studentSelect.innerHTML = state.students.map((s, i) => `<option value="${i}">${escapeHtml(s.name)}${s.valid ? '' : ' ⚠'}</option>`).join('');
    const valid = state.students.filter((s) => s.valid).length;
    const flagged = state.students.length - valid;
    stats.innerHTML = `
      <div class="stat"><strong>${state.students.length}</strong><span>student names found</span></div>
      <div class="stat"><strong>${valid}</strong><span>complete schedules</span></div>
      <div class="stat"><strong>${flagged}</strong><span>need review</span></div>`;

    state.issues = state.students.filter((s) => !s.valid);
    if (state.issues.length) {
      warningPanel.classList.remove('is-hidden');
      warningPanel.innerHTML = `<strong>${state.issues.length} schedules need review.</strong> PDFs can still be generated; missing periods will be visibly marked.
        <details><summary>Show flagged students</summary><ul>${state.issues.map((s) => {
          const notes = [];
          if (s.missing.length) notes.push(`missing ${s.missing.join(', ')}`);
          if (s.duplicate.length) notes.push(`duplicate ${s.duplicate.join(', ')}`);
          if (s.unexpected.length) notes.push(`unexpected ${s.unexpected.join(', ')}`);
          return `<li><b>${escapeHtml(s.name)}</b>: ${escapeHtml(notes.join('; '))}</li>`;
        }).join('')}</ul></details>`;
    } else warningPanel.classList.add('is-hidden');
    workspace.classList.remove('is-hidden');
    renderPreview();
  }

  function renderPreview() {
    const student = state.students[Number(studentSelect.value) || 0];
    if (!student) return;
    preview.innerHTML = `
      <h3 class="preview-title">${escapeHtml(eventTitle.value)}</h3>
      <div class="preview-name">Student Schedule: ${escapeHtml(student.name)}</div>
      <table class="schedule-table">
        <thead><tr><th>Time</th><th>Period</th><th>Class/Course</th><th>Teacher</th><th>Location</th></tr></thead>
        <tbody>${student.rows.map((r) => `<tr class="${r.missing ? 'missing-row' : ''}">
          <td>${escapeHtml(r.time || 'Missing')}</td><td>${escapeHtml(r.period)}</td><td>${escapeHtml(r.course || 'No course found')}</td><td>${escapeHtml(r.teacher)}</td><td>${escapeHtml(r.location)}</td>
        </tr>`).join('')}</tbody>
      </table>`;
  }

  function wrapText(text, font, size, maxWidth, maxLines = 3) {
    const words = clean(text).split(' ').filter(Boolean);
    if (!words.length) return [''];
    const lines = [];
    let current = '';
    for (const word of words) {
      const test = current ? `${current} ${word}` : word;
      if (font.widthOfTextAtSize(test, size) <= maxWidth || !current) current = test;
      else { lines.push(current); current = word; }
    }
    if (current) lines.push(current);
    if (lines.length > maxLines) {
      const kept = lines.slice(0, maxLines);
      let last = kept[maxLines - 1];
      while (last && font.widthOfTextAtSize(`${last}...`, size) > maxWidth) last = last.slice(0, -1);
      kept[maxLines - 1] = `${last}...`;
      return kept;
    }
    return lines;
  }

  function drawCentered(page, text, font, size, y, color) {
    const safeText = clean(text);
    const width = font.widthOfTextAtSize(safeText, size);
    page.drawText(safeText, { x: (612 - width) / 2, y, size, font, color });
  }

  async function createPdf(students) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const doc = await PDFDocument.create();
    doc.setTitle(eventTitle.value || 'Student Schedules');
    doc.setAuthor('Cape Fear Academy');
    const regular = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const navy = rgb(11 / 255, 35 / 255, 70 / 255);
    const gold = rgb(201 / 255, 149 / 255, 33 / 255);
    const white = rgb(1, 1, 1);
    const ink = rgb(30 / 255, 36 / 255, 48 / 255);
    const muted = rgb(94 / 255, 102 / 255, 116 / 255);
    const stripe = rgb(244 / 255, 246 / 255, 249 / 255);
    const missingFill = rgb(1, .94, .93);
    const line = rgb(.78, .81, .85);
    const x0 = 38;
    const widths = [72, 60, 220, 116, 68];
    const headers = ['Time', 'Period', 'Class/Course', 'Teacher', 'Location'];

    for (const student of students) {
      const page = doc.addPage([612, 792]);
      drawCentered(page, eventTitle.value, bold, 19, 724, navy);
      drawCentered(page, `Student Schedule: ${student.name}`, bold, 13, 690, gold);
      let y = 646;
      const headerH = 30;
      let x = x0;
      for (let i = 0; i < widths.length; i++) {
        page.drawRectangle({ x, y, width: widths[i], height: headerH, color: navy, borderColor: white, borderWidth: .5 });
        page.drawText(headers[i], { x: x + 7, y: y + 10, font: bold, size: 8.3, color: white });
        x += widths[i];
      }
      y -= 1;

      for (let r = 0; r < student.rows.length; r++) {
        const row = student.rows[r];
        const values = [row.time || 'Missing', row.period, row.course || 'No course found', row.teacher, row.location];
        const wrapped = values.map((v, i) => wrapText(v, regular, 8.1, widths[i] - 14, i === 2 ? 3 : 2));
        const maxLines = Math.max(...wrapped.map((lines) => lines.length));
        const rowH = Math.max(42, maxLines * 10 + 16);
        y -= rowH;
        x = x0;
        for (let i = 0; i < widths.length; i++) {
          page.drawRectangle({ x, y, width: widths[i], height: rowH, color: row.missing ? missingFill : (r % 2 ? stripe : white), borderColor: line, borderWidth: .55 });
          const chosenFont = i === 0 || i === 4 ? bold : regular;
          const chosenColor = row.missing ? rgb(.58, .18, .15) : (i === 1 ? muted : ink);
          wrapped[i].forEach((text, lineNo) => page.drawText(text, { x: x + 7, y: y + rowH - 15 - lineNo * 10, font: chosenFont, size: 8.1, color: chosenColor }));
          x += widths[i];
        }
      }
      page.drawText('Generated from the uploaded schedule spreadsheet', { x: 38, y: 34, size: 7.3, font: regular, color: muted });
    }
    return doc.save();
  }

  function safeFilename(name) {
    return clean(name).replace(/[^a-z0-9]+/gi, '_').replace(/^_|_$/g, '') || 'Student';
  }

  function downloadBlob(bytes, filename, type) {
    const blob = bytes instanceof Blob ? bytes : new Blob([bytes], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function withBusy(button, message, task) {
    const old = button.textContent;
    button.disabled = true;
    button.textContent = message;
    downloadNote.textContent = '';
    try { await task(); }
    catch (error) { downloadNote.textContent = `Could not generate the file: ${error.message}`; }
    finally { button.disabled = false; button.textContent = old; }
  }

  async function handleFile(file) {
    if (!file) return;
    fileStatus.classList.remove('error');
    fileStatus.textContent = `Reading ${file.name}…`;
    workspace.classList.add('is-hidden');
    try {
      const rows = await parseXlsx(file);
      state.students = buildStudents(rows);
      if (!state.students.length) throw new Error('No student schedule rows were found.');
      state.sourceName = file.name;
      fileStatus.textContent = `${file.name} loaded successfully.`;
      renderWorkspace();
      workspace.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (error) {
      fileStatus.classList.add('error');
      fileStatus.textContent = `Could not read this spreadsheet: ${error.message}`;
    }
  }

  fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));
  for (const eventName of ['dragenter', 'dragover']) dropZone.addEventListener(eventName, (e) => { e.preventDefault(); dropZone.classList.add('is-dragging'); });
  for (const eventName of ['dragleave', 'drop']) dropZone.addEventListener(eventName, (e) => { e.preventDefault(); dropZone.classList.remove('is-dragging'); });
  dropZone.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
  studentSelect.addEventListener('change', renderPreview);
  eventTitle.addEventListener('input', renderPreview);

  el('downloadSelected').addEventListener('click', (e) => withBusy(e.currentTarget, 'Generating…', async () => {
    const student = state.students[Number(studentSelect.value) || 0];
    const pdf = await createPdf([student]);
    downloadBlob(pdf, `${safeFilename(student.name)}_Schedule.pdf`, 'application/pdf');
    downloadNote.textContent = `Downloaded ${student.name}'s schedule.`;
  }));

  el('downloadCombined').addEventListener('click', (e) => withBusy(e.currentTarget, 'Building PDF…', async () => {
    const pdf = await createPdf(state.students);
    downloadBlob(pdf, 'Curriculum_Night_Student_Schedules.pdf', 'application/pdf');
    downloadNote.textContent = `Downloaded one ${state.students.length}-page PDF.`;
  }));

  el('downloadZip').addEventListener('click', (e) => withBusy(e.currentTarget, 'Building ZIP…', async () => {
    const zip = new JSZip();
    for (let i = 0; i < state.students.length; i++) {
      const student = state.students[i];
      e.currentTarget.textContent = `PDF ${i + 1} of ${state.students.length}…`;
      zip.file(`${safeFilename(student.name)}_Schedule.pdf`, await createPdf([student]));
      if (i % 12 === 0) await new Promise((resolve) => setTimeout(resolve, 0));
    }
    const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    downloadBlob(blob, 'Curriculum_Night_Individual_Schedules.zip', 'application/zip');
    downloadNote.textContent = `Downloaded ${state.students.length} individual PDFs in one ZIP file.`;
  }));
})();
