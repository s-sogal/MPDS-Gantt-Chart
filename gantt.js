// ─────────────────────────────────────────
// Tournament Gantt Planner — gantt.js
// ─────────────────────────────────────────

// ── Constants ─────────────────────────────

const NEEDS_BOOKING = ['amtrak', 'fly', 'plane', 'flight', 'airplane'];
const NEEDS_CAR     = ['car', 'drive', 'van'];

const GUESSES = {
  name:      ['name', 'tournament', 'event', 'title', 'competition'],
  date:      ['date', 'when', 'day', 'datetime'],
  loc:       ['location', 'city', 'venue', 'place', 'site', 'where', 'address'],
  transport: ['transport', 'travel', 'mode', 'vehicle', 'transit', 'transportation'],
  debaters:  ['debater', 'athlete', 'student', 'participant', 'count', 'num', 'size'],
};

// ── State ─────────────────────────────────

let parsedRows    = [];
let columnHeaders = [];
let ganttTasks    = [];
let ganttChart    = null;

// ── Budget inputs ─────────────────────────

function onTimelineSettingsChange() {
  if (window._tournamentData && window._tournamentData.length) recomputeTasks();
}

function onBudgetChange() {
  onTimelineSettingsChange();
}

function getBudgetSettings() {
  const confirmDays     = readLeadDays('confirm-lead',     7);
  const confirmDuration = readLeadDays('confirm-duration', 3);
  const bookingDays     = readLeadDays('booking-lead',     14);
  const bookingDuration = readLeadDays('booking-duration', 5);
  const budgets = [1, 2, 3]
    .map(n => {
      const deadlineVal = document.getElementById(`budget-deadline-${n}`).value;
      const leadDays    = readLeadDays(`budget-lead-${n}`, n === 1 ? 14 : n === 2 ? 7 : 3);
      const deadline    = deadlineVal ? new Date(deadlineVal + 'T12:00:00') : null;
      return { deadline, leadDays, num: n };
    })
    .filter(s => s.deadline !== null)
    .sort((a, b) => a.deadline - b.deadline); // chronological so staggering works
  return { confirmDays, confirmDuration, bookingDays, bookingDuration, budgets };
}

function readLeadDays(id, fallback) {
  const input = document.getElementById(id);
  const value = parseInt(input.value, 10);
  if (!Number.isFinite(value) || value < 1) return fallback;
  const max = parseInt(input.max, 10);
  return Number.isFinite(max) ? Math.min(value, max) : value;
}

// ── Drag & drop ───────────────────────────

const dz = document.getElementById('dropzone');
dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('dragover'); });
dz.addEventListener('dragleave', ()  => dz.classList.remove('dragover'));
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.classList.remove('dragover');
  if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
});

function handleFileInput(input) {
  if (input.files[0]) processFile(input.files[0]);
}

// ── File parsing ──────────────────────────

function processFile(file) {
  // Reset input so re-selecting the same file fires onchange again
  document.getElementById('file-input').value = '';

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      if (!ws) { showStatus('No sheet found in file.', true); return; }

      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!data.length || !data[0].length) { showStatus('File appears empty.', true); return; }

      columnHeaders = data[0].map(String);
      parsedRows    = data.slice(1).filter(r => r.some(c => c !== ''));

      if (!parsedRows.length) { showStatus('No data rows found — check the file.', true); return; }

      document.getElementById('file-banner').classList.add('visible');
      document.getElementById('file-name-display').textContent = file.name;
      document.getElementById('file-rows-display').textContent =
        `${parsedRows.length} tournament${parsedRows.length !== 1 ? 's' : ''} found`;

      document.getElementById('format-hint').style.display = 'none';
      document.getElementById('preview-section').classList.remove('visible');
      clearGanttChart();

      populateMapper(columnHeaders);
      document.getElementById('mapper-section').classList.add('visible');
      autoApplyMapping();
    } catch (err) {
      console.error('File parse error:', err);
      showStatus(`Could not read file: ${err.message || 'unsupported format'}`, true);
    }
  };
  reader.onerror = () => showStatus('File could not be read.', true);
  reader.readAsArrayBuffer(file);
}

// ── Column mapper ─────────────────────────

function populateMapper(headers) {
  ['map-name','map-date','map-loc','map-transport','map-debaters'].forEach(id => {
    const sel = document.getElementById(id);
    sel.innerHTML = '<option value="">— Not mapped —</option>';
    headers.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value = i; opt.textContent = h;
      sel.appendChild(opt);
    });
  });
}

function bestGuess(headers, keys) {
  for (const k of keys) {
    const idx = headers.findIndex(h => h.toLowerCase().includes(k));
    if (idx !== -1) return String(idx);
  }
  return '';
}

function autoApplyMapping() {
  const h = columnHeaders.map(x => x.toLowerCase());
  document.getElementById('map-name').value      = bestGuess(h, GUESSES.name);
  document.getElementById('map-date').value      = bestGuess(h, GUESSES.date);
  document.getElementById('map-loc').value       = bestGuess(h, GUESSES.loc);
  document.getElementById('map-transport').value = bestGuess(h, GUESSES.transport);
  document.getElementById('map-debaters').value  = bestGuess(h, GUESSES.debaters);
  applyMapping(true);
}

function applyMapping(silent) {
  const nameIdx  = document.getElementById('map-name').value;
  const dateIdx  = document.getElementById('map-date').value;
  const locIdx   = document.getElementById('map-loc').value;
  const transIdx = document.getElementById('map-transport').value;
  const debIdx   = document.getElementById('map-debaters').value;

  if (nameIdx === '' || dateIdx === '') {
    if (!silent) showStatus('Please map at least the Name and Date columns.', true);
    return;
  }

  const ni = parseInt(nameIdx);
  const di = parseInt(dateIdx);
  const li = locIdx   !== '' ? parseInt(locIdx)   : null;
  const ti = transIdx !== '' ? parseInt(transIdx) : null;
  const xi = debIdx   !== '' ? parseInt(debIdx)   : null;

  const warnings = [];
  window._tournamentData = [];

  parsedRows.forEach((row, ri) => {
    const name      = String(row[ni] || '').trim();
    const rawDate   = row[di];
    const loc       = li !== null ? String(row[li] || '').trim() : '';
    const transport = ti !== null ? String(row[ti] || '').toLowerCase().trim() : '';
    const debaters  = xi !== null ? (parseInt(row[xi]) || null) : null;

    if (!name) return;

    const date = parseDate(rawDate);
    if (!date) warnings.push(`Row ${ri+2} ("${name}"): could not parse date "${rawDate}"`);

    window._tournamentData.push({ name, date, loc, transport, debaters });
  });

  recomputeTasks(warnings);
  if (!silent) showStatus(`Loaded ${window._tournamentData.length} tournament${window._tournamentData.length !== 1 ? 's' : ''}`);
}

// ── Task computation ──────────────────────

function recomputeTasks(warnings) {
  warnings = warnings || [];
  const { confirmDays, confirmDuration, bookingDays, bookingDuration, budgets } = getBudgetSettings();
  const data = window._tournamentData || [];

  ganttTasks = [];

  data.forEach(t => {
    const { name, date, loc, transport, debaters } = t;
    if (!date) return;

    // ① Confirm teams: bar ends at dueDate (X days before), starts confirmDuration days earlier
    const confirmDue   = offsetDays(date, -confirmDays);
    const confirmStart = offsetDays(confirmDue, -confirmDuration);
    addGanttTask({
      name, loc, transport, debaters, tournDate: date,
      task:       'Confirm team members',
      dueDate:    confirmDue,
      barStart:   confirmStart,
      barEnd:     confirmDue,
      daysBefore: daysBetween(confirmDue, date),
      type:       'confirm',
    });

    // ② Book transport: same bar logic with booking lead + duration
    if (NEEDS_BOOKING.includes(transport) || NEEDS_CAR.includes(transport)) {
      const taskLabel  = NEEDS_BOOKING.includes(transport) ? `Book ${transport} tickets` : 'Reserve rental car / van';
      const bookDue    = offsetDays(date, -bookingDays);
      const bookStart  = offsetDays(bookDue, -bookingDuration);
      addGanttTask({
        name, loc, transport, debaters, tournDate: date,
        task:       taskLabel,
        dueDate:    bookDue,
        barStart:   bookStart,
        barEnd:     bookDue,
        daysBefore: daysBetween(bookDue, date),
        type:       'book',
      });
    }

    // ③ Budget request — staggered: earliest deadline after tournament date, or last deadline if all have passed
    const assignedBudget = budgets.find(b => date < b.deadline) || budgets[budgets.length - 1];
    if (assignedBudget) {
      const { deadline, leadDays, num } = assignedBudget;
      const budgetDue = offsetDays(deadline, -leadDays);
      addGanttTask({
        name, loc, transport, debaters, tournDate: date,
        task:       `Submit budget request (DL ${num})`,
        dueDate:    budgetDue,
        barStart:   budgetDue,
        barEnd:     deadline,
        daysBefore: daysBetween(budgetDue, date),
        type:       'budget',
        budgetNum:  num,
        deadline,
      });
    }
  });

  renderPreview(warnings);
}

function addGanttTask(task) {
  ganttTasks.push({
    ...task,
    dueDate:   new Date(task.dueDate),
    tournDate: new Date(task.tournDate),
    barStart:  task.barStart ? new Date(task.barStart) : null,
    barEnd:    task.barEnd   ? new Date(task.barEnd)   : null,
    deadline:  task.deadline ? new Date(task.deadline) : null,
  });
}

// ── Date helpers ──────────────────────────

function parseDate(raw) {
  if (raw instanceof Date && !isNaN(raw)) return raw;
  if (raw === '' || raw === null || raw === undefined) return null;
  if (typeof raw === 'string') {
    const trimmed = raw.trim();
    const d = new Date(trimmed);
    if (!isNaN(d)) return d;
    const num = parseFloat(trimmed);
    if (!isNaN(num) && /^\d+(\.\d+)?$/.test(trimmed) && num > 25569)
      return new Date(Math.round((num - 25569) * 86400 * 1000));
    return null;
  }
  if (typeof raw === 'number' && !isNaN(raw) && raw > 25569)
    return new Date(Math.round((raw - 25569) * 86400 * 1000));
  return null;
}

function offsetMonths(d, m) {
  const r = new Date(d); r.setMonth(r.getMonth() + m); return r;
}
function offsetDays(d, days) {
  const r = new Date(d);
  r.setDate(r.getDate() + days);
  return r;
}
function daysBetween(a, b) {
  return Math.round((b - a) / 86400000);
}
function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}
function taskBarStartDate(task) { return task.barStart || task.dueDate; }
function taskBarEndDate(task)   { return task.barEnd   || task.dueDate; }

// ── Preview rendering ─────────────────────

function renderPreview(warnings) {
  const tbody   = document.getElementById('preview-body');
  const warnDiv = document.getElementById('warnings');
  document.getElementById('preview-section').classList.add('visible');

  if (!ganttTasks.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="5">No tasks generated — check column mapping and dates.</td></tr>';
    document.getElementById('row-count').textContent = '0 tasks';
    warnDiv.innerHTML = '';
    clearGanttChart();
    return;
  }

  let lastName = null;
  const rows = ganttTasks.map(t => {
    const isFirst = t.name !== lastName;
    lastName = t.name;

    const nameCell = isFirst
      ? `<div class="tourn-name">${esc(t.name)}</div>${t.loc ? `<div class="tourn-loc">&#128205; ${esc(t.loc)}</div>` : ''}`
      : '';

    const debatersCell = (isFirst && t.debaters)
      ? `<span class="debater-badge">${t.debaters} debater${t.debaters !== 1 ? 's' : ''}</span>`
      : '';

    const pillClass = t.type === 'budget'
      ? `task-budget task-budget-${t.budgetNum || 1}`
      : ({ confirm:'task-confirm', book:'task-book' }[t.type] || 'task-confirm');

    return `<tr>
      <td>${nameCell}</td>
      <td>${debatersCell}</td>
      <td><span class="task-pill ${pillClass}"><span class="task-dot"></span>${esc(t.task)}</span></td>
      <td><span class="date-val">${fmtDate(t.dueDate)}</span></td>
      <td>${t.daysBefore > 0
        ? `<span class="ahead-val">${t.daysBefore}</span><span class="days-lbl"> days before</span>`
        : `<span style="color:var(--muted)">—</span>`}</td>
    </tr>`;
  });

  tbody.innerHTML = rows.join('');
  document.getElementById('row-count').textContent =
    `${ganttTasks.length} task${ganttTasks.length !== 1 ? 's' : ''}`;

  warnDiv.innerHTML = warnings.length
    ? warnings.map(w => `
        <div class="warning-item">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
            <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
          </svg>${esc(w)}
        </div>`).join('')
    : '';

  renderGanttChart();
}

// ── Highcharts Gantt chart ─────────────────

function renderGanttChart() {
  const section = document.getElementById('chart-section');

  if (!ganttTasks.length) {
    clearGanttChart();
    return;
  }

  const tournSeen = new Set();
  const tournamentNames = [];
  ganttTasks.forEach(t => {
    if (!tournSeen.has(t.name)) { tournSeen.add(t.name); tournamentNames.push(t.name); }
  });

  const { budgets } = getBudgetSettings();
  const BUDGET_CHART = { 1: '#f0b840', 2: '#e07824', 3: '#d44030' };

  const confirmData = [], bookData = [], milestones = [];
  const budgetData  = { 1: [], 2: [], 3: [] };
  const mileSeen    = new Set();

  ganttTasks.forEach(t => {
    const y = tournamentNames.indexOf(t.name);
    const start = taskBarStartDate(t).getTime();
    const end   = taskBarEndDate(t).getTime();
    if (t.type === 'confirm') {
      confirmData.push({ name: t.task, start, end, y });
    } else if (t.type === 'book') {
      bookData.push({ name: t.task, start, end, y });
    } else if (t.type === 'budget' && t.deadline) {
      (budgetData[t.budgetNum] || budgetData[1]).push({ name: t.task, start, end, y });
    }
    if (t.tournDate && !mileSeen.has(t.name)) {
      mileSeen.add(t.name);
      milestones.push({ name: t.name, start: t.tournDate.getTime(), end: t.tournDate.getTime(), y, milestone: true });
    }
  });

  const plotLines = [{
    value: Date.now(), color: '#e8ff47', width: 1.5, dashStyle: 'ShortDash', zIndex: 5,
    label: { text: 'Today', align: 'left', style: { color: '#e8ff47', fontSize: '10px', fontFamily: "'DM Mono', monospace" } },
  }];
  budgets.forEach(({ deadline, num }) => {
    plotLines.push({
      value: deadline.getTime(), color: BUDGET_CHART[num], width: 1.5, dashStyle: 'ShortDash', zIndex: 5,
      label: { text: `Budget DL ${num}`, align: 'left', style: { color: BUDGET_CHART[num], fontSize: '10px', fontFamily: "'DM Mono', monospace" } },
    });
  });

  const rowH   = 48;
  const height = Math.max(340, tournamentNames.length * rowH + 180);
  document.getElementById('gantt-container').style.height = height + 'px';

  if (ganttChart) { ganttChart.destroy(); ganttChart = null; }
  section.style.display = 'block';

  ganttChart = Highcharts.ganttChart('gantt-container', {
    chart: {
      backgroundColor: '#13161d',
      style: { fontFamily: "'DM Mono', monospace" },
      borderRadius: 14,
      spacingTop: 24,
      spacingBottom: 16,
    },
    title: {
      text: 'Tournament Prep Timeline', align: 'left',
      style: { color: '#dde2ee', fontFamily: "'Syne', sans-serif", fontWeight: '700', fontSize: '15px' },
    },
    subtitle: {
      text: `${tournamentNames.length} tournament${tournamentNames.length !== 1 ? 's' : ''} · ${ganttTasks.length} task${ganttTasks.length !== 1 ? 's' : ''}`,
      align: 'left',
      style: { color: '#5a6278', fontFamily: "'DM Mono', monospace", fontSize: '11px' },
    },
    xAxis: {
      plotLines,
      dateTimeLabelFormats: { week: "%b %e", month: "%b '%y" },
      gridLineColor: '#252a38', lineColor: '#252a38', tickColor: '#252a38',
      labels: { style: { color: '#8891a8', fontSize: '10px' } },
    },
    yAxis: {
      type: 'category',
      categories: tournamentNames,
      gridLineColor: '#252a38',
      labels: { style: { color: '#dde2ee', fontSize: '11px', fontWeight: '600' } },
      plotBands: tournamentNames.map((_, i) => ({
        from: i - 0.5, to: i + 0.5,
        color: i % 2 === 0 ? 'rgba(255,255,255,0.025)' : 'rgba(0,0,0,0)',
      })),
    },
    navigator: {
      enabled: tournamentNames.length > 6,
      handles: { backgroundColor: '#2f3547', borderColor: '#5a6278' },
      maskFill: 'rgba(71,200,255,0.06)',
      series: { color: '#2E75B6', type: 'gantt' },
      xAxis: { gridLineColor: '#252a38', labels: { style: { color: '#5a6278' } } },
    },
    scrollbar:     { enabled: tournamentNames.length > 6 },
    rangeSelector: { enabled: false },
    legend: {
      enabled: true,
      backgroundColor: '#1a1e28', borderColor: '#252a38', borderWidth: 1, borderRadius: 8,
      itemStyle:      { color: '#8891a8', fontFamily: "'DM Mono', monospace", fontSize: '11px' },
      itemHoverStyle: { color: '#dde2ee' },
    },
    tooltip: {
      outside: true,
      backgroundColor: '#1a1e28', borderColor: '#2f3547', borderRadius: 8,
      style: { color: '#dde2ee', fontFamily: "'DM Mono', monospace", fontSize: '11px' },
      formatter: function () {
        const p = this.point;
        const s = Highcharts.dateFormat('%b %e, %Y', p.start);
        const e = Highcharts.dateFormat('%b %e, %Y', p.end);
        const hdr = `<span style="color:${this.color}">■</span> <b>${this.series.name}</b><br/>`;
        if (p.milestone) return hdr + `<b>${p.name}</b><br/>${s}`;
        return hdr + `${p.name}<br/>From: ${s}<br/>To: ${e}`;
      },
    },
    series: [
      {
        name: 'Confirm Teams', color: '#4a8fd4', borderColor: 'rgba(74,143,212,0.5)', borderRadius: 4,
        dataLabels: { enabled: true, format: '{point.name}', align: 'left', padding: 6,
          style: { color: '#fff', textOutline: 'none', fontSize: '10px', fontWeight: '500' } },
        data: confirmData,
      },
      {
        name: 'Book Transport', color: '#5ec27a', borderColor: 'rgba(94,194,122,0.5)', borderRadius: 4,
        dataLabels: { enabled: true, format: '{point.name}', align: 'left', padding: 6,
          style: { color: '#fff', textOutline: 'none', fontSize: '10px', fontWeight: '500' } },
        data: bookData,
      },
      ...[1, 2, 3].filter(n => budgetData[n].length).map(n => ({
        name: `Budget DL ${n}`, color: BUDGET_CHART[n],
        borderColor: BUDGET_CHART[n] + '80', borderRadius: 4,
        dataLabels: { enabled: true, format: '{point.name}', align: 'left', padding: 6,
          style: { color: n === 1 ? '#1a1000' : '#fff', textOutline: 'none', fontSize: '10px', fontWeight: '500' } },
        data: budgetData[n],
      })),
      {
        name: 'Tournament Date', color: '#e8ff47', marker: { symbol: 'diamond' },
        data: milestones,
      },
    ],
    exporting: {
      enabled: true, allowHTML: true,
      sourceWidth: 1400, sourceHeight: height + 60,
      filename: 'tournament_gantt',
      chartOptions: { chart: { backgroundColor: '#13161d' } },
      buttons: {
        contextButton: {
          theme: {
            fill: '#1a1e28', stroke: '#2f3547', 'stroke-width': 1, r: 6,
            style: { color: '#8891a8' },
            states: { hover: { fill: '#252a38' } },
          },
        },
      },
    },
    credits: { enabled: false },
  });

  ganttChart.reflow();
}

function clearGanttChart() {
  if (ganttChart) { ganttChart.destroy(); ganttChart = null; }
  const section = document.getElementById('chart-section');
  const container = document.getElementById('gantt-container');
  section.style.display = 'none';
  container.innerHTML = '';
  container.removeAttribute('data-highcharts-chart');
}

// ── Highcharts export helpers ─────────────────────────────────────────────

function exportGanttPDF() {
  if (!ganttChart) { showStatus('No chart to export yet.', true); return; }
  ganttChart.exportChartLocal({ type: 'application/pdf', filename: 'tournament_gantt' });
  showStatus('Exporting PDF…');
}

// ── Excel Gantt export (client-side, styled) ──────────────────────────────

const C = {
  NAVY:'1F3864', BLUE:'2E75B6', BLUE_LIGHT:'D6E4F0',
  GREEN:'70AD47', GOLD:'BF8F00', GOLD_LIGHT:'FFF2CC',
  BAR_CONFIRM:'2E75B6', BAR_BOOK:'1F5C2E', BAR_BUDGET:'BF8F00',
  BAR_BUDGET_1:'BF8F00', BAR_BUDGET_2:'C06000', BAR_BUDGET_3:'903020',
  ROW_ALT:'F2F7FB', WHITE:'FFFFFF', GRID:'BDD7EE', TEXT:'1F3864', RED:'C00000',
};
const BUDGET_BAR_COLORS = { 1: 'BF8F00', 2: 'C06000', 3: '903020' };

function xfFill(hex){ return { patternType:'solid', fgColor:{rgb:hex} }; }
function xfFont(o){   return { name:'Arial', sz:o.sz||9, bold:!!o.bold, italic:!!o.italic, color:{rgb:o.color||C.TEXT} }; }
function xfAlign(h,v){ return { horizontal:h||'left', vertical:v||'center', wrapText:true }; }
function xfBorder(c){ const s={style:'thin',color:{rgb:c||C.GRID}}; return {top:s,bottom:s,left:s,right:s}; }
function cSt(fill,font,align,border){ return { fill:xfFill(fill), font, alignment:align, border:border||xfBorder() }; }

function colLetter(i){ let s='',n=i+1; while(n>0){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);} return s; }
function R(r,c){ return `${colLetter(c)}${r+1}`; }

function exportExcel() {
  if (!ganttTasks.length) { showStatus('Nothing to export yet.', true); return; }

  const { budgets: budgetSettings } = getBudgetSettings();
  const deadlines = budgetSettings.map(s => s.deadline).filter(Boolean);
  const tournMap = new Map();
  ganttTasks.forEach(t => { if (!tournMap.has(t.name)) tournMap.set(t.name, t); });
  const tournaments = [...tournMap.values()];

  const allDates = ganttTasks.flatMap(t => [taskBarStartDate(t), t.dueDate, taskBarEndDate(t), t.tournDate].filter(Boolean));
  if (!allDates.length) { showStatus('No valid dates found.', true); return; }

  let cs0 = new Date(Math.min(...allDates));
  let ce0 = new Date(Math.max(...allDates));
  cs0.setDate(cs0.getDate() - 14 - ((cs0.getDay()+6)%7));
  ce0.setDate(ce0.getDate() + 21);
  ce0.setDate(ce0.getDate() + (7-(ce0.getDay()+6)%7)%7);

  const mondays = [];
  for(let d=new Date(cs0); d<=ce0; d.setDate(d.getDate()+7)) mondays.push(new Date(d));

  const today = new Date(); today.setHours(0,0,0,0);
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  const COL_TOURN=0, COL_TASK=1, COL_DEB=2, COL_DUE=3, COL_D0=4;
  const ROW_TITLE=0, ROW_QTR=1, ROW_MON=2, ROW_WEEK=3, ROW_DATA=4;
  const TOTAL = COL_D0 + mondays.length;

  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];

  function sc(r,c,v,style){ ws[R(r,c)] = { v, t:typeof v==='number'?'n':'s', s:style }; }

  // Title row
  for(let c=0;c<TOTAL;c++)
    sc(ROW_TITLE,c,c===0?'Tournament Prep — Gantt Chart':'',cSt(C.NAVY,xfFont({sz:14,bold:true,color:C.WHITE}),xfAlign('left','center')));
  merges.push({s:{r:ROW_TITLE,c:0},e:{r:ROW_TITLE,c:TOTAL-1}});

  // Quarter/month spans
  const qSpans={},mSpans={};
  mondays.forEach((mon,i)=>{
    const dc=COL_D0+i;
    const q=`Q${Math.floor(mon.getMonth()/3)+1} ${mon.getFullYear()}`;
    const mk=`${mon.getFullYear()}-${mon.getMonth()}`;
    const ml=`${MONTHS[mon.getMonth()]} ${mon.getFullYear()}`;
    qSpans[q]=qSpans[q]?{...qSpans[q],max:dc}:{min:dc,max:dc,label:q};
    mSpans[mk]=mSpans[mk]?{...mSpans[mk],max:dc}:{min:dc,max:dc,label:ml};
  });

  // Left-side headers
  [ROW_QTR,ROW_MON,ROW_WEEK].forEach(r=>{
    const bg=r===ROW_QTR?C.NAVY:r===ROW_MON?C.BLUE:C.BLUE_LIGHT;
    ['Tournament','Task','Debaters','Due Date'].forEach((h,c)=>{
      sc(r,c,r===ROW_WEEK?h:'',cSt(bg,xfFont({sz:8,bold:true,color:r===ROW_WEEK?C.TEXT:C.WHITE}),xfAlign('center','center')));
    });
  });

  // Quarter row merges
  Object.values(qSpans).forEach(({min,max,label})=>{
    for(let c=min;c<=max;c++) sc(ROW_QTR,c,c===min?label:'',cSt(C.NAVY,xfFont({sz:8,bold:true,color:C.WHITE}),xfAlign('center','center')));
    if(min<max) merges.push({s:{r:ROW_QTR,c:min},e:{r:ROW_QTR,c:max}});
  });

  // Month row merges
  Object.values(mSpans).forEach(({min,max,label})=>{
    for(let c=min;c<=max;c++) sc(ROW_MON,c,c===min?label:'',cSt(C.BLUE,xfFont({sz:8,bold:true,color:C.WHITE}),xfAlign('center','center')));
    if(min<max) merges.push({s:{r:ROW_MON,c:min},e:{r:ROW_MON,c:max}});
  });

  // Week row
  mondays.forEach((mon,i)=>{
    const dc=COL_D0+i;
    const isTod=mon<=today && today<new Date(mon.getTime()+7*864e5);
    if(isTod) sc(ROW_QTR,dc,'TODAY',cSt(C.GREEN,xfFont({sz:7,bold:true,color:C.WHITE}),xfAlign('center','center')));
    sc(ROW_WEEK,dc,mon.getDate(),cSt(isTod?C.GREEN:C.BLUE_LIGHT,xfFont({sz:7,bold:isTod,color:isTod?C.WHITE:C.TEXT}),xfAlign('center','center')));
    // Budget deadline markers on month row
    deadlines.forEach((dl, di) => {
      if(dl && mon<=dl && dl<new Date(mon.getTime()+7*864e5))
        sc(ROW_MON,dc,`BUDGET DL ${di+1}`,cSt(BUDGET_BAR_COLORS[di+1]||C.GOLD,xfFont({sz:7,bold:true,color:C.WHITE}),xfAlign('center','center')));
    });
  });

  // Data rows
  let cr=ROW_DATA;
  const TCOLORS = t => t.type === 'budget'
    ? (BUDGET_BAR_COLORS[t.budgetNum] || C.BAR_BUDGET)
    : ({ confirm: C.BAR_CONFIRM, book: C.BAR_BOOK }[t.type] || C.BAR_CONFIRM);
  const TLABELS = t => t.type === 'budget'
    ? `Budget Request (DL ${t.budgetNum})`
    : ({ confirm: 'Confirm Team Members', book: 'Book Transport' }[t.type] || t.task);

  tournaments.forEach(tourn=>{
    const tTasks=ganttTasks.filter(t=>t.name===tourn.name);
    const phaseLabel=`${tourn.name}${tourn.loc?'   •   '+tourn.loc:''}${tourn.debaters?'   •   '+tourn.debaters+' debaters':''}`;

    // Phase header
    for(let c=0;c<TOTAL;c++){
      const isMk=tourn.tournDate&&mondays.some((m,i)=>COL_D0+i===c&&m<=tourn.tournDate&&tourn.tournDate<new Date(m.getTime()+7*864e5));
      sc(cr,c,c===0?phaseLabel:(isMk?'▲':''),cSt(isMk?C.GREEN:C.NAVY,isMk?xfFont({sz:7,bold:true,color:C.WHITE}):xfFont({sz:9,bold:true,color:C.WHITE}),xfAlign(c===0?'left':'center','center')));
    }
    merges.push({s:{r:cr,c:0},e:{r:cr,c:COL_D0-1}});
    cr++;

    tTasks.forEach((task,ti)=>{
      const alt=ti%2===1, bg=alt?C.ROW_ALT:C.WHITE;
      const barColor=TCOLORS(task);

      sc(cr,COL_TOURN, ti===0?tourn.name:'',  cSt(bg,xfFont({sz:8,bold:ti===0}),xfAlign('left','center')));
      sc(cr,COL_TASK,  task.task,               cSt(bg,xfFont({sz:8}),            xfAlign('left','center')));
      sc(cr,COL_DEB,   tourn.debaters||'',      cSt(bg,xfFont({sz:8}),            xfAlign('center','center')));
      sc(cr,COL_DUE,   task.dueDate?fmtDate(task.dueDate):'—', cSt(bg,xfFont({sz:8,italic:true}),xfAlign('center','center')));

      mondays.forEach((mon,i)=>{
        const dc=COL_D0+i, me=new Date(mon.getTime()+7*864e5);
        const bs=taskBarStartDate(task), be=taskBarEndDate(task);
        const inB=bs&&be&&mon<new Date(be.getTime()+7*864e5)&&me>bs;
        const isF=inB&&bs>=mon&&bs<me;
        const isTod=mon<=today&&today<me;
        const isMk=tourn.tournDate&&mon<=tourn.tournDate&&tourn.tournDate<me;
        const isBudgetDL=deadlines.some(dl=>dl&&mon<=dl&&dl<me);

        let fill=bg, lbl='', fnt=xfFont({sz:7});
        if(inB){
          fill=barColor;
          const txtColor = task.type==='budget'&&task.budgetNum===1 ? C.TEXT : C.WHITE;
          fnt=xfFont({sz:7,bold:true,color:txtColor});
          if(isF) lbl=TLABELS(task);
        } else if(isMk){ fill=C.GREEN; fnt=xfFont({sz:7,bold:true,color:C.WHITE}); }
        else if(isBudgetDL&&!inB){ fill=C.GOLD_LIGHT; }
        else if(isTod){ fill='E8F4EA'; }

        sc(cr,dc,lbl,cSt(fill,fnt,xfAlign('left','center')));
      });
      cr++;
    });

    // Spacer
    for(let c=0;c<TOTAL;c++) sc(cr,c,'',cSt('F0F0F0',xfFont({sz:4}),xfAlign('center','center')));
    cr++;
  });

  // Legend
  [[C.BAR_CONFIRM,'Confirm Team Members'],[C.BAR_BOOK,'Book Transport'],[C.BAR_BUDGET,'Submit Budget Request'],[C.GREEN,'Tournament Date']].forEach(([color,label],i)=>{
    const bc=i*3;
    sc(cr,bc,' ',cSt(color,xfFont({}),xfAlign('center','center')));
    sc(cr,bc+1,label,cSt(C.WHITE,xfFont({sz:8}),xfAlign('left','center')));
    sc(cr,bc+2,'',cSt(C.WHITE,xfFont({}),xfAlign('center','center')));
  });

  const colWidths=[{wch:22},{wch:26},{wch:10},{wch:13}];
  mondays.forEach(()=>colWidths.push({wch:3.8}));

  ws['!ref']    = `A1:${colLetter(TOTAL-1)}${cr+2}`;
  ws['!merges'] = merges;
  ws['!cols']   = colWidths;
  ws['!rows']   = [{hpt:28},{hpt:14},{hpt:14},{hpt:12}];
  ws['!freeze'] = {xSplit:COL_D0, ySplit:ROW_DATA};

  XLSX.utils.book_append_sheet(wb, ws, 'Gantt Chart');

  // Sheet 2: Summary
  const ws2={};
  const bdlHeaders = budgetSettings.map(s=>`Budget DL ${s.num} Request By`);
  const bdlDlHeaders = budgetSettings.map(s=>`Budget Deadline ${s.num}`);
  const h2=['#','Tournament','Date','Location','Transport','Debaters','Book By',...bdlHeaders,...bdlDlHeaders];
  h2.forEach((h,ci)=>{ ws2[R(0,ci)]={v:h,t:'s',s:cSt(C.NAVY,xfFont({sz:9,bold:true,color:C.WHITE}),xfAlign('center','center'))}; });

  const seen2=new Set(); let r2=1,idx2=1;
  ganttTasks.forEach(t=>{
    if(!seen2.has(t.name)){
      seen2.add(t.name);
      const bookTask = ganttTasks.find(x=>x.name===t.name&&x.type==='book');
      const bdlRequestCols = budgetSettings.map(s=>{
        const bt=ganttTasks.find(x=>x.name===t.name&&x.type==='budget'&&x.budgetNum===s.num);
        return bt?fmtDate(bt.dueDate):'N/A';
      });
      const bdlDateCols = budgetSettings.map(s=>fmtDate(s.deadline));
      const alt=idx2%2===0;
      [idx2,t.name,t.tournDate?fmtDate(t.tournDate):'',t.loc||'',t.transport||'',
       t.debaters||'',bookTask?fmtDate(bookTask.dueDate):'N/A',
       ...bdlRequestCols,...bdlDateCols,
      ].forEach((v,ci)=>{
        ws2[R(r2,ci)]={v:String(v),t:'s',s:cSt(alt?C.ROW_ALT:C.WHITE,xfFont({sz:9}),xfAlign(ci===0||ci>=5?'center':'left','center'))};
      });
      r2++;idx2++;
    }
  });
  const totalSumCols = 7 + budgetSettings.length * 2;
  ws2['!ref']=`A1:${colLetter(totalSumCols-1)}${r2}`;
  ws2['!cols']=[{wch:4},{wch:26},{wch:14},{wch:20},{wch:12},{wch:10},{wch:16},...budgetSettings.flatMap(()=>[{wch:18},{wch:16}])];
  XLSX.utils.book_append_sheet(wb,ws2,'Summary');

  XLSX.writeFile(wb,'tournament_gantt.xlsx');
  showStatus('Gantt chart exported!');
}

// ── Template download ─────────────────────

function downloadTemplate() {
  const d = (y,m,day) => new Date(y, m-1, day); // local date, avoids UTC offset shifts
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Tournament Name','Date','Location','Transport','Debaters'],
    ['Spring Invitational',   d(2026,4,15), 'Chicago, IL',      'fly',    8],
    ['Regional Championships',d(2026,5,20), 'Boston, MA',       'amtrak', 6],
    ['District Qualifier',    d(2026,6,10), 'Silver Spring, MD','car',    4],
    ['State Finals',          d(2026,7, 8), 'Richmond, VA',     'metro',  10],
  ]);
  // Apply date format to column B so Excel shows dates, not serial numbers
  for (let r = 1; r <= 4; r++) {
    const cell = ws[XLSX.utils.encode_cell({r, c:1})];
    if (cell) cell.z = 'yyyy-mm-dd';
  }
  ws['!cols']=[{wch:26},{wch:14},{wch:22},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb,ws,'Tournaments');
  XLSX.writeFile(wb,'tournament_template.xlsx');
}

// ── Reset ─────────────────────────────────

function clearFile() {
  parsedRows=[]; columnHeaders=[]; ganttTasks=[]; window._tournamentData=[];
  clearGanttChart();
  document.getElementById('file-banner').classList.remove('visible');
  document.getElementById('mapper-section').classList.remove('visible');
  document.getElementById('preview-section').classList.remove('visible');
  document.getElementById('format-hint').style.display='';
  document.getElementById('file-input').value='';
  document.getElementById('warnings').innerHTML='';
}

// ── Utilities ─────────────────────────────

function showStatus(msg, err) {
  const el=document.getElementById('status-msg');
  el.textContent=(err?'':'✓ ')+msg;
  el.className='status-msg show '+(err?'err':'ok');
  setTimeout(()=>el.classList.remove('show'),3500);
}

function esc(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
