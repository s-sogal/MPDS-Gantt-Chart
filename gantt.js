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

function onBudgetChange() {
  if (window._tournamentData && window._tournamentData.length) recomputeTasks();
}

function getBudgetSettings() {
  const deadlineVal = document.getElementById('budget-deadline').value;
  const leadDays    = parseInt(document.getElementById('budget-lead').value) || 7;
  const deadline    = deadlineVal ? new Date(deadlineVal + 'T12:00:00') : null;
  return { deadline, leadDays };
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
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      if (!data.length) { showStatus('File appears empty.', true); return; }

      columnHeaders = data[0].map(String);
      parsedRows    = data.slice(1).filter(r => r.some(c => c !== ''));

      document.getElementById('file-banner').classList.add('visible');
      document.getElementById('file-name-display').textContent = file.name;
      document.getElementById('file-rows-display').textContent =
        `${parsedRows.length} tournament${parsedRows.length !== 1 ? 's' : ''} found`;

      document.getElementById('format-hint').style.display = 'none';
      document.getElementById('preview-section').classList.remove('visible');

      populateMapper(columnHeaders);
      document.getElementById('mapper-section').classList.add('visible');
      autoApplyMapping();
    } catch (err) {
      showStatus('Could not read file — please use .xlsx, .xls, or .csv', true);
    }
  };
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
  const { deadline, leadDays } = getBudgetSettings();
  const data = window._tournamentData || [];

  ganttTasks = [];

  data.forEach(t => {
    const { name, date, loc, transport, debaters } = t;
    if (!date) return;

    // ① Confirm teams — 1 month before, always
    const confirmDate = offsetMonths(date, -1);
    ganttTasks.push({
      name, loc, transport, debaters, tournDate: date,
      task:       'Confirm teams being sent',
      dueDate:    confirmDate,
      daysBefore: daysBetween(confirmDate, date),
      type:       'confirm',
    });

    // ② Book transport — amtrak/fly 1 month before; car/van 2 weeks before
    if (NEEDS_BOOKING.includes(transport)) {
      const bookDate = offsetMonths(date, -1);
      ganttTasks.push({
        name, loc, transport, debaters, tournDate: date,
        task:       `Book ${transport} tickets`,
        dueDate:    bookDate,
        daysBefore: daysBetween(bookDate, date),
        type:       'book',
      });
    } else if (NEEDS_CAR.includes(transport)) {
      const bookDate = offsetDays(date, -14);
      ganttTasks.push({
        name, loc, transport, debaters, tournDate: date,
        task:       'Reserve rental car / van',
        dueDate:    bookDate,
        daysBefore: daysBetween(bookDate, date),
        type:       'book',
      });
    }

    // ③ Budget request — leadDays before the deadline
    if (deadline) {
      const budgetDate = offsetDays(deadline, -leadDays);
      ganttTasks.push({
        name, loc, transport, debaters, tournDate: date,
        task:       'Submit budget request',
        dueDate:    budgetDate,
        daysBefore: daysBetween(budgetDate, date),
        type:       'budget',
        deadline,
      });
    }
  });

  renderPreview(warnings);
}

// ── Date helpers ──────────────────────────

function parseDate(raw) {
  if (raw instanceof Date && !isNaN(raw)) return raw;
  if (raw === '' || raw === null || raw === undefined) return null;
  const num = parseFloat(raw);
  if (!isNaN(num) && num > 30000) return new Date(Math.round((num - 25569) * 86400 * 1000));
  const d = new Date(String(raw));
  return isNaN(d) ? null : d;
}

function offsetMonths(d, m) {
  const r = new Date(d); r.setMonth(r.getMonth() + m); return r;
}
function offsetDays(d, days) {
  return new Date(d.getTime() + days * 86400000);
}
function daysBetween(a, b) {
  return Math.round((b - a) / 86400000);
}
function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}
function taskBarEndDate(task) {
  return task.type === 'budget' && task.deadline ? task.deadline : task.tournDate;
}

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

    const pillClass = { confirm:'task-confirm', book:'task-book', budget:'task-budget' }[t.type] || 'task-confirm';

    return `<tr>
      <td>${nameCell}</td>
      <td>${debatersCell}</td>
      <td><span class="task-pill ${pillClass}"><span class="task-dot"></span>${esc(t.task)}</span></td>
      <td style="white-space:nowrap;color:var(--text)">${fmtDate(t.dueDate)}</td>
      <td>${t.daysBefore > 0
        ? `<span class="ahead-val">${t.daysBefore}</span> <span style="color:var(--muted)">days before</span>`
        : '<span style="color:var(--muted)">—</span>'}</td>
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

  // Unique tournaments in insertion order
  const tournSeen = new Set();
  const tournamentNames = [];
  ganttTasks.forEach(t => {
    if (!tournSeen.has(t.name)) { tournSeen.add(t.name); tournamentNames.push(t.name); }
  });

  const { deadline } = getBudgetSettings();

  // Build series data
  const confirmData = [], bookData = [], budgetData = [], milestones = [];
  const mileSeen = new Set();

  ganttTasks.forEach(t => {
    const y = tournamentNames.indexOf(t.name);
    if (t.type === 'confirm') {
      confirmData.push({ name: t.task, start: t.dueDate.getTime(), end: taskBarEndDate(t).getTime(), y });
    } else if (t.type === 'book') {
      bookData.push({ name: t.task, start: t.dueDate.getTime(), end: taskBarEndDate(t).getTime(), y });
    } else if (t.type === 'budget' && t.deadline) {
      budgetData.push({ name: t.task, start: t.dueDate.getTime(), end: taskBarEndDate(t).getTime(), y });
    }
    if (t.tournDate && !mileSeen.has(t.name)) {
      mileSeen.add(t.name);
      milestones.push({ name: t.name, start: t.tournDate.getTime(), end: t.tournDate.getTime(), y, milestone: true });
    }
  });

  // xAxis plot lines
  const plotLines = [{
    value: Date.now(), color: '#e8ff47', width: 1.5, dashStyle: 'ShortDash', zIndex: 5,
    label: { text: 'Today', align: 'left', style: { color: '#e8ff47', fontSize: '10px', fontFamily: "'DM Mono', monospace" } },
  }];
  if (deadline) {
    plotLines.push({
      value: deadline.getTime(), color: '#f0c040', width: 1.5, dashStyle: 'ShortDash', zIndex: 5,
      label: { text: 'Budget Deadline', align: 'left', style: { color: '#f0c040', fontSize: '10px', fontFamily: "'DM Mono', monospace" } },
    });
  }

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
      {
        name: 'Budget Request', color: '#f0c040', borderColor: 'rgba(240,192,64,0.5)', borderRadius: 4,
        dataLabels: { enabled: true, format: '{point.name}', align: 'left', padding: 6,
          style: { color: '#0b0d11', textOutline: 'none', fontSize: '10px', fontWeight: '500' } },
        data: budgetData,
      },
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
  document.getElementById('chart-section').style.display = 'none';
}

// ── Highcharts export helpers ─────────────────────────────────────────────

function exportGanttPNG() {
  if (!ganttChart) { showStatus('No chart to export yet.', true); return; }
  ganttChart.exportChartLocal({ type: 'image/png', filename: 'tournament_gantt', scale: 2 });
  showStatus('Exporting PNG…');
}

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
  ROW_ALT:'F2F7FB', WHITE:'FFFFFF', GRID:'BDD7EE', TEXT:'1F3864', RED:'C00000',
};

function xfFill(hex){ return { patternType:'solid', fgColor:{rgb:hex} }; }
function xfFont(o){   return { name:'Arial', sz:o.sz||9, bold:!!o.bold, italic:!!o.italic, color:{rgb:o.color||C.TEXT} }; }
function xfAlign(h,v){ return { horizontal:h||'left', vertical:v||'center', wrapText:true }; }
function xfBorder(c){ const s={style:'thin',color:{rgb:c||C.GRID}}; return {top:s,bottom:s,left:s,right:s}; }
function cSt(fill,font,align,border){ return { fill:xfFill(fill), font, alignment:align, border:border||xfBorder() }; }

function colLetter(i){ let s='',n=i+1; while(n>0){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);} return s; }
function R(r,c){ return `${colLetter(c)}${r+1}`; }

function exportExcel() {
  if (!ganttTasks.length) { showStatus('Nothing to export yet.', true); return; }

  const { deadline } = getBudgetSettings();
  const tournMap = new Map();
  ganttTasks.forEach(t => { if (!tournMap.has(t.name)) tournMap.set(t.name, t); });
  const tournaments = [...tournMap.values()];

  const allDates = ganttTasks.flatMap(t => [t.dueDate, taskBarEndDate(t), t.tournDate].filter(Boolean));
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
    // Budget deadline marker on month row
    if(deadline && mon<=deadline && deadline<new Date(mon.getTime()+7*864e5))
      sc(ROW_MON,dc,'BUDGET DL',cSt(C.GOLD,xfFont({sz:7,bold:true,color:C.WHITE}),xfAlign('center','center')));
  });

  // Data rows
  let cr=ROW_DATA;
  const TCOLORS={confirm:C.BAR_CONFIRM, book:C.BAR_BOOK, budget:C.BAR_BUDGET};
  const TLABELS={confirm:'Confirm Teams', book:'Book Transport', budget:'Budget Request'};

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
      const barColor=TCOLORS[task.type]||C.BAR_CONFIRM;

      sc(cr,COL_TOURN, ti===0?tourn.name:'',  cSt(bg,xfFont({sz:8,bold:ti===0}),xfAlign('left','center')));
      sc(cr,COL_TASK,  task.task,               cSt(bg,xfFont({sz:8}),            xfAlign('left','center')));
      sc(cr,COL_DEB,   tourn.debaters||'',      cSt(bg,xfFont({sz:8}),            xfAlign('center','center')));
      sc(cr,COL_DUE,   task.dueDate?fmtDate(task.dueDate):'—', cSt(bg,xfFont({sz:8,italic:true}),xfAlign('center','center')));

      mondays.forEach((mon,i)=>{
        const dc=COL_D0+i, me=new Date(mon.getTime()+7*864e5);
        const bs=task.dueDate, be=taskBarEndDate(task);
        const inB=bs&&be&&mon<new Date(be.getTime()+7*864e5)&&me>bs;
        const isF=inB&&bs>=mon&&bs<me;
        const isTod=mon<=today&&today<me;
        const isMk=tourn.tournDate&&mon<=tourn.tournDate&&tourn.tournDate<me;
        const isBudgetDL=deadline&&mon<=deadline&&deadline<me;

        let fill=bg, lbl='', fnt=xfFont({sz:7});
        if(inB){
          fill=barColor; fnt=xfFont({sz:7,bold:true,color:C.WHITE});
          if(isF) lbl=TLABELS[task.type]||task.task;
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
  [[C.BAR_CONFIRM,'Confirm Teams (1 month before)'],[C.BAR_BOOK,'Book Transport'],[C.BAR_BUDGET,'Submit Budget Request'],[C.GREEN,'Tournament Date']].forEach(([color,label],i)=>{
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
  const h2=['#','Tournament','Date','Location','Transport','Debaters','Book By','Budget Request By','Budget Deadline'];
  h2.forEach((h,ci)=>{ ws2[R(0,ci)]={v:h,t:'s',s:cSt(C.NAVY,xfFont({sz:9,bold:true,color:C.WHITE}),xfAlign('center','center'))}; });

  const seen2=new Set(); let r2=1,idx2=1;
  ganttTasks.forEach(t=>{
    if(!seen2.has(t.name)){
      seen2.add(t.name);
      const bookTask  =ganttTasks.find(x=>x.name===t.name&&x.type==='book');
      const budgTask  =ganttTasks.find(x=>x.name===t.name&&x.type==='budget');
      const alt=idx2%2===0;
      [idx2,t.name,t.tournDate?fmtDate(t.tournDate):'',t.loc||'',t.transport||'',
       t.debaters||'',
       bookTask?fmtDate(bookTask.dueDate):'N/A',
       budgTask?fmtDate(budgTask.dueDate):'No deadline set',
       deadline?fmtDate(deadline):'Not set',
      ].forEach((v,ci)=>{
        ws2[R(r2,ci)]={v:String(v),t:'s',s:cSt(alt?C.ROW_ALT:C.WHITE,xfFont({sz:9}),xfAlign(ci===0||ci>=5?'center':'left','center'))};
      });
      r2++;idx2++;
    }
  });
  ws2['!ref'] =`A1:${colLetter(8)}${r2}`;
  ws2['!cols']=[{wch:4},{wch:26},{wch:14},{wch:20},{wch:12},{wch:10},{wch:16},{wch:20},{wch:16}];
  XLSX.utils.book_append_sheet(wb,ws2,'Summary');

  XLSX.writeFile(wb,'tournament_gantt.xlsx');
  showStatus('Gantt chart exported!');
}

// ── Template download ─────────────────────

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Tournament Name','Date','Location','Transport','Debaters'],
    ['Spring Invitational',   '2026-04-15','Chicago, IL',      'fly',    8],
    ['Regional Championships','2026-05-20','Boston, MA',        'amtrak', 6],
    ['District Qualifier',    '2026-06-10','Silver Spring, MD', 'car',    4],
    ['State Finals',          '2026-07-08','Richmond, VA',      'metro',  10],
  ]);
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
