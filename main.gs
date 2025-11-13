var CONFIG = {
  SHEETS: ['Nov.'],
  HEADER_ROW: 2,
  SNAPSHOT_KEEP: 45,
  TITLE: 'Daily Changes',
  TIMEZONE: 'Asia/Bangkok',
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  DELIVERY_METHOD: 'OA_BROADCAST',     // use OA path (we'll push to group)
  OA_CHANNEL_ACCESS_TOKEN: '<your-new-token-here>',
  OA_GROUP_ID: '<OA_GROUP_ID>', 
  PREVIEW_MODE: 'BY_CHANGE_DATE',    
  HEADERS: {
    task:        'Task',
    requestBy:   'Request by',
    assignedTo:  'Assigned to',
    resources:   'Resources',
    startDate:   'Start Date',
    changeDate:  'Acted Date',       // ← renamed
    nextDueDate: 'Next Due Date',
    dueDate:     'Due Date',
    progress:    'Progress',
    note:        'Note',
    meetingTime: 'Meeting Time',
    by:          'By'                // ← NEW
  },
  WATCH_FIELDS: {
    task:        false,
    resources:   false,
    startDate:   false,
    changeDate:  false,              // stamped by script
    nextDueDate: false,         
    note:        false,      // ← stamp when edited
    dueDate:     false,               // ← stamp when edited
    meetingTime: false,
    progress:    false,
    by:          false
  },
  REQUIRE_BY_MATCH: true,
  SEND_IF_EMPTY: false,
  MAX_ROWS_PER_MESSAGE: 10,
  SEPARATOR_LINE: '------------------------------------',
  REALTIME_MIN_GAP_MS: 3000
};

// เปิด–ปิดงานประจำเช้า/เย็น (default = เปิด)
CONFIG.ENABLE_DAILY = true;

// อ่าน/ตั้งค่าสวิตช์จาก Properties (จะอยู่รอดแม้แก้สคริปต์)
function dailyEnabled_(){
  const p = PropertiesService.getDocumentProperties();
  const v = p.getProperty('DAILY_ENABLED');
  return v == null ? (CONFIG.ENABLE_DAILY !== false) : (v === '1');
}
function setDailyEnabled(on){
  PropertiesService.getDocumentProperties().setProperty('DAILY_ENABLED', on ? '1' : '0');
  ensureTriggersAlive_(); // สร้าง/ลบทริกเกอร์ตามสวิตช์ทันที
}

function getEditorEmail_() {
  try {
    var em = Session.getActiveUser().getEmail();
    return (em && /@/.test(em)) ? em : '';
  } catch (e) { return ''; }
}

/* ======================= DIFF FIELDS (used everywhere) ======================= */
var DIFF_FIELDS = [
  { key: 'task',        label: () => CONFIG.HEADERS.task,        type: 'text' },
  { key: 'requestBy',   label: () => CONFIG.HEADERS.requestBy,   type: 'text' },
  { key: 'assignedTo',  label: () => CONFIG.HEADERS.assignedTo,  type: 'text' },
  { key: 'resources',   label: () => CONFIG.HEADERS.resources,   type: 'text' },
  { key: 'startDate',   label: () => CONFIG.HEADERS.startDate,   type: 'date' },
  { key: 'nextDueDate', label: () => CONFIG.HEADERS.nextDueDate, type: 'date' },
  { key: 'dueDate',     label: () => CONFIG.HEADERS.dueDate,     type: 'date' },
  { key: 'meetingTime', label: () => CONFIG.HEADERS.meetingTime, type: 'time' },
  { key: 'note',        label: () => CONFIG.HEADERS.note,        type: 'text' }, // <-- add
  { key: 'progress',    label: () => CONFIG.HEADERS.progress,    type: 'text' },
  // { key: 'by',        ... }  // keep "By" OUT of diffs on purpose
];
function fieldTypeForKey_(key){ var f = DIFF_FIELDS.find(function(x){ return x.key === key; }); return f ? f.type : 'text'; }
function fieldLabelForKey_(key){ var f = DIFF_FIELDS.find(function(x){ return x.key === key; }); return f ? f.label() : key; }

function purgeLegacyTriggers_() {
  const legacy = new Set(['runMorningSend', 'initializeSnapshotNightly']);
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (legacy.has(fn)) ScriptApp.deleteTrigger(t);
  });
}

function ensureTriggersAlive_() {
  purgeDuplicateTimeTriggers_();
  purgeDuplicateRealtimeTriggers_();
  purgeLegacyTriggers_();

  let haveNight=false, haveMorning=false, haveRT=false;
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'runNightlyAlert') haveNight = true;
    if (f === 'runDueTodaySend') haveMorning = true;
    if (f === 'onEdit' || f === 'onMyEdit') haveRT = true;
  });

  // ⏸ เคารพสวิตช์
  if (!dailyEnabled_()) {
    removeNightAndMorningTriggers();   // ถ้าปิด ให้ลบทิ้งเผื่อยังเหลือ
  } else {
    if (!haveNight || !haveMorning) createNightAndMorningTriggersUsingConfig();
  }

  if (!haveRT) createRealtimeEditTrigger();
}
/* ======================= TIMEZONE ======================= */
function tz_() {
  if (CONFIG.TIMEZONE && typeof CONFIG.TIMEZONE === 'string') return CONFIG.TIMEZONE;
  try {
    var ss = SpreadsheetApp.getActive();
    var t1 = ss && ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : null;
    if (t1 && typeof t1 === 'string') return t1;
  } catch (e) {}
  try {
    var t2 = Session.getScriptTimeZone();
    if (t2 && typeof t2 === 'string') return t2;
  } catch (e) {}
  return 'Etc/GMT';
}

/** === DEBUG: same idea as menu "Initialize / Repair (once)" === */
function debugInitializeRepairOnce() {
  initializeProject_();              // stamps UIDs, seeds baseline, ensures triggers
  Logger.log('Initialize/Repair done.');
}

/** === DEBUG: hard reset just the daily time triggers (keep one of each) === */
function debugHardResetDailyTriggers() {
  hardResetDailyTriggers_();         // deletes any dupes and recreates clean pair
  Logger.log('Daily time triggers reset. Morning=%s Night=%s',
             getMorningTime_(), getNightTime_());
}

/** === DEBUG: full repair (daily pair + realtime onEdit) === */
function debugRepairAllTriggers() {
  hardResetDailyTriggers_();         // clean morning/night
  removeRealtimeTrigger();           // drop any existing onEdit
  createRealtimeEditTrigger();       // add a fresh onEdit
  Logger.log('Repaired: morning/night + realtime onEdit.');
  debugListTriggers();               // print what’s active
}

/** === DEBUG: print current trigger list & next run times to Logs === */
function debugListTriggers() {
  const tz = tz_();
  const trig = ScriptApp.getProjectTriggers();
  Logger.log('Timezone: %s  Now: %s', tz,
             Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'));
  trig.forEach((t,i) => Logger.log('%d) %s', i+1, t.getHandlerFunction()));
  Logger.log('Next morning run: %s', nextRunInTz_(getMorningTime_()));
  Logger.log('Next night   run: %s', nextRunInTz_(getNightTime_()));
}

/** === DEBUG: open the status dialog (same as menu "Developer Setting") === */
function debugShowTriggerStatus() {
  showTriggerStatus();
}

/** === DEBUG: clear the 8-minute cooldown keys (if you want to re-run immediately) === */
function debugClearCooldownGuards() {
  const p = PropertiesService.getDocumentProperties();
  ['RUN_GUARD_runNightlyAlert','RUN_GUARD_runDueTodaySend'].forEach(k => p.deleteProperty(k));
  Logger.log('Cleared cooldown guards.');
}

function purgeDuplicateTimeTriggers_(){
  const wanted = new Set(['runDueTodaySend','runNightlyAlert']);
  const seen = { runDueTodaySend: false, runNightlyAlert: false };

  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (!wanted.has(fn)) return;
    if (!seen[fn]) { seen[fn] = true; return; }  // keep the first one
    // delete extras
    ScriptApp.deleteTrigger(t);
  });
}

/* ======================= MENU ======================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Kanban Alert')
    .addSeparator()
    .addItem("▶ Manually Notification. 🔔", 'openReviewSidebar')
    .addItem('▶ Preview Change Summary Notification. ☀️', 'previewNightlyAlertMessage')
    .addItem('▶ Preview Tomorrow Task Notification. 🌙', 'previewDueListMessage')
    // ไม่มี Pause / Resume / Initialize อีกแล้ว
    .addToUi();

  try { seedCellCacheOnOpen_(); } catch(_){}
  try { openStickyUiIfAny_ && openStickyUiIfAny_(); } catch(_){}
}


function pauseDailySends(){
  setDailyEnabled(false);
  try { SpreadsheetApp.getActive().toast('Paused: ปิดงานประจำเช้า/เย็นแล้ว'); } catch(_){}
}
function resumeDailySends(){
  setDailyEnabled(true);
  try { SpreadsheetApp.getActive().toast('Resumed: เปิดงานประจำเช้า/เย็นแล้ว'); } catch(_){}
}

function initializeProject_() {
  ensureTriggersAlive_();                 
  watchedSheets_().forEach(sh => {
    const lastCol = sh.getLastColumn();
    if (lastCol <= 0) return;
    const header = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
    const idx = indexMap_(header);
    ensureRowUidsForSheet_(sh, idx);
  });
  seedBaselineSnapshotIfMissing_();       // safe now (installable context)
  seedCellCacheOnOpen_();
}

function ensureRowUidsForSheet_(sh, idx) {
  const lastRow = sh.getLastRow();
  const startRow = CONFIG.HEADER_ROW + 1;
  const rows = Math.max(0, lastRow - startRow + 1);
  if (rows <= 0) return [];
  const taskCol = idx.task;
  const rng = sh.getRange(startRow, taskCol, rows, 1);
  const notes = rng.getNotes(); // 2D array
  let changed = false;
  for (let r = 0; r < rows; r++) {
    let uid = (notes[r][0] || '').trim();
    if (!uid) { uid = Utilities.getUuid(); notes[r][0] = uid; changed = true; }
  }
  if (changed) rng.setNotes(notes);
  return notes.map(row => row[0]);
}

function reindexSnapshotToUid_(prevSnap, todayRows) {
  prevSnap = prevSnap || {};
  // If keys already look like UUIDs, keep as-is
  const looksUid = Object.keys(prevSnap).some(k => String(k).indexOf('-') !== -1);
  if (looksUid) return prevSnap;
  const out = {};
  Object.keys(prevSnap).forEach(k => {
    const s = prevSnap[k] || {};
    if (s.uid) out[String(s.uid)] = s;
  });
  if (Object.keys(out).length) return out;
  const prevByRow = prevSnap;
  (todayRows || []).forEach(r => {
    const byRow = prevByRow[String(r.row)];
    if (byRow) out[String(r.uid)] = byRow;
  });
  return out;
}

function keyForRow_(rowObj){
  return String(rowObj.uid || rowObj.row);
}
// ------------------------------------------------------------

function showLockedEditModal_() {
  try {
    const info = getEditLockInfo();
    const html = HtmlService.createHtmlOutput(buildLockedEditHtml_(info))
      .setWidth(420).setHeight(260);
    SpreadsheetApp.getUi().showModalDialog(html, 'Editing Locked');
  } catch (_) { /* ignore */ }
}

function revertFromCacheOrOld_(rng, e) {
  const sh = rng.getSheet();
  const rows = rng.getNumRows(), cols = rng.getNumColumns();
  const top = rng.getRow(), left = rng.getColumn();
  const cache = loadCellCache_(sh);
  if (rows === 1 && cols === 1) {
    const key = top + ':' + left;
    const prev = (typeof e.oldValue !== 'undefined') ? e.oldValue : cache[key];
    if (typeof prev !== 'undefined') rng.setValue(prev); // else no-op
    return;
  }
  const after = rng.getValues();
  const out = [];
  for (let r = 0; r < rows; r++) {
    const rowArr = [];
    for (let c = 0; c < cols; c++) {
      const k = (top + r) + ':' + (left + c);
      rowArr.push((k in cache) ? cache[k] : after[r][c]); // ← keep as-is if unknown
    }
    out.push(rowArr);
  }
  rng.setValues(out);
}

function isReverting_() {
  const p = PropertiesService.getDocumentProperties();
  return p.getProperty('REVERTING') === '1';
}
function setRevertingFlag_(on) {
  const p = PropertiesService.getDocumentProperties();
  if (on) p.setProperty('REVERTING', '1');
  else p.deleteProperty('REVERTING');
}

function openChangeLogDialog() {
  var html = HtmlService.createHtmlOutput(buildChangeLogDialogHtml_())
    .setWidth(720).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Change Log');
}

function getChangeLogMeta() {
  var tz = tz_();
  return {
    tz: tz,
    now: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss')
  };
}

function getChangeLogData(opts) {
  const sh = getTargetSheet_();
  const suf = sheetKeySuffix_(sh);

  const from = (opts && opts.from) ? opts.from : '';
  const to   = (opts && opts.to)   ? opts.to   : '';
  const tz = tz_();

  const p = PropertiesService.getDocumentProperties();
  const raw = p.getProperty('KANBAN_CHANGELOG' + suf) || '[]';  // ← per-sheet
  let arr = []; try { arr = JSON.parse(raw); } catch(_) {}

  function inRange(ts) {
    const d = new Date(ts);
    const yyyyMMdd = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (from && yyyyMMdd < from) return false;
    if (to   && yyyyMMdd > to)   return false;
    return true;
  }

  return arr
  .filter(e => inRange(e.ts))
  .sort((a,b) => a.ts - b.ts)
  .map(e => {
    const type = fieldTypeForKey_(e.field);
    const before = normForType_(type, e.from);
    const after  = normForType_(type, e.to);
    return {
      when: Utilities.formatDate(new Date(e.ts), tz, 'dd/MM/yyyy HH:mm:ss'),
      whenTs: e.ts,
      row: e.row,
      uid: e.uid || null,                 // ← NEW
      fieldKey: e.field,
      fieldLabel: fieldLabelForKey_(e.field),
      beforeDisp: dispForType_(type, before),
      afterDisp:  dispForType_(type, after),
      task: e.task || '',
      by:   e.by   || '',
      who:  e.who  || ''
    };
  });
}

function clearChangelogToday_() {
  const sh = SpreadsheetApp.getActiveSheet();
  const suf = sheetKeySuffix_(sh);
  const tz = tz_();
  const p = PropertiesService.getDocumentProperties();
  let arr = []; try { arr = JSON.parse(p.getProperty('KANBAN_CHANGELOG' + suf) || '[]'); } catch(_) {}
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  arr = arr.filter(e => Utilities.formatDate(new Date(e.ts), tz, 'yyyy-MM-dd') !== today);
  p.setProperty('KANBAN_CHANGELOG' + suf, JSON.stringify(arr));
}

function getSheetMeta() {
  var sh  = getTargetSheet_();
  var url = SpreadsheetApp.getActive().getUrl();
  return { url: url, gid: sh.getSheetId(), name: sh.getName() };
}

function getStartOfDayBaseline_(sh, actedDay){
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const latestDay = p.getProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf);
  const prevDay   = p.getProperty('KANBAN_SNAPSHOT_PREV_DATE'   + suf);
  // If a snapshot exists for "today", use yesterday’s instead
  if (latestDay === actedDay && prevDay) {
    return loadSnapshotByDay_(prevDay, sh);
  }
  // Otherwise the latest snapshot is already yesterday (or earlier)
  if (latestDay) return loadSnapshotByDay_(latestDay, sh);
  return {}; // no baseline yet
}

// Build "yesterday vs realtime" diff rows
function getRealtimeCompareData(opts) {
  opts = opts || {};
  const tz = tz_();
  const sh = getTargetSheet_();
  const latest   = loadLatestSnapshot_(sh);        // {day, snap}
  const tableNow = readTableForSheet_(sh);
  const lastSnap = reindexSnapshotToUid_(latest.snap || {}, tableNow); // ← normalize keys
  const actedDay = latest.day;
  const nowSnap  = buildSnapshot_(tableNow);       // UID-keyed
  // live maps keyed by UID (fallback row for legacy)
  const liveMap = Object.create(null);
  tableNow.forEach(r => { liveMap[keyForRow_(r)] = r; });

  // latest edit today per UID
  const todayLog = loadChangelogForToday_(sh);
  const latestByKey = Object.create(null);
  todayLog.forEach(e => {
    const k = String(e.uid || e.row);
    if (!latestByKey[k] || e.ts > latestByKey[k].ts) latestByKey[k] = { ts: e.ts, who: e.who || '' };
  });

  // union of keys from baseline & now
  const keys = Object.keys(nowSnap);
  Object.keys(lastSnap).forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); });

  const out = [];
  keys.forEach(k => {
    const before = lastSnap[k] || {};
    const after  = nowSnap[k]  || {};
    const diffs = [];

    DIFF_FIELDS.forEach(f => {
      const b = normForType_(f.type, before[f.key]);
      const a = normForType_(f.type, after[f.key]);
      if ((b || '') !== (a || '')) {
        diffs.push({
          fieldKey:   f.key,
          fieldLabel: fieldLabelForKey_(f.key),
          beforeDisp: dispForType_(f.type, b),
          afterDisp:  dispForType_(f.type, a)
        });
      }
    });

    if (!diffs.length) return;

    const live = liveMap[k] || {};
    const last = latestByKey[k] || null;

    out.push({
      row: +live.row || 0, // keep showing current row number
      task: normalizeText_(live.task),
      by: normalizeText_(live.by),
      progress: normalizeText_(live.progress),
      dueDisp: displayDate_(normalizeDate_(live.dueDate)),
      badge:   deadlineBadge_(live.dueDate) || '',
      actedToday: isActedOnCustomDay_(live.changeDate, todayCustomDate_()),
      lastTs: last ? last.ts : 0,
      lastTsDisp: last ? Utilities.formatDate(new Date(last.ts), tz, 'dd/MM/yyyy HH:mm:ss') : '',
      who: last ? (last.who || '') : '',
      snapshotDayDisp: actedDay ? Utilities.formatDate(toDate_(actedDay), tz, 'dd/MM/yyyy') : '-',
      diffs: diffs
    });
  });

  return out;
}

function buildChangeLogDialogHtml_() {
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Change Log</title>
<style>
  :root { --ok:#137333; --okbg:#e8f7ed; --bad:#a50e0e; --badbg:#fdecea; --muted:#666; }
  body { font:13px Arial, sans-serif; margin:0; padding:14px; color:#222; background:#fff; }
  h2 { margin:0 0 10px; font-size:18px; }
  .meta { color: var(--muted); margin-bottom: 12px; }
  .grid { display:grid; grid-template-columns: 1fr; gap:10px; }
  .card { border:1px solid #e6e6e6; border-radius:10px; padding:12px; box-shadow:0 1px 0 rgba(0,0,0,.03); background:#fff; }
  .row { display:flex; flex-wrap:wrap; gap:8px 12px; align-items:center; }
  .label { color: var(--muted); }
  input, select, button { font:inherit; padding:6px 8px; border:1px solid #ddd; border-radius:8px; background:#fff; }
  button { cursor:pointer; }
  button:hover { background:#f7f7f7; }
  table { border-collapse:collapse; width:100%; font-size:12px; }
  th, td { border:1px solid #eee; padding:6px 8px; vertical-align:top; }
  th { background:#fafafa; position:sticky; top:0; z-index:1; user-select:none; }
  th.sortable { cursor:pointer; }
  th.sortable .arrow { font-size:11px; opacity:.6; margin-left:4px; }
  .scroll { max-height: 340px; overflow:auto; border:1px solid #eee; border-radius:8px; }
  .mono { font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
  mark { background: #fff2ac; padding:0 2px; border-radius:3px; }
  .footer { display:flex; justify-content:flex-end; gap:8px; margin-top:10px; }
  .chip { display:inline-flex; gap:6px; align-items:center; padding:2px 8px; border:1px solid #eee; border-radius:999px; background:#fafafa; }
  .chips { display:flex; flex-wrap:wrap; gap:6px; }
  .muted { color:#666; }
</style>
</head>
<body>
  <h2>Change Log</h2>
  <div id="meta" class="meta">Loading…</div>

  <div class="grid">
    <!-- Filters / Controls -->
    <div class="card">
      <div class="row">
        <span class="label">Quick range</span>
        <select id="range">
          <option value="today">Today</option>
          <option value="7d">Last 7 days</option>
          <option value="30d" selected>Last 30 days</option>
          <option value="all">All</option>
        </select>

        <span class="label">From</span><input type="date" id="from">
        <span class="label">To</span><input type="date" id="to">
        <span class="label">Field</span><select id="field"></select>
        <span class="label">Row</span><input type="number" id="qRow" style="width:90px;" min="0" placeholder="#">
        <span class="label">Task</span><input type="text" id="qTask" placeholder="contains…">
        <span class="label">By</span><input type="text" id="qBy" placeholder="contains…">
        <span class="label">Editor</span><input type="text" id="qWho" placeholder="email contains…">

        <span class="label">Rows/page</span>
        <select id="pageSize">
          <option>50</option><option selected>200</option><option>1000</option><option>All</option>
        </select>

        <span class="label">Auto-refresh</span>
        <select id="auto">
          <option value="0" selected>Off</option>
          <option value="30">30s</option>
          <option value="60">1m</option>
          <option value="300">5m</option>
        </select>

        <button id="btnReload">Reload</button>
        <button id="btnReset">Reset</button>
        <button id="btnClearToday">Clear today</button>
        <button id="btnExportVisible">Export visible CSV</button>
        <button id="btnExportFull">Export full-range CSV</button>
      </div>
      <div class="row small" id="summary"></div>
    </div>

    <!-- Quick stats -->
    <div class="card" id="statsCard" style="display:none;">
      <div class="row">
        <div style="flex:1;">
          <div class="label">Top Fields (shown)</div>
          <div class="chips" id="statsFields"></div>
        </div>
        <div style="flex:1;">
          <div class="label">Top Editors (shown)</div>
          <div class="chips" id="statsEditors"></div>
        </div>
      </div>
    </div>

    <!-- Table -->
    <div class="card">
      <div class="scroll">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-k="when">When<span class="arrow"></span></th>
              <th class="sortable" data-k="row">#Row<span class="arrow"></span></th>
              <th class="sortable" data-k="fieldLabel">Field<span class="arrow"></span></th>
              <th>Before</th>
              <th>After</th>
              <th class="sortable" data-k="task">Task<span class="arrow"></span></th>
              <th class="sortable" data-k="by">By<span class="arrow"></span></th>
              <th class="sortable" data-k="who">Editor<span class="arrow"></span></th>
              <th>Open</th>
            </tr>
          </thead>
          <tbody id="tbody"><tr><td colspan="9" class="small">Loading…</td></tr></tbody>
        </table>
      </div>

      <!-- Pagination -->
      <div class="row" id="pager" style="justify-content:flex-end; margin-top:8px;">
        <button id="prev">Prev</button>
        <span class="label" id="pageInfo">Page 1/1</span>
        <button id="next">Next</button>
      </div>
    </div>
  </div>

  <div class="footer">
    <button onclick="google.script.host.close()">Close</button>
  </div>

<script>
let RAW = [];               // all rows for selected date range
let SORT = { key: 'when', dir: 'desc' };
let PAGE = 1;
let PAGE_SIZE = 200;
let SHEET_META = null;
let autoTimer = null;

function setMeta(d){ document.getElementById('meta').textContent = 'Timezone: ' + d.tz + ' • Now: ' + d.now; }
function setSheetMeta(m){ SHEET_META = m; }

function dstrLocal(d){
  const z = new Date(d); z.setMinutes(z.getMinutes()-z.getTimezoneOffset());
  return z.toISOString().slice(0,10);
}
function rangeToDates() {
  const sel = document.getElementById('range').value;
  if (sel === 'today') return {from: dstrLocal(new Date()), to: dstrLocal(new Date())};
  if (sel === '7d' || sel === '30d') {
    const days = sel === '7d' ? 7 : 30;
    const f = new Date(Date.now() - (days-1)*24*3600*1000);
    return {from: dstrLocal(f), to: dstrLocal(new Date())};
  }
  return {from:'', to:''}; // all
}

function loadFieldsDropdown() {
  const sel = document.getElementById('field');
  sel.innerHTML = '<option value="">Any</option>';
}

function reload(){
  let {from,to} = rangeToDates();
  const fInp = document.getElementById('from').value;
  const tInp = document.getElementById('to').value;
  if (fInp) from = fInp; if (tInp) to = tInp;

  google.script.run.withSuccessHandler(function(rows){
    RAW = rows || [];
    // populate field dropdown with distinct labels
    const sel = document.getElementById('field');
    const keep = sel.value || '';
    const labels = Array.from(new Set(RAW.map(r=>r.fieldLabel))).sort();
    sel.innerHTML = '<option value="">Any</option>' + labels.map(l=>'<option>'+escapeHtml(l)+'</option>').join('');
    if (labels.includes(keep)) sel.value = keep;

    PAGE = 1;
    render();
  }).getChangeLogData({from,to});
}

function getFilters(){
  return {
    field: document.getElementById('field').value || '',
    row: (document.getElementById('qRow').value||'').trim(),
    task: (document.getElementById('qTask').value||'').toLowerCase().trim(),
    by:   (document.getElementById('qBy').value||'').toLowerCase().trim(),
    who:  (document.getElementById('qWho').value||'').toLowerCase().trim()
  };
}

function applyFilters(rows){
  const f = getFilters();
  if (f.field) rows = rows.filter(r => r.fieldLabel === f.field);
  if (f.row)   rows = rows.filter(r => String(r.row) === String(f.row));
  if (f.task)  rows = rows.filter(r => String(r.task||'').toLowerCase().includes(f.task));
  if (f.by)    rows = rows.filter(r => String(r.by||'').toLowerCase().includes(f.by));
  if (f.who)   rows = rows.filter(r => String(r.who||'').toLowerCase().includes(f.who));
  return rows;
}

function sortRows(rows){
  const k = SORT.key, d = SORT.dir === 'asc' ? 1 : -1;
  return rows.sort((a,b)=>{
    if (k === 'row')  return (+a.row - +b.row) * d;
    if (k === 'when') return ((a.whenTs || 0) - (b.whenTs || 0)) * d;
    const sx = String(a[k]||'').toLowerCase(), sy = String(b[k]||'').toLowerCase();
    if (sx < sy) return -1*d; if (sx > sy) return 1*d; return 0;
  });
}


function paginate(rows){
  const sizeSel = document.getElementById('pageSize').value;
  PAGE_SIZE = sizeSel === 'All' ? Infinity : parseInt(sizeSel,10);
  if (!isFinite(PAGE_SIZE) || PAGE_SIZE <= 0) PAGE_SIZE = Infinity;

  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / (isFinite(PAGE_SIZE)? PAGE_SIZE : total)));
  if (PAGE > pages) PAGE = pages;

  const start = isFinite(PAGE_SIZE) ? (PAGE-1)*PAGE_SIZE : 0;
  const slice = isFinite(PAGE_SIZE) ? rows.slice(start, start+PAGE_SIZE) : rows.slice();

  document.getElementById('pageInfo').textContent = 'Page ' + PAGE + '/' + pages;
  document.getElementById('prev').disabled = (PAGE<=1);
  document.getElementById('next').disabled = (PAGE>=pages);

  return slice;
}

function rowLink(row){
  if (!SHEET_META) return '#';
  const base = SHEET_META.url.split('#')[0];
  return base + '#gid=' + SHEET_META.gid + '&range=A' + row + ':A' + row;
}

function computeStats(rows){
  const f = {}; const w = {};
  rows.forEach(r => {
    f[r.fieldLabel] = (f[r.fieldLabel]||0)+1;
    if (r.who) w[r.who] = (w[r.who]||0)+1;
  });
  function top(obj, n){
    return Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,8)
      .map(([k,v])=>({k,v}));
  }
  return { fields: top(f,8), editors: top(w,8) };
}

function render(){
  // pipeline
  let rows = applyFilters(RAW.slice());
  rows = sortRows(rows);
  const visible = paginate(rows);

  // stats
  const stats = computeStats(rows);
  const statsCard = document.getElementById('statsCard');
  statsCard.style.display = rows.length ? '' : 'none';
  document.getElementById('statsFields').innerHTML =
    stats.fields.map(x=>'<span class="chip">'+escapeHtml(x.k)+' · '+x.v+'</span>').join('');
  document.getElementById('statsEditors').innerHTML =
    stats.editors.map(x=>'<span class="chip">'+escapeHtml(x.k)+' · '+x.v+'</span>').join('');

  // table
  const tb = document.getElementById('tbody');
  if (!visible.length) { tb.innerHTML = '<tr><td colspan="9" class="small">No entries.</td></tr>'; }
  else {
    tb.innerHTML = visible.map(r => {
      const marked = markDiff(r.beforeDisp, r.afterDisp);
      return '<tr>'
        + td(r.when) + td(r.row) + td(r.fieldLabel)
        + tdMonoHTML(marked.a) + tdMonoHTML(marked.b)
        + td(r.task||'') + td(r.by||'') + td(r.who||'')
        + '<td><a href="'+rowLink(r.row)+'" target="_blank" title="Open this row">Open</a></td>'
        + '</tr>';
    }).join('');
  }

  document.getElementById('summary').textContent =
    'Total (range): ' + RAW.length + ' • Shown (after filters): ' + rows.length;
  paintSortArrows();
}

function paintSortArrows(){
  document.querySelectorAll('th.sortable .arrow').forEach(el=>el.textContent='');
  const th = document.querySelector('th.sortable[data-k="'+SORT.key+'"] .arrow');
  if (th) th.textContent = SORT.dir === 'asc' ? '▲' : '▼';
}

// ===== table cell helpers =====
function td(s){ return '<td>'+escapeHtml(s)+'</td>'; }
function tdMonoHTML(html){ return '<td class="mono">'+html+'</td>'; }
function escapeHtml(s){
  return String(s==null?'':s).replace(/[&<>\"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// Highlight middle diff (simple LCP/LCS)
function markDiff(a, b){
  a = String(a ?? ''); b = String(b ?? '');
  if (a === b) return { a: escapeHtml(a), b: escapeHtml(b) };
  let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++;
  let j = 0; while (j < a.length - i && j < b.length - i && a[a.length-1-j] === b[b.length-1-j]) j++;
  function wrap(s){
    const pre = escapeHtml(s.slice(0,i));
    const mid = escapeHtml(s.slice(i, s.length - j));
    const suf = escapeHtml(s.slice(s.length - j));
    return pre + '<mark>' + mid + '</mark>' + suf;
  }
  return { a: wrap(a), b: wrap(b) };
}

// ===== wiring / events =====
function exportVisibleCSV(){
  // Export currently filtered + sorted rows (not just the page)
  let rows = sortRows(applyFilters(RAW.slice()));
  const headers = ['When','Row','Field','Before','After','Task','By','Editor'];
  function esc(v){ return '"' + String(v==null?'':v).replace(/"/g,'""') + '"'; }
  const lines = [headers.map(esc).join(',')];
  rows.forEach(r=>{
    lines.push([r.when, r.row, r.fieldLabel, r.beforeDisp, r.afterDisp, r.task||'', r.by||'', r.who||''].map(esc).join(','));
  });
  const blob = new Blob([lines.join('\\n')], {type:'text/csv'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url;
  a.download = 'change-log-visible-' + Date.now() + '.csv';
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 800);
}

function exportFullCSV(){
  let {from,to} = rangeToDates();
  const fInp = document.getElementById('from').value;
  const tInp = document.getElementById('to').value;
  if (fInp) from = fInp; if (tInp) to = tInp;
  google.script.run.withSuccessHandler(function(res){
    const blob = new Blob([res.csv], {type:'text/csv'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = res.filename;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 800);
  }).getChangeLogCsv({from,to});
}

function clearToday(){
  if (!confirm("Clear today's changelog entries?")) return;
  google.script.run.withSuccessHandler(reload).clearChangelogToday_();
}

function resetFilters(){
  document.getElementById('range').value = '30d';
  ['from','to','field','qRow','qTask','qBy','qWho'].forEach(id=>document.getElementById(id).value='');
  PAGE = 1; SORT = {key:'when', dir:'desc'}; render();
  reload();
}

function handleSort(e){
  const th = e.target.closest('th.sortable'); if (!th) return;
  const k = th.getAttribute('data-k');
  if (SORT.key === k) SORT.dir = (SORT.dir === 'asc' ? 'desc' : 'asc');
  else { SORT.key = k; SORT.dir = 'asc'; }
  render();
}

function wire(){
  // Buttons
  document.getElementById('btnReload').onclick = reload;
  document.getElementById('btnReset').onclick = resetFilters;
  document.getElementById('btnClearToday').onclick = clearToday;
  document.getElementById('btnExportVisible').onclick = exportVisibleCSV;
  document.getElementById('btnExportFull').onclick = exportFullCSV;

  // Inputs
  document.getElementById('range').onchange = reload;
  ['from','to','field','qRow','qTask','qBy','qWho'].forEach(id=>{
    const el = document.getElementById(id);
    el.onchange = render; el.oninput = render;
  });
  document.getElementById('pageSize').onchange = ()=>{ PAGE = 1; render(); };

  // Pager
  document.getElementById('prev').onclick = ()=>{ PAGE = Math.max(1, PAGE-1); render(); };
  document.getElementById('next').onclick = ()=>{ PAGE = PAGE+1; render(); };

  // Sortable headers
  document.querySelectorAll('th.sortable').forEach(th => th.addEventListener('click', handleSort));

  // Auto-refresh
  document.getElementById('auto').onchange = function(){
    if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
    const secs = parseInt(this.value,10)||0;
    if (secs > 0) autoTimer = setInterval(reload, secs*1000);
  };
}

function boot(){
  wire();
  loadFieldsDropdown();
  google.script.run.withSuccessHandler(setMeta).getChangeLogMeta();
  google.script.run.withSuccessHandler(setSheetMeta).getSheetMeta();
  reload();
}
boot();
</script>
</body>
</html>`;
}


// Seed a baseline snapshot the very first time (or after clearing props) — for ALL watched tabs
function seedBaselineSnapshotIfMissing_() {
  const p = PropertiesService.getDocumentProperties();
  watchedSheets_().forEach(sh => {
    const suf = sheetKeySuffix_(sh);
    if (!p.getProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf)) {
      const nowData = readTableForSheet_(sh);
      saveSnapshot_(buildSnapshot_(nowData), sh);
    }
  });
}

// 7) seed cache for all tabs เมื่อเปิดไฟล์
function seedCellCacheOnOpen_() {
  try {
    watchedSheets_().forEach(sh => {
      const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
      const startRow = CONFIG.HEADER_ROW + 1;
      const rows = Math.max(0, lastRow - startRow + 1);
      if (rows <= 0) { saveCellCache_({}, sh); return; }
      const MAX_CELLS = 15000;
      const takeRows = Math.max(0, Math.min(rows, Math.floor(MAX_CELLS / Math.max(1, lastCol))));
      if (takeRows <= 0) { saveCellCache_({}, sh); return; }
      const vals = sh.getRange(startRow, 1, takeRows, lastCol).getValues();
      const cache = {};
      for (let r = 0; r < vals.length; r++) for (let c = 0; c < lastCol; c++) {
        cache[(startRow + r) + ':' + (1 + c)] = vals[r][c];
      }
      saveCellCache_(cache, sh);
    });
  } catch (_) {}
}
/* ======================= EDIT-WINDOW INFO + COUNTDOWN MODAL ======================= */
// When will editing unlock next?
function computeNextUnlockDate_() {
  const tz = tz_();
  const now = new Date();
  const mt = parseHHmm_(getMorningTime_(), '06:00');
  const nt = parseHHmm_(getNightTime_(),   '21:00');
  function atToday(h, m) {
    const base = Utilities.formatDate(now, tz, 'yyyy/MM/dd') + ' 00:00:00';
    const d = new Date(base);
    d.setHours(h); d.setMinutes(m); d.setSeconds(0); d.setMilliseconds(0);
    return d;
  }
  const morningToday = atToday(mt.h, mt.m);
  const nightToday   = atToday(nt.h, nt.m);
  const within = isWithinEditWindow_();
  if (within) return null;
  if (now.getTime() <= morningToday.getTime()) {
    return morningToday; // unlock later today
  }
  const tomorrowMorning = new Date(morningToday.getTime() + 24*3600*1000);
  return tomorrowMorning;
}
function getEditLockInfo() {
  const tz = tz_();
  const p = PropertiesService.getDocumentProperties();
  const lockOn = p.getProperty('EDIT_LOCK_ON') === '1';
  const allowedNow = isWithinEditWindow_();

  let unlockAt = null, unlockAtMs = null, unlockAtDisp = '-';
  if (lockOn && !allowedNow) {
    unlockAt = computeNextUnlockDate_();
    if (unlockAt) {
      unlockAtMs = unlockAt.getTime();
      unlockAtDisp = Utilities.formatDate(unlockAt, tz, 'dd/MM/yyyy HH:mm');
    }
  }
  return {
    tz, lockOn, allowedNow,
    unlockAtMs, unlockAtDisp,
    nowDisp: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'),
    morning: getMorningTime_(),
    night: getNightTime_()
  };
}
function buildLockedEditHtml_(info) {
  // we’ll render a live countdown until this timestamp (ms since epoch)
  var targetMs = Number(info.unlockAtMs || 0);
  var msgTop   = 'Editing is locked. Window: ' + info.morning + ' → ' + info.night;
  var tz       = info.tz || tz_();
  var unlockAt = info.unlockAtDisp || '-';
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Editing Locked</title>
<style>
  :root{
    --bg:#fff; --panel:#ffffff; --ink:#1f2328; --muted:#6b7280; --accent:#137333;
    --border:#e5e7eb; --shadow:0 10px 30px rgba(0,0,0,.12);
    --okbg:#e8f7ed; --ok:#137333; --pill:#eef2ff;
  }
  *{box-sizing:border-box}
  body{margin:0;background:#f6f7f9;font:14px/1.4 system-ui, -apple-system, Segoe UI, Roboto, Arial, "Noto Sans", sans-serif;color:var(--ink)}
  .wrap{display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}
  .card{
    width:420px; background:var(--panel); border:1px solid var(--border);
    border-radius:16px; box-shadow:var(--shadow); padding:18px 18px 16px; text-align:center;
  }
  h2{margin:0 0 8px; font-size:18px}
  .muted{color:var(--muted)}
  .pill{display:inline-flex;align-items:center;gap:8px;margin:8px 0 12px;padding:6px 10px;border-radius:999px;background:var(--okbg);color:var(--ok);font-weight:600;font-size:12px}
  /* Big digital countdown */
  .timer{
    font:700 36px/1.05 ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace;
    letter-spacing:1px; margin:10px 0 8px;
  }
  .colon{animation:blink 1s steps(1,end) infinite}
  @keyframes blink{50%{opacity:.2}}
  .under{font-size:12px;color:var(--muted)}
  /* progress bar that drains to 0% as we approach unlock */
  .bar{height:8px; background:#eef2f7; border-radius:999px; overflow:hidden; margin:10px 0 4px}
  .fill{height:100%; width:100%;
        background:linear-gradient(90deg,#22c55e,#16a34a);
        transition:width .2s linear}
  /* tiny chips for hm/s readout */
  .chips{display:flex;justify-content:center;gap:8px;margin-top:8px}
  .chip{padding:2px 8px;border-radius:999px;background:#f3f4f6;color:#374151;font-size:12px}
  .footer{display:flex;justify-content:center;margin-top:12px}
  .btn{border:1px solid var(--border);background:#fff;padding:8px 12px;border-radius:10px;cursor:pointer}
  .btn:hover{background:#f8fafc}
</style>
</head>
<body>
  <div class="wrap">
    <div class="card" role="dialog" aria-live="polite">
      <h2>Editing Locked</h2>
      <div class="muted">${msgTop}</div>
      <div class="pill">
        <span>🔒 Edits are blocked right now</span>
      </div>
      <div id="timer" class="timer">--<span class="colon">:</span>--<span class="colon">:</span>--</div>
      <div class="bar" aria-hidden="true"><div id="fill" class="fill" style="width:100%"></div></div>
      <div class="under">Unlocks at <b>${unlockAt}</b> (${tz})</div>
      <div class="chips">
        <div id="dChip" class="chip" style="display:none;"></div>
        <div id="hChip" class="chip">--h</div>
        <div id="mChip" class="chip">--m</div>
        <div id="sChip" class="chip">--s</div>
      </div>
      <div class="footer"><button class="btn" onclick="google.script.host.close()">OK</button></div>
    </div>
  </div>
<script>
  const target = ${targetMs};
  const startedAt = Date.now();
  const totalMs   = Math.max(1, target - startedAt);
  function z(n){ return String(n).padStart(2,'0'); }
  function render(ms){
    ms = Math.max(0, ms);
    const s  = Math.floor(ms/1000);
    const d  = Math.floor(s/86400);
    const hh = Math.floor((s % 86400) / 3600);
    const mm = Math.floor((s % 3600) / 60);
    const ss = s % 60;
    const t = (d>0 ? (d + 'd ' + z(hh) + '<span class="colon">:</span>' + z(mm) + '<span class="colon">:</span>' + z(ss))
                   : (z(hh) + '<span class="colon">:</span>' + z(mm) + '<span class="colon">:</span>' + z(ss)));
    document.getElementById('timer').innerHTML = t;
    const dChip = document.getElementById('dChip');
    dChip.style.display = d>0 ? '' : 'none';
    dChip.textContent   = d + 'd';
    document.getElementById('hChip').textContent = (d>0 ? hh : Math.floor(s/3600)) + 'h';
    document.getElementById('mChip').textContent = mm + 'm';
    document.getElementById('sChip').textContent = ss + 's';
    const pct = Math.max(0, Math.min(100, Math.round(ms * 100 / totalMs)));
    document.getElementById('fill').style.width = pct + '%';
  }
  function tick(){
    const left = target - Date.now();
    render(left);
    if (left <= 0) setTimeout(()=>google.script.host.close(), 900);
  }
  render(target - Date.now());
  setInterval(tick, 250);
</script>
</body>
</html>`;
}
/* ======================= DATE/TIME PARSERS ======================= */
function excelSerialToDate_(n) {
  if (typeof n !== 'number') n = Number(n);
  if (!isFinite(n)) return null;
  var ms = Math.round((n - 25569) * 86400000);
  return new Date(ms);
}
function toDate_(v) {
  if (v instanceof Date) return v;

  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) return excelSerialToDate_(v);
    if (v >= 0 && v < 2)        return excelSerialToDate_(v);
    return new Date(Math.round(v));
  }
  if (typeof v === 'string') {
    var s = v.trim();
    if (!s) return null;
    if (/^\d+(\.\d+)?$/.test(s)) {
      var num = +s;
      if (num > 20000 && num < 80000) return excelSerialToDate_(num);
      if (num >= 0 && num < 2)        return excelSerialToDate_(num);
    }
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
    if (m) return new Date(+m[3], +m[2]-1, +m[1], +(m[4]||'0'), +(m[5]||'0'));
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
    if (m) return new Date(+m[1], +m[2]-1, +m[3], +(m[4]||'0'), +(m[5]||'0'));
    var d3 = new Date(s);
    if (!isNaN(d3)) return d3;
  }
  return null;
}
function normalizeDate_(v) { var d = toDate_(v); if (!d) return null; return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function normalizeTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) return Utilities.formatDate(excelSerialToDate_(v), tz_(), 'HH:mm');
    if (v >= 0 && v < 2) {
      var total = Math.round(v * 24 * 60), hh = Math.floor(total / 60), mm = total % 60;
      return ('0'+hh).slice(-2) + ':' + ('0'+mm).slice(-2);
    }
  }
  if (typeof v === 'string') {
    var s = v.trim().replace('.', ':');
    var m = s.match(/^(\d{1,2})(?::|\.)(\d{2})$/);
    if (m) return ('0'+m[1]).slice(-2) + ':' + ('0'+m[2]).slice(-2);
    if (/^\d+(\.\d+)?$/.test(s)) return normalizeTime_(+s);
  }
  if (v === '' || v == null) return null;
  return String(v);
}
function normalizeText_(v) { if (v === '' || v == null) return null; return String(v).trim(); }
function displayDate_(yyyyMMddOrNull) {
  if (!yyyyMMddOrNull) return '-';
  var y = +yyyyMMddOrNull.slice(0, 4), m = +yyyyMMddOrNull.slice(5, 7), d = +yyyyMMddOrNull.slice(8, 10);
  return Utilities.formatDate(new Date(y, m - 1, d), tz_(), 'dd/MM/yyyy');
}
function displayTime_(hhmmOrNull) { return hhmmOrNull || '-'; }
/* ======================= DEADLINE HELPERS ======================= */
function daysLeft_(dueCell) {
  var d = toDate_(dueCell);
  if (!d) return null;
  var tz = tz_();
  var end = new Date(Utilities.formatDate(d, tz, 'yyyy/MM/dd') + ' 00:00:00');
  var start = new Date(Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd') + ' 00:00:00');
  return Math.round((end - start) / 86400000);
}
function deadlineBadge_(dueCell) {
  var dl = daysLeft_(dueCell);
  if (dl === null) return '';
  if (dl < 0) return ' [overdue]';
  if (dl === 0) return ' [due today]';
  return ' [' + dl + ' day(s) left]';
}
function isUrgentTomorrow_(dueCell) { return daysLeft_(dueCell) === 1; }
/* ======================= SCHEDULE STORAGE / UTIL ======================= */
// Persisted schedule times (defaults)
function getMorningTime_() { return PropertiesService.getDocumentProperties().getProperty('MORNING_TIME') || '06:00'; }
function getNightTime_()   { return PropertiesService.getDocumentProperties().getProperty('NIGHT_TIME')   || '21:00'; }
function setScheduleTimes_(morningHHmm, nightHHmm) {
  var p = PropertiesService.getDocumentProperties();
  p.setProperty('MORNING_TIME', morningHHmm);
  p.setProperty('NIGHT_TIME',   nightHHmm);
}
function parseHHmm_(s, fallback) {
  if (typeof s !== 'string') s = String(s || '');
  var m = s.trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return parseHHmm_(fallback || '00:00');
  var h = +m[1], mm = +m[2];
  if (h < 0 || h > 23 || mm < 0 || mm > 59) return parseHHmm_(fallback || '00:00');
  return { h: h, m: mm, hhmm: ('0'+h).slice(-2)+':'+('0'+mm).slice(-2) };
}
function recreateDailyTrigger_(handlerName, hh, mm) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(handlerName).timeBased().atHour(hh).nearMinute(mm).everyDays(1).create();
}
function nextRunInTz_(hhmm) {
  var tz = tz_();
  var parts = parseHHmm_(hhmm, hhmm);
  var now = new Date();
  var nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd') + ' 00:00:00';
  var start = new Date(Utilities.formatDate(new Date(nowStr), tz, 'yyyy/MM/dd HH:mm:ss'));
  start.setHours(parts.h); start.setMinutes(parts.m); start.setSeconds(0); start.setMilliseconds(0);
  if (start.getTime() <= now.getTime()) start = new Date(start.getTime() + 24*3600*1000);
  return Utilities.formatDate(start, tz, 'dd/MM/yyyy HH:mm');
}
// Tie custom day to morning time
function getCustomDayStartHour_() {
  var p = PropertiesService.getDocumentProperties();
  var v = p.getProperty('CUSTOM_DAY_START_HOUR');
  if (v == null) {
    var mt = parseHHmm_(getMorningTime_(), '06:00'); // default 6
    p.setProperty('CUSTOM_DAY_START_HOUR', String(mt.h));
    return mt.h;
  }
  var h = +v;
  if (!isFinite(h)) h = 6;
  return Math.min(23, Math.max(0, Math.floor(h)));
}
function todayCustomDate_() {
  var tz = tz_();
  var now = new Date();
  var startHour = getCustomDayStartHour_();
  var hourNow = +Utilities.formatDate(now, tz, 'HH');
  var base = (hourNow < startHour) ? new Date(now.getTime() - 24*3600*1000) : now;
  return Utilities.formatDate(base, tz, 'yyyy-MM-dd');
}
function isActedOnCustomDay_(actedDateCell, dayStr) {
  var d = normalizeDate_(actedDateCell);
  if (!d) return false;
  return d === dayStr;
}
/* ======================= TRIGGER CREATION (uses saved times) ======================= */
function createNightAndMorningTriggersUsingConfig() {
  purgeDuplicateTimeTriggers_();      
  var m = parseHHmm_(getMorningTime_(), '06:00');
  var n = parseHHmm_(getNightTime_(),   '21:00');
  recreateDailyTrigger_('runDueTodaySend', m.h, m.m);
  recreateDailyTrigger_('runNightlyAlert', n.h, n.m);
}

function removeNightAndMorningTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'runNightlyAlert' || h === 'runDueTodaySend') ScriptApp.deleteTrigger(t);
  });
}
/* ======================= TRIGGER STATUS DATA (for UI) ======================= */
function getTriggerStatusData() {
  var tz = tz_();
  var trig = ScriptApp.getProjectTriggers();
  var p = PropertiesService.getDocumentProperties();
  var NIGHT_FN   = 'runNightlyAlert';
  var MORNING_FN = 'runDueTodaySend';
  var hasNight    = trig.some(function (t) { return t.getHandlerFunction() === NIGHT_FN; });
  var hasMorning  = trig.some(function (t) { return t.getHandlerFunction() === MORNING_FN; });
  var hasRealtime = trig.some(function (t) { return t.getHandlerFunction() === 'onEdit'; });
  var tsNight   = p.getProperty('LAST_RUN_runNightlyAlert');
  var tsMorning = p.getProperty('LAST_RUN_runDueTodaySend');
  var editLock  = p.getProperty('EDIT_LOCK_ON') === '1';
  var morningTime = getMorningTime_();
  var nightTime   = getNightTime_();
  function fmt(ts){ return ts ? Utilities.formatDate(new Date(+ts), tz, 'dd/MM/yyyy HH:mm:ss') : '-'; }
  return {
    tz: tz,
    now: Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'),
    morning: {
      on: hasMorning, name: '06:00 Due-Today Send',
      time: morningTime, next: nextRunInTz_(morningTime), last: fmt(tsMorning),
      explain: 'Sends “Due Today” (Next Due Date = today), excluding rows where Progress = "Cancelled".'
    },
    night: {
      on: hasNight, name: '21:00 Nightly Alert',
      time: nightTime, next: nextRunInTz_(nightTime), last: fmt(tsNight),
      explain: 'Sends Daily Changes (Acted today & By matches Request by/Assigned to), then snapshots today for tomorrow.'
    },
    realtime: {
      on: hasRealtime, name: 'Realtime Edit Trigger',
      explain: 'Stamps Acted Date/By (watched columns) and logs per-edit diffs for the sidebar.'
    },
    editLock: editLock
  };
}
function getChangeLogCsv(opts) {
  opts = opts || {};
  var rows = getChangeLogData({ from: opts.from || '', to: opts.to || '' });
  var headers = ['When','Row','Field','Before','After','Task','By','Editor'];
  function esc(v){ return '"' + String(v == null ? '' : v).replace(/"/g,'""') + '"'; }
  var lines = [headers.map(esc).join(',')];
  rows.forEach(function(r){
    lines.push([r.when, r.row, r.fieldLabel, r.beforeDisp, r.afterDisp, r.task||'', r.by||'', r.who||''].map(esc).join(','));
  });
  var fname = 'change-log-' + (opts.from||'all') + '_to_' + (opts.to||'all') + '.csv';
  return { filename: fname, csv: lines.join('\n') };
}
/* ======================= SAVE TIMES (from UI) ======================= */
function saveScheduleTimes(morningHHmm, nightHHmm) {
  var m = parseHHmm_(morningHHmm, '06:00');
  var n = parseHHmm_(nightHHmm,   '21:00');
  setScheduleTimes_(m.hhmm, n.hhmm);
  // keep custom day boundary aligned with morning time
  PropertiesService.getDocumentProperties().setProperty('CUSTOM_DAY_START_HOUR', String(m.h));
  createNightAndMorningTriggersUsingConfig();
  return getTriggerStatusData();
}
/* ======================= SHOW TRIGGER STATUS (UI) ======================= */
function showTriggerStatus() {
  var html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Kanban Alert – Trigger Status</title>
<style>
  :root { --ok:#137333; --okbg:#e8f7ed; --bad:#a50e0e; --badbg:#fdecea; --muted:#666; }
  body { font: 13px Arial, sans-serif; margin:0; padding:14px; color:#222; }
  h2 { margin:0 0 10px; font-size:18px; }
  .meta { color: var(--muted); margin-bottom: 12px; }
  .grid { display:grid; grid-template-columns: 1fr; gap:10px; }
  .card { border:1px solid #e6e6e6; border-radius:10px; padding:12px; box-shadow:0 1px 0 rgba(0,0,0,.03); }
  .row { display:flex; justify-content:space-between; align-items:center; gap:8px; margin:6px 0; }
  .label { color: var(--muted); }
  .chip { display:inline-flex; align-items:center; gap:6px; padding:2px 8px; border-radius:999px; font-weight:600; font-size:12px; }
  .on  { background: var(--okbg); color: var(--ok); }
  .off { background: var(--badbg); color: var(--bad); }
  .small { color: var(--muted); font-size:12px; }
  .actions { margin-top:8px; display:flex; flex-wrap:wrap; gap:8px; }
  button { cursor:pointer; border:1px solid #ddd; background:#fff; border-radius:8px; padding:6px 10px; }
  button:hover { background:#f7f7f7; }
  input[type=time] { padding:4px 6px; border:1px solid #ddd; border-radius:6px; font:inherit; }
  .footer { display:flex; justify-content:flex-end; gap:8px; margin-top:14px; }
  .desc { margin-top:6px; color:#555; }
</style>
</head>
<body>
  <h2>Kanban Alert — Trigger Status</h2>
  <div class="meta" id="meta">Loading…</div>
  <div class="grid">
    <div class="card" id="night">
      <div class="row">
        <div><strong>Night Snapshot</strong></div>
        <div class="chip" id="nightChip">…</div>
      </div>
      <div class="row">
        <div class="label">Time (daily)</div>
        <div><input id="nightTime" type="time" step="60" value="20:00"></div>
      </div>
      <div class="row"><div class="label">Next run</div><div id="nightNext">-</div></div>
      <div class="row"><div class="label">Last run</div><div id="nightLast">-</div></div>
      <div class="desc" id="nightExplain">-</div>
      <div class="actions">
        <button onclick="runAlertNow()">Run alert now</button>
        <button onclick="runSnapshotNow()">Run snapshot now</button>
      </div>
    </div>
    <div class="card" id="morning">
      <div class="row">
        <div><strong>Morning Send</strong></div>
        <div class="chip" id="morningChip">…</div>
      </div>
      <div class="row">
        <div class="label">Time (daily)</div>
        <div><input id="morningTime" type="time" step="60" value="06:00"></div>
      </div>
      <div class="row"><div class="label">Next run</div><div id="morningNext">-</div></div>
      <div class="row"><div class="label">Last run</div><div id="morningLast">-</div></div>
      <div class="desc" id="morningExplain">-</div>
      <div class="actions">
        <button onclick="runSendNow()">Run send now</button>
        <button onclick="testSend()">Send test message</button>
      </div>
    </div>
    <div class="card" id="rt">
      <div class="row">
        <div><strong>Realtime Edit Trigger</strong></div>
        <div class="chip" id="rtChip">…</div>
      </div>
      <div class="small">Stamps Acted Date/By (watched columns) and logs step-by-step edits for the manual sidebar.</div>
      <div class="actions">
        <button onclick="enableRealtime()">Enable</button>
        <button onclick="disableRealtime()">Disable</button>
      </div>
    </div>
    <div class="card" id="locks">
      <div class="row"><strong>Edit Window</strong><span class="small">Allowed Morning → Night</span></div>
      <div class="row">
        <div class="small">With “Edit lock” on, edits are blocked between Night+1m and next Morning.</div>
        <div>
          <label class="small"><input id="lockToggle" type="checkbox" onchange="toggleLock(this.checked)"> Lock edits outside window</label>
        </div>
      </div>
    </div>
  </div>
  <div class="footer">
    <button onclick="saveTimes()">Save times</button>
    <button onclick="refresh()">Refresh</button>
    <button onclick="google.script.host.close()">Close</button>
  </div>
<script>
function badge(el, on) {
  el.textContent = on ? 'ON ✅' : 'OFF ❌';
  el.className = 'chip ' + (on ? 'on' : 'off');
}
function render(data){
  document.getElementById('meta').textContent = 'Timezone: ' + data.tz + ' • Now: ' + data.now;
  // Night
  badge(document.getElementById('nightChip'), data.night.on);
  document.getElementById('nightTime').value = data.night.time;
  document.getElementById('nightNext').textContent = data.night.next;
  document.getElementById('nightLast').textContent = data.night.last;
  document.getElementById('nightExplain').textContent = data.night.explain;
  // Morning
  badge(document.getElementById('morningChip'), data.morning.on);
  document.getElementById('morningTime').value = data.morning.time;
  document.getElementById('morningNext').textContent = data.morning.next;
  document.getElementById('morningLast').textContent = data.morning.last;
  document.getElementById('morningExplain').textContent = data.morning.explain;
  // Realtime + lock
  badge(document.getElementById('rtChip'), data.realtime.on);
  document.getElementById('lockToggle').checked = !!data.editLock;
}
function refresh(){ google.script.run.withSuccessHandler(render).getTriggerStatusData(); }
function saveTimes(){
  const m = document.getElementById('morningTime').value || '06:00';
  const n = document.getElementById('nightTime').value || '20:00';
  google.script.run.withSuccessHandler(render).saveScheduleTimes(m, n);
}
function enableRealtime(){ google.script.run.withSuccessHandler(refresh).createRealtimeEditTrigger(); }
function disableRealtime(){ google.script.run.withSuccessHandler(refresh).removeRealtimeTrigger(); }
function runSnapshotNow(){ google.script.run.withSuccessHandler(refresh).initializeSnapshotNightly(); }
function runSendNow(){ google.script.run.withSuccessHandler(refresh).runDueTodaySend(); }
function runAlertNow(){ google.script.run.withSuccessHandler(refresh).runNightlyAlert(); }
function testSend(){ google.script.run.testSend(); }
function toggleLock(on){ google.script.run.withSuccessHandler(refresh).setEditLock(on); }
refresh();
</script>
</body>
</html>
  `).setWidth(560).setHeight(640);

  SpreadsheetApp.getUi().showModalDialog(html, 'Trigger Status');
}

/* ======================= SNAPSHOT STORAGE (per-day) ======================= */
// Save & load snapshots by custom-day key (yyyy-MM-dd)
function saveSnapshotForDay_(day, obj, sh) {
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const prev = p.getProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf);
  p.setProperty('KANBAN_SNAPSHOT_' + day + suf, JSON.stringify(obj));
  if (prev && prev !== day) p.setProperty('KANBAN_SNAPSHOT_PREV_DATE' + suf, prev);
  p.setProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf, day);
  p.setProperty('KANBAN_SNAPSHOT_TS' + suf, String(Date.now()));
  pruneOldSnapshots_(sh);
}

function loadSnapshotByDay_(day, sh) {
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const raw = p.getProperty('KANBAN_SNAPSHOT_' + day + suf);
  return raw ? JSON.parse(raw) : {};
}
function loadLatestSnapshot_(sh) {
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const d = p.getProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf);
  return d ? { day: d, snap: loadSnapshotByDay_(d, sh) } : { day: null, snap: {} };
}

function saveSnapshot_(obj, sh) {
  const day = todayCustomDate_();
  saveSnapshotForDay_(day, obj, sh);
}

/* ======================= CHANGELOG (per-day) ======================= */
function appendChangelog_(entry) {
  const sh = entry.sh || SpreadsheetApp.getActiveSheet();
  const suf = sheetKeySuffix_(sh);
  const p = PropertiesService.getDocumentProperties();
  let arr = []; try { arr = JSON.parse(p.getProperty('KANBAN_CHANGELOG' + suf) || '[]'); } catch (_) {}

  // ensure who
  let who = getEditorEmail_(); try { who = Session.getActiveUser().getEmail() || who; } catch (_) {}
  entry.who = entry.who || who;

  // NEW: if uid missing, read from Task note
  try {
    if (!entry.uid && entry.row && entry.idx && entry.idx.task) {
      const note = sh.getRange(entry.row, entry.idx.task, 1, 1).getNote();
      if (note) entry.uid = note;
    }
  } catch(_){}

  delete entry.sh;
  arr.push(entry);
  // ////////////////////
  let s = JSON.stringify(arr);
  while (s.length > 8000 && arr.length > 1) { arr = arr.slice(Math.floor(arr.length/2)); s = JSON.stringify(arr); }
  ////////////////////////////
  p.setProperty('KANBAN_CHANGELOG' + suf, s);
}

function loadChangelogForToday_(sh) {
  const suf = sheetKeySuffix_(sh);
  const p = PropertiesService.getDocumentProperties();
  const raw = p.getProperty('KANBAN_CHANGELOG' + suf) || '[]';
  const arr = JSON.parse(raw);
  const tz = tz_(); const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return arr.filter(e => Utilities.formatDate(new Date(e.ts), tz, 'yyyy-MM-dd') === today);
}

/* ======================= REALTIME TRIGGERS ======================= */
function createRealtimeEditTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEdit' || t.getHandlerFunction() === 'onMyEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEdit')                      // <— use onEdit directly
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

function removeRealtimeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* ======================= EDIT LOCK ======================= */
function setEditLock(on) { PropertiesService.getDocumentProperties().setProperty('EDIT_LOCK_ON', on ? '1' : '0'); }
function isWithinEditWindow_() {
  var tz = tz_();
  var now = new Date();
  var morning = parseHHmm_(getMorningTime_(), '06:00');
  var night   = parseHHmm_(getNightTime_(),   '21:00');
  var hh = +Utilities.formatDate(now, tz, 'HH');
  var mm = +Utilities.formatDate(now, tz, 'mm');
  var minutes = hh*60 + mm;
  var start = morning.h*60 + morning.m;
  var end   = night.h*60 + night.m;
  if (end >= start) {
    return minutes >= start && minutes <= end; // same day window
  } else {
    // wrapped window (rare if night < morning)
    return minutes >= start || minutes <= end;
  }
}

/* ======================= PER-CELL CACHE (for robust oldValue) ======================= */
// 2) Cache ต่อแท็บ
function loadCellCache_(sh) {
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  if (p.getProperty('KANBAN_CACHE_DAY' + suf) !== todayCustomDate_()) return {};
  const key = 'KANBAN_CC' + suf;
  const json = CacheService.getDocumentCache().get(key);
  if (!json) return {};
  try { return JSON.parse(json); } catch (_) { return {}; }
}

function saveCellCache_(obj, sh) {
  const p = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const key = 'KANBAN_CC' + suf;
  const json = JSON.stringify(obj || {});
  if (json.length > 80000) CacheService.getDocumentCache().remove(key);
  else CacheService.getDocumentCache().put(key, json, 21600); // 6h
  p.setProperty('KANBAN_CACHE_DAY' + suf, todayCustomDate_());
}

function watchedSheets_() {
  const ss = SpreadsheetApp.getActive();
  const names = Array.isArray(CONFIG.SHEETS) && CONFIG.SHEETS.length
    ? CONFIG.SHEETS : [/* none */];
  return names.map(n => ss.getSheetByName(n)).filter(Boolean);
}

function isWatchedSheet_(sh) {
  if (!sh) return false;
  const names = Array.isArray(CONFIG.SHEETS) && CONFIG.SHEETS.length
    ? CONFIG.SHEETS
    : []; // no single-sheet fallback
  return names.indexOf(sh.getName()) !== -1;
}

function sheetKeySuffix_(sh) { return '_S' + sh.getSheetId(); }

/* ======================= onEdit: stamp & log ======================= */
function onEdit(e) {
  
  var props = PropertiesService.getDocumentProperties();
  try {
    if (!e || !e.triggerUid) return;
    if (!e || !e.range) return;
    var rng = e.range;
    var sh  = rng.getSheet();
    if (!isWatchedSheet_(sh)) return;    
    if (rng.getRow() <= CONFIG.HEADER_ROW) {
      try {
        var cacheH = loadCellCache_(sh), vH = rng.getValues();
        var tH = rng.getRow(), lH = rng.getColumn();
        for (var rH = 0; rH < rng.getNumRows(); rH++) {
          for (var cH = 0; cH < rng.getNumColumns(); cH++) {
            cacheH[(tH + rH) + ':' + (lH + cH)] = vH[rH][cH];
          }
        }
        saveCellCache_(cacheH, sh);
      } catch (_) {}
      return;
    }
    if (isReverting_()) { setRevertingFlag_(false); return; }
    var lockOn        = props.getProperty('EDIT_LOCK_ON') === '1';
    var allowedNow    = isWithinEditWindow_();
    var isInstallable = !!(e && e.triggerUid); // simple triggers have no triggerUid

    if (lockOn && !allowedNow) {
      // Always at least a toast (simple triggers can't open modals)
      try { SpreadsheetApp.getActive().toast('Edits are locked. Your change was undone.'); } catch (_) {}
      // Pretty modal if installable trigger
      if (isInstallable) { try { showLockedEditModal_(); } catch (_) {} }
      setRevertingFlag_(true);
      try {
        revertFromCacheOrOld_(rng, e);  // uses oldValue for single cells or our cache for ranges
        SpreadsheetApp.flush();
      } finally {
        setRevertingFlag_(false);
      }
      return;
    }
    var lastCol = sh.getLastColumn();
    var header  = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
    var idx     = indexMap_(header); // throws if headers mismatch
    var beforeCache = loadCellCache_(sh);                  // cell-level "before" values
    var vals   = rng.getValues();
    var top    = rng.getRow(), left = rng.getColumn();
    var rows   = rng.getNumRows(), cols = rng.getNumColumns();
    var rowValsCache = Object.create(null);
    function getRowVals(rowNum) {
      var k = String(rowNum);
      if (!rowValsCache[k]) rowValsCache[k] = sh.getRange(rowNum, 1, 1, lastCol).getValues()[0];
      return rowValsCache[k];
    }
    var rowsToStamp = Object.create(null);
    for (var r = 0; r < rows; r++) {
      var rowNum = top + r;
      for (var c = 0; c < cols; c++) {
        var colNum = left + c;
        var headerNorm = _normHeader(header[colNum - 1] || '');
        var fieldKey   = lookupFieldKeyByHeader_(headerNorm);
        if (!fieldKey) continue;

        var key    = rowNum + ':' + colNum;
        var before = (key in beforeCache) ? beforeCache[key] : undefined;
        var after  = vals[r][c];
        if (String(before) === String(after)) continue;

        // Log this single-field change
        try {
          var rowVals = getRowVals(rowNum);
          appendChangelog_({
            ts:   Date.now(),
            row:  rowNum,
            uid:  sh.getRange(rowNum, idx.task, 1, 1).getNote(),   // ← NEW
            field: fieldKey,
            from: before,
            to:   after,
            task: rowVals[idx.task - 1],
            by:   rowVals[idx.by   - 1] || '',
            idx:  idx,                                             // help appendChangelog_ fallback
            sh:   sh
          });
        } catch (_) {}

        // Mark row to stamp if the field is watched (but not the stamp fields themselves)
        if (CONFIG.WATCH_FIELDS[fieldKey] && fieldKey !== 'changeDate' && fieldKey !== 'by') {
          rowsToStamp[rowNum] = true;
        }
      }
    }
    var now = new Date();
    Object.keys(rowsToStamp).forEach(function (k) {
      var rn = +k;
      var rowVals = getRowVals(rn);
      var currentBy  = rowVals[idx.by         - 1];
      var requestBy  = rowVals[idx.requestBy  - 1];
      var assignedTo = rowVals[idx.assignedTo - 1];
      var newBy = chooseBy_(currentBy, requestBy, assignedTo);
      sh.getRange(rn, idx.changeDate, 1, 1).setValue(now);
      if (newBy !== currentBy) sh.getRange(rn, idx.by, 1, 1).setValue(newBy || '');
      beforeCache[rn + ':' + idx.changeDate] = now;
      if (newBy !== currentBy) beforeCache[rn + ':' + idx.by] = newBy || '';
    });
    vals = rng.getValues();
    for (var r2 = 0; r2 < rows; r2++) {
      for (var c2 = 0; c2 < cols; c2++) {
        beforeCache[(top + r2) + ':' + (left + c2)] = vals[r2][c2];
      }
    }
    saveCellCache_(beforeCache,sh);
  } catch (err) {
    try { setRevertingFlag_(false); } catch (_) {}
  }
}

function initializeSnapshotNightly() {
  // Snapshot ALL watched sheets (those listed in CONFIG.SHEETS),
  // rotate "prev" correctly per-sheet, prune older days,
  // and give a clear success toast.
  const actedDay = todayCustomDate_();              // yyyy-MM-dd at your custom morning boundary
  watchedSheets_().forEach(sh => {
    const data = readTableForSheet_(sh);            // read table for THIS sheet
    const snap = buildSnapshot_(data);              // normalize (uses UID as key)
    saveSnapshotForDay_(actedDay, snap, sh);        // rotates prev + prunes others for THIS sheet
  });

  PropertiesService.getDocumentProperties()
    .setProperty('LAST_RUN_initializeSnapshotNightly', String(Date.now()));

  try { SpreadsheetApp.getActive().toast('Snapshot completed for ALL watched tabs ✅'); } catch (_) {}
}

function indexMap_(headerValues) {
  var normRow = headerValues.map(_normHeader);
  var idx = {};
  for (var k in CONFIG.HEADERS) {
  var target = _normHeader(CONFIG.HEADERS[k]);
  var i = normRow.indexOf(target);
  if (i === -1) throw new Error('Header "' + CONFIG.HEADERS[k] + '" not found at row ' + CONFIG.HEADER_ROW + '. Seen: ' + headerValues.join(' | '));
    idx[k] = i + 1;
  }
  return idx;
}
function buildSnapshot_(data) {
  var snap = {};
  data.forEach(function (item) {
    const key = String(item.uid || item.row);  // ← use UID first
    snap[key] = {
      uid:         normalizeText_(item.uid),   // keep for convenience
      task:        normalizeText_(item.task),
      requestBy:   normalizeText_(item.requestBy),
      assignedTo:  normalizeText_(item.assignedTo),
      resources:   normalizeText_(item.resources),
      startDate:   normalizeDate_(item.startDate),
      changeDate:  normalizeDate_(item.changeDate),
      nextDueDate: normalizeDate_(item.nextDueDate),
      dueDate:     normalizeDate_(item.dueDate),
      meetingTime: normalizeTime_(item.meetingTime),
      note:        normalizeText_(item.note),
      progress:    normalizeText_(item.progress),
      by:          normalizeText_(item.by)
    };
  });
  return snap;
}


/* ======================= NAME / MATCH HELPERS ======================= */
function splitNames_(v) {
  if (v == null) return [];
  var s = String(v).replace(/\u00A0/g, ' ').trim();
  if (!s) return [];
  return s.split(/\s*(?:,|\/|\||และ|and|or)\s*/i)
          .map(function(t){ return t.trim(); })
          .filter(function(t){ return !!t; });
}
function byMatchesRow_(r) {
  var by = normalizeText_(r.by);
  if (!by) return false;
  var rb = splitNames_(r.requestBy);
  var ab = splitNames_(r.assignedTo);
  return rb.indexOf(by) !== -1 || ab.indexOf(by) !== -1;
}

/* ======================= DIFF UTILS ======================= */
function normForType_(type, value) {
  if (type === 'date') return normalizeDate_(value);
  if (type === 'time') return normalizeTime_(value);
  if (type === 'text') return normalizeText_(value);
  if (type === 'datetime') {
    var d = toDate_(value);
    return d ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd HH:mm') : null;
  }
  return value == null ? null : String(value);
}
function dispForType_(type, normVal) {
  if (type === 'date') return normVal ? displayDate_(normVal) : '""';
  if (type === 'time') return normVal ? displayTime_(normVal) : '""';
  if (type === 'text') return (normVal == null || normVal === '') ? '""' : '"' + normVal + '"';
  if (type === 'datetime') return normVal ? normVal : '""';
  return (normVal == null || normVal === '') ? '""' : '"' + String(normVal) + '"';
}
function diffLinesForRow_(snapRow, currentRow) {
  function isEmpty(type, v){
    return v == null || v === '';
  }
  var lines = [];
  if (!snapRow) snapRow = {};
  if (!currentRow) currentRow = {};

  DIFF_FIELDS.forEach(function (f) {
    if (f.key === 'by') return;
    var before = normForType_(f.type, snapRow[f.key]);
    var after  = normForType_(f.type, currentRow[f.key]);

    if ((before || '') !== (after || '')) {
      // ✅ skip “empty → value” (initial fill) so it won’t appear in preview/sends
      if (isEmpty(f.type, before) && !isEmpty(f.type, after)) return;

      lines.push('- ' + f.label() + ': ' +
        dispForType_(f.type, before) + ' \u2192 ' +
        dispForType_(f.type, after));
    }
  });
  return lines;
}


function rowFromSnapRow_(rownum, s) {
  return {
    row: rownum,
    task: s.task,
    requestBy: s.requestBy,
    assignedTo: s.assignedTo,
    resources: s.resources,
    startDate: s.startDate,
    changeDate: s.changeDate,
    nextDueDate: s.nextDueDate,
    dueDate: s.dueDate,
    progress: s.progress,
    meetingTime: s.meetingTime,
    note: s.note,                   
    by: s.by
  };
}

// Read all rows from a specific sheet (respects CONFIG.HEADER_ROW)
function readTableForSheet_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW) return [];
  const header = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const idx = indexMap_(header);
  const startRow = CONFIG.HEADER_ROW + 1;
  const rows = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  // NEW: ensure/load UIDs from Task notes
  const uids = ensureRowUidsForSheet_(sh, idx);

  return rows.map((r, off) => ({
    row: startRow + off,
    uid: uids[off],
    task: r[idx.task - 1],
    requestBy: r[idx.requestBy - 1],
    assignedTo: r[idx.assignedTo - 1],
    resources: r[idx.resources - 1],
    startDate: r[idx.startDate - 1],
    changeDate: r[idx.changeDate - 1],
    nextDueDate: r[idx.nextDueDate - 1],
    dueDate: r[idx.dueDate - 1],
    progress: r[idx.progress - 1],
    meetingTime: r[idx.meetingTime - 1],
    note: r[idx.note - 1],
    by: r[idx.by - 1]
  }));
}

// ==== Idempotent guards (วันละครั้ง) ====
function alreadySentForDay_(k, ymd){
  const p = PropertiesService.getDocumentProperties();
  return p.getProperty(k) === ymd;
}
function markSentForDay_(k, ymd){
  PropertiesService.getDocumentProperties().setProperty(k, ymd);
}

function runNightlyAlert() {
  // 1) จับ lock และถือไว้จนจบ
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p   = PropertiesService.getDocumentProperties();
    const key = 'RUN_GUARD_runNightlyAlert';
    const now  = Date.now();
    const last = +(p.getProperty(key) || '0');
    if (now - last < 8 * 60 * 1000) return;     // กันช็อตซ้อนช่วงสั้น ๆ
    p.setProperty(key, String(now));

    // 2) Guard แบบ "วันละครั้ง" (ยึด boundary ตาม todayCustomDate_())
    const actedDay = todayCustomDate_();
    const DAY_KEY_NIGHT = 'NIGHTLY_SENT_DAY';
    if (p.getProperty(DAY_KEY_NIGHT) === actedDay) return;

    // 3) สร้างข้อความ + ทำ snapshot (โค้ดเดิมคงไว้)
    const tz = tz_();
    const lines = [
      'Date : ' + Utilities.formatDate(toDate_(actedDay), tz, 'dd/MM/yyyy'),
      '- Daily Change',
      CONFIG.SEPARATOR_LINE
    ];
    let any = false;

    watchedSheets_().forEach(sh => {
      const prevRaw = getStartOfDayBaseline_(sh, actedDay);
      const nowRows = readTableForSheet_(sh);
      const nowSnap = buildSnapshot_(nowRows);
      const prev    = reindexSnapshotToUid_(prevRaw, nowRows);

      // ล่าสุดต่อ field ของวันนี้ (keyed by UID/row)
      const latestByKey = {};
      loadChangelogForToday_(sh).sort((a,b)=>a.ts-b.ts).forEach(e=>{
        const k = String(e.uid || e.row);
        const type   = fieldTypeForKey_(e.field);
        const before = normForType_(type, e.from);
        const after  = normForType_(type, e.to);
        if ((before||'')===(after||'')) return;
        (latestByKey[k]||(latestByKey[k]={}))[e.field] = { before, after, type };
      });

      // UID → row ปัจจุบัน
      const uidToRow = Object.create(null);
      nowRows.forEach(r => { uidToRow[keyForRow_(r)] = r.row; });

      // รวบรวมแถวที่ "Acted today" และมี diff จริง
      const rows = [];
      Object.keys(nowSnap).forEach(k => {
        const after = nowSnap[k];
        if (!isActedOnCustomDay_(after.changeDate, actedDay)) return;

        const liveRowNum = uidToRow[k] || 0;
        const tmpRow = rowFromSnapRow_(liveRowNum, after);
        if (CONFIG.REQUIRE_BY_MATCH && !byMatchesRow_(tmpRow)) return;

        const diffs = diffLinesForRow_(prev[k], after); // baseline → now
        let changeLines = diffs;

        // ถ้า baseline ไม่มี diff ให้ fallback เป็น latestByKey ของวันนี้
        const lat = latestByKey[k];
        if (!changeLines.length && lat) {
          const alt = [];
          DIFF_FIELDS.forEach(f=>{
            const v = lat[f.key]; if (!v) return;
            if ((v.before||'')===(v.after||'')) return;
            alt.push('- ' + f.label() + ': ' +
              dispForType_(f.type, v.before) + ' \u2192 ' + dispForType_(f.type, v.after));
          });
          changeLines = alt;
        }
        if (!changeLines.length) return;

        rows.push({ row: liveRowNum, after, changeLines });
      });

      // snapshot ต่อให้ไม่มี rows ที่จะส่ง
      saveSnapshotForDay_(actedDay, nowSnap, sh);

      if (!rows.length) return;

      any = true;
      lines.push('📄 Sheet: ' + sh.getName());
      rows.forEach((r,i)=>{
        const full = rowFromSnapRow_(r.row, r.after);
        lines.push(
          '#' + full.row + ' · ' + String(full.task||'-').trim() + ' ' + String(full.progress||'-').trim(),
          'Request: '      + (String(full.requestBy  || '-').trim()),
          'Assigned to: '  + (String(full.assignedTo || '-').trim()),
          'Resources: '    + (String(full.resources  || '-').trim()),
          'Start: '        + displayDate_(normalizeDate_(full.startDate)),
          'Next Due: '     + displayDate_(normalizeDate_(full.nextDueDate)),
          'Deadline: '     + displayDate_(normalizeDate_(full.dueDate)) + (deadlineBadge_(full.dueDate)||''),
          'Meeting Time: ' + displayTime_(normalizeTime_(full.meetingTime)),
          'Note: '         + (String(full.note || '-').trim()),
          'Change:'
        );
        (r.changeLines.length ? r.changeLines : ['—']).forEach(x=>lines.push(x));
        if (i < rows.length-1 && CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
      });
    });

    // 4) ส่ง / ไม่ส่ง ตามนโยบาย และ mark วัน
    if (!any && !CONFIG.SEND_IF_EMPTY) {
      // ไม่ส่ง และ "ไม่ mark" เพื่อให้ยัง Run now ได้ถ้ามีการแก้หลังจากนี้ในคืนเดียวกัน
      p.setProperty('LAST_RUN_runNightlyAlert', String(Date.now()));
      return;
    }
    if (!any) lines.push('No changes in today’s.');
    if (CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
    lines.push('Goodnight💤');

    sendLineMulti_(lines.join('\n'));
    p.setProperty('LAST_RUN_runNightlyAlert', String(Date.now()));
    p.setProperty(DAY_KEY_NIGHT, actedDay);   // นับว่าวันนี้ “ส่งแล้ว”
  } finally {
    lock.releaseLock();
  }
}

function previewNightlyAlertMessage(){
  const actedDay = todayCustomDate_();
  const tz = tz_();
  const lines = [
    'Date : ' + Utilities.formatDate(toDate_(actedDay), tz, 'dd/MM/yyyy'),
    '- Daily Change',
    CONFIG.SEPARATOR_LINE
  ];
  let any = false;

  watchedSheets_().forEach(sh => {
    const prevRaw = getStartOfDayBaseline_(sh, actedDay);
    const nowRows = readTableForSheet_(sh);
    const nowSnap = buildSnapshot_(nowRows);
    const prev    = reindexSnapshotToUid_(prevRaw, nowRows);   // ← add this


  // collect today’s latest per-field changes (keyed by UID)
  const latestByKey = {};
  loadChangelogForToday_(sh).sort((a,b)=>a.ts-b.ts).forEach(e=>{
    const k = String(e.uid || e.row);
    const type = fieldTypeForKey_(e.field);
    const before = normForType_(type, e.from);
    const after  = normForType_(type, e.to);
    if ((before||'')===(after||'')) return;
    (latestByKey[k]||(latestByKey[k]={}))[e.field] = {before, after, type};
  });

  // UID → current row #
  const uidToRow = Object.create(null);
  nowRows.forEach(r => { uidToRow[keyForRow_(r)] = r.row; });

  function keepImportant(v){
    const bEmpty = (v.before==null || v.before==='');
    const aEmpty = (v.after==null  || v.after==='');
    return !(bEmpty && !aEmpty) && ((v.before||'') !== (v.after||''));
  }
  function linesFromLatest(lat){
    const out = [];
    DIFF_FIELDS.forEach(f=>{
      const v = lat[f.key]; if (!v) return;
      if (!keepImportant(v)) return;
      out.push('- ' + f.label() + ': ' +
        dispForType_(f.type, v.before) + ' \u2192 ' + dispForType_(f.type, v.after));
    });
    return out;
  }
  const rows = [];
  Object.keys(nowSnap).forEach(k=>{
    const after = nowSnap[k];           // snapshot row
    if (!isActedOnCustomDay_(after.changeDate, actedDay)) return;

    const liveRowNum = uidToRow[k] || 0;
    const tmpRow = rowFromSnapRow_(liveRowNum, after);
    if (CONFIG.REQUIRE_BY_MATCH && !byMatchesRow_(tmpRow)) return;

    const diffs = diffLinesForRow_(prev[k], after);
    const lat   = latestByKey[k];
    const alt   = lat ? linesFromLatest(lat) : [];

    if (diffs.length || alt.length) {
      rows.push({ uid: k, row: liveRowNum, after, changeLines: diffs.length ? diffs : alt });
    }
  });

    if (!rows.length) return;
    any = true;
    lines.push('📄 Sheet: ' + sh.getName());
    rows.forEach((r,i)=>{
      const full = rowFromSnapRow_(r.row, r.after);
      lines.push(
        '#' + full.row + ' · ' + String(full.task||'-').trim() + ' ' + String(full.progress||'-').trim(),
        'Request: '      + (String(full.requestBy  || '-').trim()),
        'Assigned to: '  + (String(full.assignedTo || '-').trim()),
        'Resources: '    + (String(full.resources  || '-').trim()),
        'Start: '        + displayDate_(normalizeDate_(full.startDate)),
        'Next Due: '     + displayDate_(normalizeDate_(full.nextDueDate)),
        'Deadline: '     + displayDate_(normalizeDate_(full.dueDate)) + (deadlineBadge_(full.dueDate)||''),
        'Meeting Time: ' + displayTime_(normalizeTime_(full.meetingTime)),
        'Note: '         + (String(full.note || '-').trim()),
        'Change:'
      );
      (r.changeLines.length ? r.changeLines : ['—']).forEach(x=>lines.push(x));
      if (i < rows.length-1 && CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
    });
    if (CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
  });

  if (!any) lines.push('No changes in today’s.');
  if (CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
  lines.push('Goodnight💤');

  const html = HtmlService.createHtmlOutput(
    '<pre style="white-space:pre-wrap;font:12px ui-monospace">' +
    lines.join('\n').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') +
    '</pre>'
  ).setWidth(700).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview Change Summary Notification.');
}

function formatDueListMessage_(rows, targetYmd, isPreview) {
  var tz = tz_();
  var ymd = targetYmd || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var dateDisp = Utilities.formatDate(toDate_(ymd), tz, 'dd/MM/yyyy');
  var H = CONFIG.HEADERS;
  rows = rows.slice().sort(function (a, b) {
    function t(v) { var s = normalizeTime_(v); return s ? s : '99:99'; }
    var ta = t(a.meetingTime), tb = t(b.meetingTime);
    if (ta < tb) return -1;
    if (ta > tb) return 1;
    return (a.row || 0) - (b.row || 0);
  });
  var lines = [
    'Next Due Date : ' + dateDisp,
    '- Good Morning ☀️',
    'Today Reminder',
    CONFIG.SEPARATOR_LINE
  ];
  if (!rows.length) {
    lines.push('(no items due today)');
  } else {
    rows.forEach(function (r, i) {
      var taskName  = (String(r.task || '').trim() || '-');
      var status    = (String(r.progress || '').trim() || '-');
      var startDisp = displayDate_(normalizeDate_(r.startDate));
      var nextDisp  = displayDate_(normalizeDate_(r.nextDueDate));
      var dueDisp   = displayDate_(normalizeDate_(r.dueDate));
      var mt        = displayTime_(normalizeTime_(r.meetingTime));
      var noteTxt   = String(r.note || '').trim();
      if (noteTxt.length > 180) noteTxt = noteTxt.slice(0, 177) + '…';
      var rawBadge = deadlineBadge_(r.dueDate) || '';
      var badge = (function () {
        if (!rawBadge) return '';
        if (/\boverdue\b/i.test(rawBadge)) return ' ⚠️ overdue';
        if (/due today/i.test(rawBadge))   return ' 🟡 today';
        var m = rawBadge.match(/\[(\d+)\s+day/);
        return m ? (' 🟢 ' + m[1] + 'd left') : (' ' + rawBadge);
      })();
      lines.push('#' + r.row + ' · ' + taskName);
      lines.push(H.requestBy   + ': ' + (r.requestBy  || '-'));
      lines.push(H.assignedTo  + ': ' + (r.assignedTo || '-'));
      lines.push(H.resources   + ': ' + (r.resources  || '-'));
      lines.push(H.startDate   + ': ' + startDisp);
      lines.push(H.nextDueDate + ': ' + nextDisp);
      lines.push(H.dueDate     + ': ' + dueDisp + badge);
      if (mt && mt !== '-') lines.push(H.meetingTime + ': ' + mt + ' ⏰');
      lines.push(H.progress    + ': ' + status);
      lines.push(H.note        + ': ' + (noteTxt ? noteTxt : '-'));
      if (i < rows.length - 1 && CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
    });
  }
  lines.push(CONFIG.SEPARATOR_LINE);
  lines.push('Have a good day');
  return lines.join('\n');
}

function previewDueListMessage(){
  const tz = tz_();
  const ymd = Utilities.formatDate(new Date(Date.now()+24*3600*1000), tz, 'yyyy-MM-dd');

  let rowsAll = [];
  watchedSheets_().forEach(sh => {
    rowsAll = rowsAll.concat(getDueOnDate_ForSheet_(sh, ymd));
  });

  const msg = formatDueListMessage_(rowsAll, ymd, true);
  const html = HtmlService.createHtmlOutput(
    '<pre style="white-space:pre-wrap;font:12px ui-monospace">' +
    msg.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') +
    '</pre>'
  ).setWidth(700).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview Tomorrow Task Notification.');
}

function purgeDuplicateRealtimeTriggers_(){
  let kept = false;
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'onEdit' || f === 'onMyEdit') {
      if (!kept) { kept = true; }
      else { ScriptApp.deleteTrigger(t); }     // delete extras
    }
  });
}

function runDueTodaySend(){
  // 1) จับ lock และถือไว้จนจบ
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p   = PropertiesService.getDocumentProperties();
    const key = 'RUN_GUARD_runDueTodaySend';
    const now  = Date.now();
    const last = +(p.getProperty(key) || '0');
    if (now - last < 8 * 60 * 1000) return;
    p.setProperty(key, String(now));

    const tz = tz_();
    const todayYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // 2) Guard แบบ "วันละครั้ง"
    const DAY_KEY_MORNING = 'MORNING_SENT_DAY';
    if (p.getProperty(DAY_KEY_MORNING) === todayYmd) return;

    // 3) รวมรายการ due วันนี้ (โค้ดเดิมคงไว้)
    let rowsAll = [];
    watchedSheets_().forEach(sh => {
      rowsAll = rowsAll.concat(getDueOnDate_ForSheet_(sh, todayYmd));
    });

    // 4) ส่ง / ไม่ส่ง และ mark วัน
    if (!rowsAll.length && !CONFIG.SEND_IF_EMPTY) {
      // ไม่ส่ง และ "ไม่ mark" เพื่อให้ยัง Run now ได้ในเช้าวันเดียวกัน หากมีงานโผล่ทีหลัง
      p.setProperty('LAST_RUN_runDueTodaySend', String(Date.now()));
      return;
    }
    const msg = formatDueListMessage_(rowsAll, todayYmd, false);
    sendLineMulti_(msg);
    p.setProperty('LAST_RUN_runDueTodaySend', String(Date.now()));
    p.setProperty(DAY_KEY_MORNING, todayYmd);
  } finally {
    lock.releaseLock();
  }
}


function getDueOnDate_ForSheet_(sh, targetYmd) {
  return readTableForSheet_(sh).filter(r => {
    const nextDue = normalizeDate_(r.nextDueDate);
    const cancelled = String(r.progress || '').trim().toLowerCase() === 'cancelled';
    return !cancelled && nextDue === targetYmd;
  });
}

/* ======================= SIDEBAR (manual review) ======================= */
function openReviewSidebar() {
  var html = HtmlService.createHtmlOutput(buildSidebarHtml_()).setTitle('Manually Notification.');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPreviewChanges() {
  const actedDay = todayCustomDate_();
  const out = [];

  watchedSheets_().forEach(sh => {
    const data = readTableForSheet_(sh);

    // === BASELINE (yesterday at day-start) — reindex to UID ===
    const prevRaw = getStartOfDayBaseline_(sh, actedDay) || {};
    const prev    = reindexSnapshotToUid_(prevRaw, data);  // <-- IMPORTANT

    // base rows: acted today (+ optional by-match)
    let baseRows;
    if (CONFIG.PREVIEW_MODE === 'BY_DIFF') {
      // Compare against UID-keyed baseline
      baseRows = data.filter(r => diffLinesForRow_(prev[keyForRow_(r)], r).length > 0);
    } else { // BY_CHANGE_DATE
      baseRows = data.filter(r => isActedOnCustomDay_(r.changeDate, actedDay));
    }
    if (CONFIG.REQUIRE_BY_MATCH) baseRows = baseRows.filter(byMatchesRow_);

    // per-field changelog (today only), keyed by UID (fallback: row for legacy)
    const log = loadChangelogForToday_(sh);
    const allowed = Object.create(null); baseRows.forEach(r => allowed[keyForRow_(r)] = true);
    const latest = {};
    log.sort((a,b)=>a.ts-b.ts).forEach(e=>{
      const k = String(e.uid || e.row);
      if (!allowed[k]) return;
      const type = fieldTypeForKey_(e.field);
      const before = normForType_(type, e.from);
      const after  = normForType_(type, e.to);
      if ((before||'')===(after||'')) return;
      (latest[k]||(latest[k]={}))[e.field] = {before, after, type};
    });

    baseRows.forEach(ch => {
      const key = keyForRow_(ch);             // UID preferred, else row
      const l = latest[key];

      // 1) True diffs between baseline (UID-keyed) and now
      const snapDiffs = diffLinesForRow_(prev[key], ch);

      // 2) If we have today’s per-field logs, prefer those lines
      let lines = [];
      if (l) {
        DIFF_FIELDS.forEach(f=>{
          const v = l[f.key]; if (!v) return;
          if ((v.before||'') === (v.after||'')) return;
          lines.push('- ' + f.label() + ': ' +
            dispForType_(f.type, v.before) + ' → ' + dispForType_(f.type, v.after));
        });
      } else {
        lines = snapDiffs.slice();
      }

      if (!lines.length) return; // nothing truly changed → skip

      out.push({
        sheet: sh.getName(),
        row: ch.row,
        task: ch.task,
        requestBy: ch.requestBy,
        status: ch.progress,
        by: ch.by,
        startDisp: displayDate_(normalizeDate_(ch.startDate)),
        nextDisp:  displayDate_(normalizeDate_(ch.nextDueDate)),
        dueDisp:   displayDate_(normalizeDate_(ch.dueDate)),
        badges: (isUrgentTomorrow_(ch.dueDate) ? ' 🔴 [URGENT]' : '') + deadlineBadge_(ch.dueDate),
        changeBits: lines.join('\n'),
        preselect: true
      });
    });
  });

  return out;
}

function buildSidebarHtml_() {
  var html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    body { font: 13px Arial, sans-serif; padding: 12px; }
    h2 { margin: 0 0 8px; font-size: 16px; }
    .muted { color: #666; }
    .card { border: 1px solid #e5e5e5; border-radius: 8px; padding: 8px 10px; margin-bottom: 8px; }
    .rowhead { font-weight: bold; margin-bottom: 4px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 12px; margin: 6px 0; }
    .changes { margin: 4px 0; white-space: pre-wrap; }
    .controls { display:flex; gap:6px; margin-top: 10px; }
    textarea { width: 100%; min-height: 48px; }
    .small { font-size: 12px; }
    .tag { background:#f2f2f2; padding:1px 6px; border-radius: 999px; margin-left: 6px; }
    .toolbar { display:flex; gap:6px; margin-bottom: 8px; }
    button { cursor:pointer; }
  </style>
</head>
<body>
  <h2>Sent Urgent: Today’s Changes</h2>
  <div class="toolbar">
    <button onclick="selectAll(true)">Select all</button>
    <button onclick="selectAll(false)">Clear</button>
    <button onclick="reload()">Refresh</button>
  </div>
  <div id="list">Loading…</div>
  <div class="controls">
    <button onclick="send()">Send selected</button>
    <button onclick="closeMe()">Close</button>
  </div>
<script>
let items = [];
function reload(){
  google.script.run
    .withSuccessHandler(render)
    .withFailureHandler(err => {
      const box = document.getElementById('list');
      const msg = (err && err.message) ? err.message : String(err);
      box.innerHTML = '<div class="muted">Error: ' + escapeHtml(msg) + '</div>';
    })
    .getPreviewChanges();
}
function selectAll(v){ document.querySelectorAll('input[type=checkbox][data-k]').forEach(cb => cb.checked = v); }
function closeMe(){ google.script.host.close(); }
function render(data){
  items = data;
  const box = document.getElementById("list");
  if (!data.length) { box.innerHTML = '<div class="muted">No matching rows.</div>'; return; }

  box.innerHTML = data.map((it, idx) => {
    const chk = '<input type="checkbox" data-k="'+idx+'" ' + (it.preselect ? "checked" : "") + '/>';
    const hdr = '📄 '+escapeHtml(it.sheet||"-")+' · #'+it.row+' · '+escapeHtml(it.task||"-")+
            ' <span class="tag">'+escapeHtml(it.status||"-")+'</span>';
    const grid = [
      '<div class="grid small">',
      '  <div><b>Request</b><br>'+escapeHtml(it.requestBy||"-")+'</div>',
      '  <div><b>Start</b><br>'+escapeHtml(it.startDisp||"-")+'</div>',
      '  <div><b>Next Due</b><br>'+escapeHtml(it.nextDisp||"-")+'</div>',
      '  <div><b>Deadline</b><br>'+escapeHtml(it.dueDisp||"-")+' '+escapeHtml(it.badges||"")+'</div>',
      '</div>'
    ].join("\\n");
    const changes = '<div class="changes small"><b>Change:</b>\\n'+escapeHtml(it.changeBits||"-")+'</div>';
    const edit = [
      '<details>',
      '  <summary class="small">Edit message for this row (optional)</summary>',
      '  <textarea data-edit="'+idx+'" placeholder="Custom text to replace the \\"Change\\" section"></textarea>',
      '</details>'
    ].join("\\n");
    return '<div class="card">'+chk+' <span class="rowhead">'+hdr+'</span>'+grid+changes+edit+'</div>';
  }).join("");
}
function send(){
  const payload = [];
  document.querySelectorAll('input[type=checkbox][data-k]').forEach(cb=>{
    if (!cb.checked) return;
    const idx = +cb.getAttribute("data-k");
    const t = document.querySelector('textarea[data-edit="'+idx+'"]');
    const custom = t && t.value.trim() ? t.value.trim() : null;
    payload.push({ idx, custom });
  });
  if (!payload.length){ alert("Select at least one row to send."); return; }
  google.script.run.withSuccessHandler(()=>{
    alert("Sent to LINE.");
    google.script.host.close();
  }).sendSelectedChanges(payload, items);
}
function escapeHtml(s){
  return String(s||"").replace(/[&<>\\\"\\']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", "\\\"":"&quot;", "'":"&#39;"
  })[c]);
}
reload();
</script>
</body>
</html>`;
  return html;
}

function sendSelectedChanges(selection, items) {
  var ts = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm');
  var lines = ['Urgent: Today’s Changes', '', ts]; // ← change header for MANUAL only
  selection.forEach(function (sel, i) {
    var it = items[sel.idx];
    var changeLine = (sel.custom && sel.custom.trim()) ? sel.custom.trim() : (it.changeBits || '-');
    if (i > 0 && CONFIG.SEPARATOR_LINE) lines.push(CONFIG.SEPARATOR_LINE);
    lines.push(
      'Sheet: ' + (it.sheet || '-'),
      'Meeting List: #' + it.row,
      'Project: ' + (it.task || '-'),
      'Request: ' + (it.requestBy || '-'),
      'Start: ' + it.startDisp,
      'Next Due: ' + it.nextDisp,
      'Deadline: ' + it.dueDisp + (it.badges || ''),
      'By: ' + (it.by || '-'),
      'Status: ' + (it.status || '-'),
      'Change:\n' + changeLine
    );
  });
  sendLineMulti_(lines.join('\n'));
}

/* ======================= UTIL ======================= */
function _normHeader(s) {
  return String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim().toLowerCase();
}

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActive();
  const active = ss.getActiveSheet();
  const names = Array.isArray(CONFIG.SHEETS) && CONFIG.SHEETS.length
    ? CONFIG.SHEETS.slice()
    : [active.getName()]; // fallback: whatever is active

  if (names.includes(active.getName())) return active;

  const first = ss.getSheetByName(names[0]);
  return first || active;
}

function lookupFieldKeyByHeader_(editedHeaderNorm) {
  for (var key in CONFIG.HEADERS) if (_normHeader(CONFIG.HEADERS[key]) === editedHeaderNorm) return key;
  return null;
}

function chooseBy_(currentBy, requestBy, assignedTo) {
  var cb = normalizeText_(currentBy);
  var rbList = splitNames_(requestBy);
  var atList = splitNames_(assignedTo);
  if (cb && (rbList.indexOf(cb) !== -1 || atList.indexOf(cb) !== -1)) return cb; // ok as-is
  if (atList.length === 1) return atList[0]; // single assignee
  if (rbList.length === 1) return rbList[0]; // single requester
  return cb || null; // ambiguous → keep
}

/* ======================= SENDING ======================= */
function getMessageLimit_() { return CONFIG.DELIVERY_METHOD === 'OA_BROADCAST' ? 4900 : 950; }
function splitMessage_(text, limit) {
  var parts = [], remaining = text;
  while (remaining.length > limit) {
    var cut = remaining.lastIndexOf('\n', limit);
    if (cut <= 0) cut = limit;
    parts.push(remaining.slice(0, cut));
    remaining = remaining.slice(cut);
    if (remaining.startsWith('\n')) remaining = remaining.slice(1);
  }
  parts.push(remaining);
  return parts;
}

function sendLineMulti_(message) {
  var chunks = splitMessage_(message, getMessageLimit_());
  chunks.forEach(function (chunk, i) {
    var suffix = (chunks.length > 1) ? '\n\n(' + (i + 1) + '/' + chunks.length + ')' : '';
    sendLine_(chunk + suffix);
  });
}

function sendLine_(message) {
  if (CONFIG.DELIVERY_METHOD === 'OA_BROADCAST') {
    return sendLineOA_(message);
  } else {
    try { return sendLineNotify_(message); }
    catch (e) {
      if (CONFIG.OA_CHANNEL_ACCESS_TOKEN) { try { return sendLineOA_(message + '\n\n(sent via OA fallback)'); } catch (_) {} }
      throw e;
    }
  }
}

function hardResetDailyTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'runDueTodaySend' || f === 'runNightlyAlert') ScriptApp.deleteTrigger(t);
  });
  createNightAndMorningTriggersUsingConfig();  // recreates clean pair
}

function pruneOldSnapshots_(sh) {
  const p   = PropertiesService.getDocumentProperties();
  const suf = sheetKeySuffix_(sh);
  const PREFIX = 'KANBAN_SNAPSHOT_';
  const KEEP  = +CONFIG.SNAPSHOT_KEEP || 45;
  const props = p.getProperties();
  const days = Object.keys(props)
    .filter(k => k.startsWith(PREFIX) && k.endsWith(suf))
    .map(k => k.substring(PREFIX.length, k.length - suf.length))
    .filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d))
    .sort(); // yyyy-MM-dd sorts lexicographically
  if (days.length <= KEEP) return;
  const toDelete = days.slice(0, days.length - KEEP);
  toDelete.forEach(d => p.deleteProperty(PREFIX + d + suf));
  const remaining = days.slice(days.length - KEEP);
  const latest = p.getProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf);
  let prev = null;
  if (latest) {
    prev = remaining.filter(d => d < latest).pop() || null;
  } else if (remaining.length) {
    p.setProperty('KANBAN_SNAPSHOT_LATEST_DATE' + suf, remaining[remaining.length - 1]);
    prev = remaining.length > 1 ? remaining[remaining.length - 2] : null;
  }
  if (prev) p.setProperty('KANBAN_SNAPSHOT_PREV_DATE' + suf, prev);
  else p.deleteProperty('KANBAN_SNAPSHOT_PREV_DATE' + suf);
}

function sendLineNotify_(message) {
  var url = 'https://notify-api.line.me/api/notify';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: { message: message },
    headers: { Authorization: 'Bearer ' + CONFIG.LINE_NOTIFY_TOKEN },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) throw new Error('LINE Notify error: ' + res.getResponseCode() + ' ' + res.getContentText());
}

function sendLineOA_(message) {
  var headers = { Authorization: 'Bearer ' + CONFIG.OA_CHANNEL_ACCESS_TOKEN };
  var body = { messages: [{ type: 'text', text: message }] };
  var url;
  if (CONFIG.OA_GROUP_ID && /^C/.test(CONFIG.OA_GROUP_ID)) {
    // Send ONLY to this group
    url = 'https://api.line.me/v2/bot/message/push';
    body.to = CONFIG.OA_GROUP_ID;
  } else {
    // Fallback: broadcast (only if you intentionally leave OA_GROUP_ID empty)
    url = 'https://api.line.me/v2/bot/message/broadcast';
  }
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    headers: headers,
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('LINE OA error: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
}
/* ======================= TEST ======================= */
function testSend() {
  sendLineMulti_(CONFIG.TITLE + '\n\nThis is a test message at ' +
    Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm'));
}
