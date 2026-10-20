/**
 * Ready Set Go - live scoring system.
 * Bound to a Google Sheet. Deploy as a Web App (Execute as: Me, Access: Anyone with the link).
 */

var BOOSTS = [10, 20, 35, 50];
var UNDO_WINDOW_MS = 60 * 1000;
// Max uses of each boost per professional at the planned event size. Scaled to the number of
// active students (rounded up), so 22 students gives exactly these numbers.
var BASE_LIMITS = { 10: 8, 20: 7, 35: 6, 50: 5 };
var BASE_STUDENTS = 22;

var TAB = {
  scores: 'Scores',
  students: 'Students',
  pros: 'Professionals',
  board: 'Leaderboard',
  monitor: 'Live Monitor'
};

// Scores columns (1-indexed)
var C = { ts: 1, entry: 2, pro: 3, student: 4, boost: 5, note: 6, voided: 7 };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ready Set Go')
    .addItem('Set up sheet tabs', 'setup')
    .addItem('Rebuild leaderboard', 'rebuildViews')
    .addItem('Add student', 'addStudentPrompt')
    .addItem('Add professional', 'addProfessionalPrompt')
    .addItem('Remove student', 'removeStudentPrompt')
    .addItem('Remove professional', 'removeProfessionalPrompt')
    .addItem('Set admin passcode', 'setAdminPasscode')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var defs = [
    [TAB.scores, ['Timestamp', 'Entry ID', 'Professional ID', 'Student ID', 'Boost', 'Note', 'Voided']],
    [TAB.students, ['ID', 'Name', 'Major', 'Checked In', 'Active']],
    [TAB.pros, ['ID', 'Name', 'Company', 'Active']],
    [TAB.board, ['Rank', 'Student ID', 'Student', 'Total', 'Interviews', 'Average', 'Highest Boost']],
    [TAB.monitor, ['Professional', 'Company', 'Entries', 'Last Entry']]
  ];
  defs.forEach(function (d) {
    var sh = ss.getSheetByName(d[0]) || ss.insertSheet(d[0]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, d[1].length).setValues([d[1]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // Older sheets: make sure the Active column header exists (blank = active, N = removed)
  ss.getSheetByName(TAB.students).getRange(1, 5).setValue('Active').setFontWeight('bold');
  ss.getSheetByName(TAB.pros).getRange(1, 4).setValue('Active').setFontWeight('bold');
  // IDs are plain text so "01" stays "01"
  ss.getSheetByName(TAB.students).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(TAB.pros).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(TAB.scores).getRange('C:D').setNumberFormat('@');
  ss.getSheetByName(TAB.board).getRange('B:B').setNumberFormat('@');
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function doGet(e) {
  var admin = e && e.parameter && e.parameter.page === 'admin';
  return HtmlService.createHtmlOutputFromFile(admin ? 'admin' : 'index')
    .setTitle(admin ? 'Ready Set Go Admin' : 'Ready Set Go Scoring')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- JSON API for the GitHub Pages front end ----------
// POST body {"fn":"submitScore","args":[...]} -> {"result":...} or {"error":"..."}.
// Requests arrive without the visitor's Google sign-in, so school accounts and
// multiple signed-in accounts can't break them. Only these functions are callable.
var API = {
  getProfessionals: getProfessionals, lookupStudent: lookupStudent, getHistory: getHistory,
  getRemaining: getRemaining, submitScore: submitScore, undoLast: undoLast,
  adminLogin: adminLogin, adminList: adminList, adminAddStudent: adminAddStudent,
  adminImportStudents: adminImportStudents, adminUpdateStudent: adminUpdateStudent,
  adminSetStudentFlag: adminSetStudentFlag, adminAddProfessional: adminAddProfessional,
  adminUpdateProfessional: adminUpdateProfessional, adminSetProfessionalActive: adminSetProfessionalActive,
  adminOverview: adminOverview, adminSetVoided: adminSetVoided,
  adminDeleteStudent: adminDeleteStudent, adminDeleteProfessional: adminDeleteProfessional
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (!Object.prototype.hasOwnProperty.call(API, req.fn)) throw new Error('Unknown function.');
    out = { result: API[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- helpers ----------

function normId_(v) {
  var s = String(v == null ? '' : v).trim().toUpperCase();
  if (/^\d+$/.test(s)) return String(parseInt(s, 10));
  return s;
}

function isActive_(v) { return String(v == null ? '' : v).trim().toUpperCase() !== 'N'; }

function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function rows_(name, width) {
  var sh = sheet_(name);
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getValues();
}

// ---------- client API ----------

function getProfessionals() {
  return rows_(TAB.pros, 4)
    .filter(function (r) { return r[0] !== '' && r[1] !== '' && isActive_(r[3]); })
    .map(function (r) { return { id: normId_(r[0]), name: String(r[1]), company: String(r[2] || '') }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

function lookupStudent(id) {
  var want = normId_(id);
  if (!want) return null;
  var rows = rows_(TAB.students, 5);
  for (var i = 0; i < rows.length; i++) {
    if (normId_(rows[i][0]) === want && isActive_(rows[i][4])) return { id: want, name: String(rows[i][1]) };
  }
  return null;
}

/** Returns this professional's entries so far (for the on-page history). */
function getHistory(proId) {
  var pro = normId_(proId);
  var names = {};
  rows_(TAB.students, 2).forEach(function (r) { names[normId_(r[0])] = String(r[1]); });
  return rows_(TAB.scores, 7)
    .filter(function (r) { return normId_(r[C.pro - 1]) === pro && r[C.voided - 1] !== 'Y'; })
    .map(function (r) {
      var sid = normId_(r[C.student - 1]);
      return { student: sid, name: names[sid] || '', boost: r[C.boost - 1] };
    })
    .reverse();
}

/** Max uses of each boost per professional, scaled to the active student count. */
function getLimits_() {
  var n = rows_(TAB.students, 5).filter(function (r) { return r[0] !== '' && isActive_(r[4]); }).length;
  var limits = {};
  BOOSTS.forEach(function (b) { limits[b] = Math.ceil(BASE_LIMITS[b] * n / BASE_STUDENTS); });
  return limits;
}

/** How many of each boost this professional has used (voided entries don't count). */
function usedBoosts_(pro, data) {
  var used = { 10: 0, 20: 0, 35: 0, 50: 0 };
  data.forEach(function (r) {
    if (r[C.voided - 1] !== 'Y' && normId_(r[C.pro - 1]) === pro) used[Number(r[C.boost - 1])]++;
  });
  return used;
}

function remaining_(pro, data) {
  var used = usedBoosts_(pro, data), limits = getLimits_(), out = {};
  BOOSTS.forEach(function (b) { out[b] = Math.max(0, limits[b] - used[b]); });
  return out;
}

/** Remaining uses of each boost for this professional, e.g. {10:8, 20:5, 35:6, 50:0}. */
function getRemaining(proId) {
  return remaining_(normId_(proId), rows_(TAB.scores, 7));
}

/**
 * Record a score. If this professional already scored this student and
 * replace is falsy, returns {needsConfirm:true, existing:<boost>} instead.
 */
function submitScore(proId, studentId, boost, note, replace) {
  boost = Number(boost);
  if (BOOSTS.indexOf(boost) === -1) return { ok: false, error: 'Invalid boost value.' };
  var pro = normId_(proId);
  var student = lookupStudent(studentId);
  if (!student) return { ok: false, error: 'Unknown student number.' };
  if (!getProfessionals().some(function (p) { return p.id === pro; }))
    return { ok: false, error: 'Unknown professional.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(TAB.scores);
    var data = rows_(TAB.scores, 7);
    var dupRow = -1;
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (r[C.voided - 1] !== 'Y' && normId_(r[C.pro - 1]) === pro && normId_(r[C.student - 1]) === student.id) {
        dupRow = i + 2;
        var existing = r[C.boost - 1];
      }
    }
    // Boost limit. A score being replaced frees its slot.
    var used = usedBoosts_(pro, data);
    if (dupRow > 0) used[Number(existing)]--;
    var limits = getLimits_();
    if (used[boost] >= limits[boost])
      return { ok: false, error: 'No +' + boost + ' scores left (limit ' + limits[boost] + ').', remaining: remaining_(pro, data) };
    if (dupRow > 0) {
      if (!replace) return { ok: false, needsConfirm: true, existing: existing, student: student.name };
      sh.getRange(dupRow, C.voided).setValue('Y');
    }
    sh.appendRow([new Date(), Utilities.getUuid(), "'" + pro, "'" + student.id, boost, note || '', '']);
    rebuildViews();
    return { ok: true, student: student.name, boost: boost, replaced: dupRow > 0, remaining: getRemaining(pro) };
  } finally {
    lock.releaseLock();
  }
}

/** Voids the professional's most recent entry if it is under 60 seconds old. */
function undoLast(proId) {
  var pro = normId_(proId);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(TAB.scores);
    var data = rows_(TAB.scores, 7);
    for (var i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      if (r[C.voided - 1] === 'Y' || normId_(r[C.pro - 1]) !== pro) continue;
      if (Date.now() - new Date(r[C.ts - 1]).getTime() > UNDO_WINDOW_MS)
        return { ok: false, error: 'Too late to undo. Ask a committee member.' };
      sh.getRange(i + 2, C.voided).setValue('Y');
      rebuildViews();
      return { ok: true };
    }
    return { ok: false, error: 'Nothing to undo.' };
  } finally {
    lock.releaseLock();
  }
}

// ---------- leaderboard + monitor ----------

function rebuildViews() {
  var scores = rows_(TAB.scores, 7).filter(function (r) { return r[C.voided - 1] !== 'Y' && r[C.student - 1] !== ''; });
  var students = rows_(TAB.students, 5).filter(function (r) { return r[0] !== '' && isActive_(r[4]); });

  var stats = {};
  students.forEach(function (r) {
    stats[normId_(r[0])] = { id: normId_(r[0]), name: String(r[1]), total: 0, n: 0, max: 0 };
  });
  var proStats = {};
  scores.forEach(function (r) {
    var s = stats[normId_(r[C.student - 1])];
    var b = Number(r[C.boost - 1]);
    if (s) { s.total += b; s.n += 1; s.max = Math.max(s.max, b); }
    var p = normId_(r[C.pro - 1]);
    var ts = new Date(r[C.ts - 1]);
    if (!proStats[p]) proStats[p] = { n: 0, last: ts };
    proStats[p].n += 1;
    if (ts > proStats[p].last) proStats[p].last = ts;
  });

  // Rank: total desc, then highest single boost desc, then average desc
  var list = Object.keys(stats).map(function (k) { return stats[k]; });
  list.sort(function (a, b) {
    return (b.total - a.total) || (b.max - a.max) ||
      ((b.n ? b.total / b.n : 0) - (a.n ? a.total / a.n : 0));
  });
  var out = list.map(function (s, i) {
    return [i + 1, s.id, s.name, s.total, s.n, s.n ? Math.round(s.total / s.n * 10) / 10 : 0, s.max];
  });
  var board = sheet_(TAB.board);
  board.getRange(2, 1, Math.max(board.getMaxRows() - 1, 1), 7).clearContent().setBackground(null);
  if (out.length) {
    board.getRange(2, 1, out.length, 7).setValues(out);
    var colors = ['#ffd700', '#c0c0c0', '#cd7f32'];
    for (var i = 0; i < Math.min(3, out.length); i++) board.getRange(2 + i, 1, 1, 7).setBackground(colors[i]);
  }

  var pros = rows_(TAB.pros, 4).filter(function (r) { return r[0] !== '' && isActive_(r[3]); });
  var mon = pros.map(function (r) {
    var ps = proStats[normId_(r[0])];
    return [r[1], r[2], ps ? ps.n : 0, ps ? ps.last : ''];
  }).sort(function (a, b) { return a[2] - b[2]; });
  var m = sheet_(TAB.monitor);
  m.getRange(2, 1, Math.max(m.getMaxRows() - 1, 1), 4).clearContent();
  if (mon.length) m.getRange(2, 1, mon.length, 4).setValues(mon);
}

// ---------- committee: add / remove people ----------
// Removal is a soft delete: Active = N. Past scores stay in the Scores tab.

function findRow_(tab, width, id) {
  var want = normId_(id);
  var rows = rows_(tab, width);
  for (var i = 0; i < rows.length; i++) if (normId_(rows[i][0]) === want) return i + 2;
  return -1;
}

function addStudentPrompt() {
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt('Student sticker number:').getResponseText();
  var name = ui.prompt('Student name:').getResponseText();
  if (!id || !name) return;
  var sh = sheet_(TAB.students);
  var row = findRow_(TAB.students, 5, id);
  if (row > 0) {
    if (isActive_(sh.getRange(row, 5).getValue())) { ui.alert('That number is already taken.'); return; }
    sh.getRange(row, 2).setValue(name);       // reactivate a removed student
    sh.getRange(row, 5).setValue('');
  } else {
    sh.appendRow(["'" + normId_(id), name, '', 'Y', '']);
  }
  rebuildViews();
}

function addProfessionalPrompt() {
  var ui = SpreadsheetApp.getUi();
  var name = ui.prompt('Professional name:').getResponseText();
  var company = ui.prompt('Company:').getResponseText();
  if (!name) return;
  // Next ID = highest existing P-number + 1, so deleted/removed rows never cause a collision
  var max = 0;
  rows_(TAB.pros, 4).forEach(function (r) {
    var m = /^P(\d+)$/.exec(normId_(r[0]));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  sheet_(TAB.pros).appendRow(["'P" + (max + 1), name, company, '']);
  rebuildViews();
}

function removeStudentPrompt() {
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt('Sticker number of the student to remove:').getResponseText();
  var row = findRow_(TAB.students, 5, id);
  if (row < 0) { ui.alert('No student with that number.'); return; }
  var sh = sheet_(TAB.students);
  var name = sh.getRange(row, 2).getValue();
  if (ui.alert('Remove ' + name + '? They leave the leaderboard but their scores stay in the Scores tab.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.getRange(row, 5).setValue('N');
  rebuildViews();
}

function removeProfessionalPrompt() {
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt('ID of the professional to remove (e.g. P3):').getResponseText();
  var row = findRow_(TAB.pros, 4, id);
  if (row < 0) { ui.alert('No professional with that ID.'); return; }
  var sh = sheet_(TAB.pros);
  var name = sh.getRange(row, 2).getValue();
  if (ui.alert('Remove ' + name + '? They can no longer submit, but their past scores still count.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.getRange(row, 4).setValue('N');
  rebuildViews();
}

// ---------- admin page (?page=admin) ----------
// Every admin call takes the passcode stored in Script Properties (menu: Set admin passcode).

function setAdminPasscode() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('New admin passcode (used on the ?page=admin page):');
  var code = r.getResponseText().trim();
  if (r.getSelectedButton() !== ui.Button.OK || !code) return;
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSCODE', code);
  ui.alert('Admin passcode saved.');
}

function checkAdmin_(pass) {
  var want = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
  if (!want) throw new Error('No admin passcode set. In the sheet: Ready Set Go > Set admin passcode.');
  if (String(pass) !== want) throw new Error('Wrong passcode.');
}

function adminLogin(pass) {
  checkAdmin_(pass);
  return adminList(pass);
}

function adminList(pass) {
  checkAdmin_(pass);
  var students = rows_(TAB.students, 5)
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return { id: normId_(r[0]), name: String(r[1]), major: String(r[2] || ''),
               checkedIn: String(r[3]).toUpperCase() === 'Y', active: isActive_(r[4]) };
    });
  var pros = rows_(TAB.pros, 4)
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return { id: normId_(r[0]), name: String(r[1]), company: String(r[2] || ''), active: isActive_(r[3]) };
    });
  var counts = { students: {}, pros: {} };
  rows_(TAB.scores, 7).forEach(function (r) {
    if (r[C.voided - 1] === 'Y') return;
    var s = normId_(r[C.student - 1]), p = normId_(r[C.pro - 1]);
    counts.students[s] = (counts.students[s] || 0) + 1;
    counts.pros[p] = (counts.pros[p] || 0) + 1;
  });
  students.forEach(function (s) { s.scores = counts.students[s.id] || 0; });
  pros.forEach(function (p) { p.scores = counts.pros[p.id] || 0; });
  students.sort(function (a, b) { return a.id.localeCompare(b.id, undefined, { numeric: true }); });
  pros.sort(function (a, b) { return a.id.localeCompare(b.id, undefined, { numeric: true }); });
  return { students: students, pros: pros };
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** Adds a student, or reactivates a removed one with the same number. */
function addStudentRow_(id, name, major) {
  id = normId_(id); name = String(name || '').trim();
  if (!id || !name) return 'Number and name are required.';
  var sh = sheet_(TAB.students);
  var row = findRow_(TAB.students, 5, id);
  if (row > 0) {
    if (isActive_(sh.getRange(row, 5).getValue())) return 'Number ' + id + ' is already taken.';
    sh.getRange(row, 2, 1, 4).setValues([[name, major || '', 'Y', '']]);   // reactivate
  } else {
    sh.appendRow(["'" + id, name, major || '', 'Y', '']);
  }
  return null;
}

function adminAddStudent(pass, id, name, major) {
  checkAdmin_(pass);
  return withLock_(function () {
    var err = addStudentRow_(id, name, major);
    if (err) return { ok: false, error: err };
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

/** text: one student per line, "number, name, major" (comma or tab separated; major optional). */
function adminImportStudents(pass, text) {
  checkAdmin_(pass);
  return withLock_(function () {
    var added = 0, errors = [];
    String(text || '').split(/\r?\n/).forEach(function (line, i) {
      if (!line.trim()) return;
      var p = line.split(/\t|,/).map(function (s) { return s.trim(); });
      var err = addStudentRow_(p[0], p[1], p[2]);
      if (err) errors.push('Line ' + (i + 1) + ': ' + err); else added++;
    });
    rebuildViews();
    return { ok: true, added: added, errors: errors, data: adminList(pass) };
  });
}

function adminUpdateStudent(pass, id, name, major) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.students, 5, id);
    if (row < 0) return { ok: false, error: 'No student with that number.' };
    if (!String(name).trim()) return { ok: false, error: 'Name is required.' };
    sheet_(TAB.students).getRange(row, 2, 1, 2).setValues([[String(name).trim(), major || '']]);
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

function adminSetStudentFlag(pass, id, field, value) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.students, 5, id);
    if (row < 0) return { ok: false, error: 'No student with that number.' };
    var sh = sheet_(TAB.students);
    if (field === 'checkedIn') sh.getRange(row, 4).setValue(value ? 'Y' : '');
    else if (field === 'active') sh.getRange(row, 5).setValue(value ? '' : 'N');
    else return { ok: false, error: 'Bad field.' };
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

function adminAddProfessional(pass, name, company) {
  checkAdmin_(pass);
  return withLock_(function () {
    name = String(name || '').trim();
    if (!name) return { ok: false, error: 'Name is required.' };
    var max = 0;
    rows_(TAB.pros, 4).forEach(function (r) {
      var m = /^P(\d+)$/.exec(normId_(r[0]));
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    sheet_(TAB.pros).appendRow(["'P" + (max + 1), name, String(company || '').trim(), '']);
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

function adminUpdateProfessional(pass, id, name, company) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.pros, 4, id);
    if (row < 0) return { ok: false, error: 'No professional with that ID.' };
    if (!String(name).trim()) return { ok: false, error: 'Name is required.' };
    sheet_(TAB.pros).getRange(row, 2, 1, 2).setValues([[String(name).trim(), String(company || '').trim()]]);
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

function adminSetProfessionalActive(pass, id, active) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.pros, 4, id);
    if (row < 0) return { ok: false, error: 'No professional with that ID.' };
    sheet_(TAB.pros).getRange(row, 4).setValue(active ? '' : 'N');
    rebuildViews();
    return { ok: true, data: adminList(pass) };
  });
}

// ---------- admin: event overview (leaderboard, live monitor, scores) ----------

/** Everything the admin Event tabs show. refresh=true recalculates the Leaderboard/Live Monitor tabs first. */
function adminOverview(pass, refresh) {
  checkAdmin_(pass);
  if (refresh) withLock_(rebuildViews);
  var ms = function (v) { return v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() : 0); };
  var sNames = {}, pNames = {};
  rows_(TAB.students, 2).forEach(function (r) { sNames[normId_(r[0])] = String(r[1]); });
  rows_(TAB.pros, 3).forEach(function (r) { pNames[normId_(r[0])] = { name: String(r[1]), company: String(r[2] || '') }; });
  var board = rows_(TAB.board, 7).filter(function (r) { return r[1] !== ''; }).map(function (r) {
    return { rank: r[0], id: normId_(r[1]), name: String(r[2]), total: r[3], n: r[4], avg: r[5], max: r[6] };
  });
  var monitor = rows_(TAB.monitor, 4).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { name: String(r[0]), company: String(r[1] || ''), n: r[2], last: ms(r[3]) };
  });
  var scores = rows_(TAB.scores, 7).filter(function (r) { return r[C.entry - 1] !== ''; }).map(function (r) {
    var p = normId_(r[C.pro - 1]), s = normId_(r[C.student - 1]);
    return { entry: String(r[C.entry - 1]), ts: ms(r[C.ts - 1]), pro: p, proName: pNames[p] ? pNames[p].name : p,
             student: s, studentName: sNames[s] || '', boost: r[C.boost - 1], note: String(r[C.note - 1] || ''),
             voided: r[C.voided - 1] === 'Y' };
  }).reverse();
  return { board: board, monitor: monitor, scores: scores };
}

function adminSetVoided(pass, entryId, voided) {
  checkAdmin_(pass);
  return withLock_(function () {
    var data = rows_(TAB.scores, 7);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][C.entry - 1]) === String(entryId)) {
        sheet_(TAB.scores).getRange(i + 2, C.voided).setValue(voided ? 'Y' : '');
        rebuildViews();
        return { ok: true, overview: adminOverview(pass, false) };
      }
    }
    return { ok: false, error: 'Entry not found.' };
  });
}

// ---------- admin: one-step delete ----------
// Removes the person (Active = N) AND voids all their scores, so they vanish from the
// leaderboard, monitor and counts. Rows stay in the sheet; re-adding the same number/ID reactivates it.

function voidScoresFor_(col, id) {
  var want = normId_(id), sh = sheet_(TAB.scores), data = rows_(TAB.scores, 7), n = 0;
  data.forEach(function (r, i) {
    if (r[C.voided - 1] !== 'Y' && normId_(r[col - 1]) === want) { sh.getRange(i + 2, C.voided).setValue('Y'); n++; }
  });
  return n;
}

function adminDeleteStudent(pass, id) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.students, 5, id);
    if (row < 0) return { ok: false, error: 'No student with that number.' };
    var n = voidScoresFor_(C.student, id);
    sheet_(TAB.students).getRange(row, 5).setValue('N');
    rebuildViews();
    return { ok: true, voided: n, data: adminList(pass) };
  });
}

function adminDeleteProfessional(pass, id) {
  checkAdmin_(pass);
  return withLock_(function () {
    var row = findRow_(TAB.pros, 4, id);
    if (row < 0) return { ok: false, error: 'No professional with that ID.' };
    var n = voidScoresFor_(C.pro, id);
    sheet_(TAB.pros).getRange(row, 4).setValue('N');
    rebuildViews();
    return { ok: true, voided: n, data: adminList(pass) };
  });
}
