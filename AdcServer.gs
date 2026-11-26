/**
 * ADC cabinets — automated dispensing cabinet stock, refill, monthly count and zeroing (NUPCO ABC-VEN).
 * Pharmacy Tracker by WaadAlwagdani
 *
 * Sheets (created by adcInstall):
 *   ADC Items         one row per medication: master data, par levels and the opening balance of the open period
 *   ADC Transactions  append-only movement log (every withdrawal, override and refill) — the audit trail
 *   ADC Counts        monthly physical counts; APPROVED rows close a month and set the next opening balance
 * Live stock on hand = opening balance of the open period + received − withdrawn in that period.
 * Formulas are never written to these tabs: every figure is recomputed by the server, and the downloadable
 * Excel copy is generated with its formulas (System Balance, Variance, New Opening, Fill %, Qty to Refill).
 * No patient data is recorded. Unit cost is optional: values in SAR are shown only for items that have a cost.
 */
const ADC = {
  ITEMS: 'ADC Items', TX: 'ADC Transactions', COUNTS: 'ADC Counts', PROP: 'adc.period', DRAFT: 'adc.countDraft', ASSIGN: 'adc.assign',
  ITEM_HEAD: ['Code', 'NUPCO Code', 'Drug & Strength', 'ADC Cluster', 'Primary Ward', 'ABC-VEN', 'Unit Cost (SAR, optional)', 'Par Min', 'Par Max',
    'Opening Balance', 'Baseline Avg Daily Use', 'Alert (HIGH / CONTROLLED)'],
  TX_HEAD: ['Tx ID', 'Timestamp', 'Month', 'Date & Shift', 'Type', 'Code', 'Drug & Strength', 'ADC Cluster', 'Destination Ward',
    'User', 'Role', 'Qty', 'Value (SAR)', 'Override reason', 'Note'],
  COUNT_HEAD: ['Month', 'Code', 'Drug & Strength', 'Opening', 'Received', 'Withdrawn', 'System Balance', 'Physical Count',
    'Variance', 'Variance (SAR)', 'Root Cause', 'Counted by', 'Approved at', 'Status'],
  CLUSTERS: [
    ['ADC-01 (ICU - FMW - FSW - Burn)', ['ICU', 'FMW', 'FSW', 'Burn']],
    ['ADC-02 (LTCU - Gyne - HDU - MSW)', ['LTCU', 'Gyne', 'HDU', 'MSW']],
    ['ADC-03 (Isolation - CCU - MMW - MS3)', ['Isolation', 'CCU', 'MMW', 'MS3']],
    ['ADC-04 (NICU - DR - AntA/B - Post-Op)', ['NICU', 'DR', 'AntA/B', 'Post-Op']]],
  OUT: ['SCANNED', 'OVERRIDE'],
  REASONS: ['Code / Rapid response', 'Patient deterioration', 'Barcode unreadable / damaged', 'Scanner down', 'Physician verbal order (STAT)', 'Other'],
  CACHE: 'adc.view.v1'
};

/* ------------------------------ install (run once from the editor) ------------------------------ */
function adcInstall() {
  assertEditor_();
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(ADC.ITEMS)) throw new Error('ADC is already installed — the tabs exist. Delete them first to reinstall.');
  const S = ADC_SEED;
  const it = ss.insertSheet(ADC.ITEMS);
  it.getRange(1, 1, 1, ADC.ITEM_HEAD.length).setValues([ADC.ITEM_HEAD]).setFontWeight('bold').setBackground('#f1efeb');
  const rows = S.items.map(function (r) { return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]]; });
  it.getRange(2, 1, rows.length, ADC.ITEM_HEAD.length).setValues(rows); it.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  it.setFrozenRows(1);
  const tx = ss.insertSheet(ADC.TX);
  tx.getRange(1, 1, 1, ADC.TX_HEAD.length).setValues([ADC.TX_HEAD]).setFontWeight('bold').setBackground('#f1efeb');
  const byCode = {}; S.items.forEach(function (r) { byCode[r[0]] = r; });
  const tr = S.tx.map(function (t) {
    const i = byCode[t[4]] || [];
    const d = String(t[2]).slice(0, 10);
    return [t[0], d, t[1], t[2], t[3], t[4], i[2] || '', i[3] || '', t[5] || '', t[6], t[7], t[8], t[9], '', t[10]];
  });
  if (tr.length) tx.getRange(2, 1, tr.length, ADC.TX_HEAD.length).setValues(tr);
  tx.setFrozenRows(1);
  const ct = ss.insertSheet(ADC.COUNTS);
  ct.getRange(1, 1, 1, ADC.COUNT_HEAD.length).setValues([ADC.COUNT_HEAD]).setFontWeight('bold').setBackground('#f1efeb');
  const cr = S.counts.map(function (c) { const i = byCode[c[1]] || [];
    return [c[0], c[1], i[2] || '', c[2], c[3], c[4], c[5], c[6], c[7], c[8], c[9], 'Imported', '', 'APPROVED']; });
  if (cr.length) ct.getRange(2, 1, cr.length, ADC.COUNT_HEAD.length).setValues(cr);
  ct.setFrozenRows(1);
  PropertiesService.getScriptProperties().setProperty(ADC.PROP, S.period);
  return 'ADC installed: ' + rows.length + ' items, ' + tr.length + ' transactions, ' + cr.length + ' count rows. Open period ' + S.period;
}

/* ------------------------------ helpers ------------------------------ */
function adcSheet_(n) { const s = SpreadsheetApp.getActive().getSheetByName(n); if (!s) throw new Error('ADC_NOT_INSTALLED'); return s; }
function adcPeriod_() { return PropertiesService.getScriptProperties().getProperty(ADC.PROP) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'); }
function adcPerms_(u) {
  const own = isOwner_(u), st = u.stations || [];
  return { withdraw: own || st.indexOf('adc') > -1, refill: own || st.indexOf('adcrefill') > -1, count: own || st.indexOf('adccount') > -1,
    sup: own || u.sysRole === 'PHARMACY_SUPERVISOR' || u.sysRole === 'DEPT_SUPERVISOR' || u.sysRole === 'QUALITY' };
}
function adcUser_(token, need) {
  // 'sup' is a read (summary, export, count history) — only actions that change data need a writable session
  const u = requireSession_(token, 'adc', need && need !== 'sup' ? { write: true } : undefined), p = adcPerms_(u);
  if (need && !p[need]) throw new Error('FORBIDDEN');
  return { u: u, p: p };
}

/* ------------------------------ cabinet assignment (set by the supervisor) ------------------------------ */
function adcAssign_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(ADC.ASSIGN) || '{}'); } catch (e) { return {}; } }
/** null = every cabinet (owner / supervisors); otherwise the cabinets assigned to this account. */
function adcScope_(a) {
  if (a.p.sup) return null;
  const m = adcAssign_();
  return ADC.CLUSTERS.map(function (c) { return c[0]; }).filter(function (c) { return (m[c] || []).indexOf(a.u.id) > -1; });
}
function adcInScope_(a, adc) { const s = adcScope_(a); if (s && s.indexOf(adc) === -1) throw new Error('ADC_NOT_ASSIGNED'); }
/** Supervisor: who is responsible for which cabinet. map = { 'ADC-01 (…)': [userId, …], … } */
function adcAssignSave(token, map) {
  const a = adcUser_(token, 'sup'); requireSession_(token, 'adc', { write: true });
  if (!map || typeof map !== 'object') throw new Error('INVALID_INPUT');
  const ids = {}; allUsers_().forEach(function (u) { ids[u.id] = 1; });
  const out = {};
  ADC.CLUSTERS.forEach(function (c) {
    const v = Array.isArray(map[c[0]]) ? map[c[0]] : [];
    out[c[0]] = v.filter(function (id, i) { return ids[id] && v.indexOf(id) === i; }).slice(0, 30);
  });
  PropertiesService.getScriptProperties().setProperty(ADC.ASSIGN, JSON.stringify(out));
  return { ok: true, assign: out, by: a.u.displayName };
}
function adcStaff_() {
  return allUsers_().filter(function (u) { return u.status === 'active' && (u.stations || []).some(function (s) { return ['adc', 'adcrefill', 'adccount'].indexOf(s) > -1; }); })
    .map(function (u) { return { id: u.id, name: u.displayName || u.username, stations: (u.stations || []).filter(function (s) { return s.indexOf('adc') === 0; }) }; });
}

function adcItems_() {
  const sh = adcSheet_(ADC.ITEMS), n = sh.getLastRow() - 1; if (n < 1) return [];
  // columns B and C are read by their header: the sheet may hold «Drug & Strength» in B and «NUPCO Code» in C, or the other way round
  const head = sh.getRange(1, 1, 1, 3).getValues()[0], nameInB = /drug|name/i.test(String(head[1])) && !/nupco/i.test(String(head[1]));
  return sh.getRange(2, 1, n, ADC.ITEM_HEAD.length).getValues().map(function (r, i) {
    return { row: i + 2, code: String(r[0]), nupco: String(nameInB ? r[2] : r[1]), name: String(nameInB ? r[1] : r[2]), adc: String(r[3]), ward: String(r[4]), abc: String(r[5]),
      cost: r[6] === '' || r[6] === null ? null : (Number(r[6]) || null), parMin: Number(r[7]) || 0, parMax: Number(r[8]) || 0, opening: Number(r[9]) || 0, base: Number(r[10]) || 0, alert: String(r[11] || '') };
  }).filter(function (x) { return x.code; });
}
function adcTx_() {
  const sh = adcSheet_(ADC.TX), n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, ADC.TX_HEAD.length).getValues();
}
function adcNum_(v) { return Number(v) || 0; }
/** Money only when the unit cost is known — otherwise blank, never a guess. */
function adcVal_(q, cost) { return cost ? Math.round(q * cost * 100) / 100 : ''; }
function adcStatus_(soh, min) { return soh <= Math.max(1, Math.round(min * 0.45)) ? 'STAT' : soh <= min ? 'REORDER' : 'OK'; }
/** Builds the live view: stock on hand per item from the opening balance + movements of the open period. */
function adcView_() {
  const period = adcPeriod_(), items = adcItems_(), tx = adcTx_(), mov = {}, use30 = {};
  const tz = Session.getScriptTimeZone(), since = Utilities.formatDate(new Date(Date.now() - 30 * 864e5), tz, 'yyyy-MM-dd');
  let maxId = 1000;
  tx.forEach(function (r) {
    const num = Number(String(r[0]).replace(/\D/g, ''));
    if (num > maxId) maxId = num;
    const code = String(r[5]), type = String(r[4]), q = adcNum_(r[11]);
    if (String(r[2]) === period) {
      const m = mov[code] || (mov[code] = { rec: 0, out: 0 });
      if (type === 'RECEIVED' || type === 'RETURN') m.rec += q; else if (ADC.OUT.indexOf(type) > -1) m.out += q;
    }
    const d = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]).slice(0, 10);
    if (ADC.OUT.indexOf(type) > -1 && d >= since) use30[code] = (use30[code] || 0) + q;
  });
  return { period: period, maxId: maxId, items: items.map(function (x) {
    const m = mov[x.code] || { rec: 0, out: 0 }, soh = x.opening + m.rec - m.out;
    const avg = use30[x.code] ? Math.round(use30[x.code] / 30 * 10) / 10 : x.base, st = adcStatus_(soh, x.parMin);
    return { code: x.code, nupco: x.nupco, name: x.name, adc: x.adc, ward: x.ward, abc: x.abc, cost: x.cost, parMin: x.parMin, parMax: x.parMax,
      opening: x.opening, rec: m.rec, out: m.out, soh: soh, fill: x.parMax ? Math.round(soh / x.parMax * 100) : 0, avg: avg,
      days: avg > 0 ? Math.round(soh / avg * 10) / 10 : null, status: st, refill: soh <= x.parMin ? Math.max(0, x.parMax - soh) : 0, alert: x.alert };
  }) };
}
function adcShift_(d) {
  const h = d.getHours(), tz = Session.getScriptTimeZone(), day = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return day + ' | ' + (h >= 7 && h < 15 ? 'صباحية (07:00-15:00)' : h >= 15 && h < 23 ? 'مسائية (15:00-23:00)' : 'ليلية (23:00-07:00)');
}
function adcRole_(u) { return { OWNER: 'Owner', PHARMACY_SUPERVISOR: 'Pharmacy supervisor', DEPT_SUPERVISOR: 'Department supervisor', PHARMACIST: 'Pharmacist', NURSE: 'Nurse', QUALITY: 'Quality' }[u.sysRole] || u.sysRole; }
function adcAppend_(rows) { if (rows.length) { const sh = adcSheet_(ADC.TX); sh.getRange(sh.getLastRow() + 1, 1, rows.length, ADC.TX_HEAD.length).setValues(rows); } }
function adcWards_() { const w = []; ADC.CLUSTERS.forEach(function (c) { c[1].forEach(function (x) { w.push({ ward: x, adc: c[0] }); }); }); return w; }

/* ------------------------------ read ------------------------------ */
function adcGet(token) {
  const a = adcUser_(token), v = adcView_(), sc = adcScope_(a), all = ADC.CLUSTERS.map(function (c) { return c[0]; });
  const ok = function (c) { return !sc || sc.indexOf(c) > -1; };
  // strength and dosage form come from the item lists (the ADC code is the item's serial number)
  const dm = {}; try { getData().drugs.forEach(function (d) { dm[d.no] = d; }); } catch (e) {}
  const items = v.items.filter(function (x) { return ok(x.adc); }).map(function (x) { const d = dm[x.code]; x.strength = d ? d.strength : ''; x.form = d ? d.form : ''; return x; });
  return { period: v.period, perms: a.p, items: items, priced: items.some(function (x) { return x.cost; }), wards: adcWards_().filter(function (w) { return ok(w.adc); }),
    clusters: all.filter(ok), scoped: !!sc, total: v.items.length,
    reasons: ADC.REASONS, draft: (a.p.count || a.p.sup) ? adcDraft_() : null, me: a.u.displayName,
    assign: a.p.sup ? adcAssign_() : null, staff: a.p.sup ? adcStaff_() : null };
}

/* ------------------------------ nursing: withdraw ------------------------------ */
function adcWithdraw(token, code, qty, ward, scanned, reason) {
  const a = adcUser_(token, 'withdraw'), q = Number(qty);
  if (!(q === Math.floor(q) && q > 0 && q <= 500)) throw new Error('INVALID_INPUT');
  const w = adcWards_().filter(function (x) { return x.ward === ward; })[0]; if (!w) throw new Error('INVALID_INPUT');
  if (!scanned && ADC.REASONS.indexOf(reason) === -1) throw new Error('ADC_REASON_REQUIRED');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = adcView_(), it = v.items.filter(function (x) { return x.code === code; })[0];
    if (!it) throw new Error('NOT_FOUND');
    if (w.adc !== it.adc) throw new Error('INVALID_INPUT');
    adcInScope_(a, it.adc);
    if (q > it.soh) throw new Error('ADC_INSUFFICIENT:' + it.soh);   // never below the stock physically in the drawer
    const now = new Date(), id = 'TX-' + (v.maxId + 1);
    adcAppend_([[id, Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), v.period, adcShift_(now), scanned ? 'SCANNED' : 'OVERRIDE',
      it.code, it.name, it.adc, ward, a.u.displayName + ' (@' + a.u.username + ')', adcRole_(a.u), q, adcVal_(q, it.cost),
      scanned ? '' : reason, scanned ? 'Scanned (BCMA)' : 'STAT override — not scanned']]);
    SpreadsheetApp.flush();
    return { txId: id, soh: it.soh - q, name: it.name };
  } finally { lock.releaseLock(); }
}

/* ------------------------------ replenishment: refill to Par Max ------------------------------ */
function adcRefill(token, codes) {
  const a = adcUser_(token, 'refill');
  if (!Array.isArray(codes) || !codes.length || codes.length > 200) throw new Error('INVALID_INPUT');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const v = adcView_(), by = {}; v.items.forEach(function (x) { by[x.code] = x; });
    const now = new Date(), tz = Session.getScriptTimeZone(); let id = v.maxId, units = 0;
    const rows = [];
    // each entry: a code (refill up to Par Max) or {code, qty} (the quantity actually put in the drawer)
    codes.forEach(function (c) {
      const code = typeof c === 'object' && c ? String(c.code || '') : String(c), it = by[code]; if (!it) return; adcInScope_(a, it.adc);
      let q = it.parMax - it.soh;
      if (typeof c === 'object' && c) { q = Number(c.qty); if (!(q === Math.floor(q) && q > 0 && q <= 100000)) throw new Error('INVALID_INPUT'); }
      if (q <= 0) return;
      // the ward chosen by the technician (one of the wards of this item's cabinet); otherwise the item's primary ward
      let ward = it.ward;
      if (typeof c === 'object' && c && c.ward) { const w = adcWards_().filter(function (x) { return x.ward === String(c.ward) && x.adc === it.adc; })[0]; if (!w) throw new Error('INVALID_INPUT'); ward = w.ward; }
      rows.push(['TX-' + (++id), Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'), v.period, adcShift_(now), 'RECEIVED', it.code, it.name, it.adc, ward,
        a.u.displayName + ' (@' + a.u.username + ')', adcRole_(a.u), q, adcVal_(q, it.cost), '', typeof c === 'object' && c ? 'Refill from main pharmacy' : 'Refill to Par Max from main pharmacy']);
      units += q;
    });
    adcAppend_(rows); SpreadsheetApp.flush();
    return { lines: rows.length, units: units };
  } finally { lock.releaseLock(); }
}

/* ------------------------------ inventory pharmacist: count and monthly zeroing ------------------------------ */
/** Chunked storage in ScriptProperties (4,000 chars/chunk stays safely under the 9KB UTF-8 property limit). */
function adcDraft_() {
  const P = PropertiesService.getScriptProperties(), raw = P.getProperty(ADC.DRAFT);
  if (!raw) return {};
  try {
    if (raw.charAt(0) === '{') return JSON.parse(raw);   // backward compatibility with legacy single-key draft
    const n = Number(raw); if (!(n > 0)) return {};
    let txt = '';
    for (let i = 0; i < n; i++) txt += (P.getProperty(ADC.DRAFT + '.' + i) || '');
    return txt ? JSON.parse(txt) : {};
  } catch (e) { return {}; }
}
function adcDraftSave_(d) {
  const P = PropertiesService.getScriptProperties(), prev = Number(P.getProperty(ADC.DRAFT)) || 0;
  const txt = JSON.stringify(d || {}), size = 4000, map = {}; let n = 0;
  for (let i = 0; i < txt.length; i += size) map[ADC.DRAFT + '.' + (n++)] = txt.slice(i, i + size);
  map[ADC.DRAFT] = String(n);
  P.setProperties(map);
  for (let i = n; i < prev; i++) P.deleteProperty(ADC.DRAFT + '.' + i);
}
function adcDraftClear_() {
  const P = PropertiesService.getScriptProperties(), prev = Number(P.getProperty(ADC.DRAFT)) || 0;
  P.deleteProperty(ADC.DRAFT);
  for (let i = 0; i < prev; i++) P.deleteProperty(ADC.DRAFT + '.' + i);
}

function adcCountSave(token, counts) {
  const a = adcUser_(token, 'count'), sc = adcScope_(a);
  if (!counts || typeof counts !== 'object') throw new Error('INVALID_INPUT');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = adcDraft_(), codes = {};
    adcItems_().forEach(function (x) { if (!sc || sc.indexOf(x.adc) > -1) codes[x.code] = 1; });
    Object.keys(counts).forEach(function (c) {
      if (!codes[c]) return; const k = counts[c] || {}, p = k.physical;
      if (p === '' || p === null || p === undefined) { delete d[c]; return; }
      const n = Number(p); if (!(n === Math.floor(n) && n >= 0 && n <= 100000)) throw new Error('INVALID_INPUT');
      const ex = String(k.expiry || '').trim(); if (ex && !/^\d{4}-(0[1-9]|1[0-2])$/.test(ex)) throw new Error('INVALID_INPUT');
      d[c] = { physical: n, expiry: ex, cause: sheetText_(k.cause, 300) };
    });
    adcDraftSave_(d);
    return { saved: Object.keys(d).length };
  } finally { lock.releaseLock(); }
}

function adcCountApprove(token) {
  const a = adcUser_(token, 'count');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const v = adcView_(), d = adcDraft_();
    const missing = v.items.filter(function (x) { return !d[x.code]; });
    if (missing.length) throw new Error('ADC_COUNT_INCOMPLETE:' + missing.length);
    // root cause is no longer required: the count screen records drug · physical quantity · expiry only
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const rows = v.items.map(function (x) { const p = d[x.code].physical, vq = p - x.soh;
      return [v.period, x.code, x.name, x.opening, x.rec, x.out, x.soh, p, vq, adcVal_(vq, x.cost),
        (d[x.code].cause || '') || (vq === 0 ? 'Zero variance' : ''), a.u.displayName + ' (@' + a.u.username + ')', now, 'APPROVED']; });
    const ct = adcSheet_(ADC.COUNTS); ct.getRange(ct.getLastRow() + 1, 1, rows.length, ADC.COUNT_HEAD.length).setValues(rows);
    // one tab per month: drug · physical quantity · expiry (the count of that month, starting fresh every month)
    const mName = adcMonthTab_(v.period, v.items.map(function (x) { return [x.name, x.code, x.adc, d[x.code].physical, adcExpTxt_(d[x.code].expiry)]; }),
      a.u.displayName, now);
    // zeroing: the physical count becomes the opening balance of the next month
    const it = adcSheet_(ADC.ITEMS), items = adcItems_(), op = {};
    v.items.forEach(function (x) { op[x.code] = d[x.code].physical; });
    it.getRange(2, 10, items.length, 1).setValues(items.map(function (x) { return [op[x.code] !== undefined ? op[x.code] : x.opening]; }));
    const p = v.period.split('-').map(Number), next = p[1] === 12 ? (p[0] + 1) + '-01' : p[0] + '-' + String(p[1] + 1).padStart(2, '0');
    PropertiesService.getScriptProperties().setProperty(ADC.PROP, next);
    adcDraftClear_();
    try { audit_('adc_count_approved', a.u.username, a.u.username, v.period + ' → ' + next, { actorId: a.u.id }); } catch (e) {}
    SpreadsheetApp.flush();
    const net = rows.reduce(function (s, r) { return s + r[8]; }, 0), priced = rows.filter(function (r) { return r[9] !== ''; });
    return { closed: v.period, next: next, sheet: mName, items: rows.length, netVariance: net,
      varianceSAR: priced.length ? Math.round(priced.reduce(function (s, r) { return s + r[9]; }, 0) * 100) / 100 : null, pricedItems: priced.length };
  } finally { lock.releaseLock(); }
}

/** 2026-11 → 11/2026 (blank stays blank). */
function adcExpTxt_(ym) { const m = /^(\d{4})-(\d{2})$/.exec(String(ym || '')); return m ? m[2] + '/' + m[1] : ''; }
/** Writes (or rewrites) the tab of one month: «جرد الكابينات YYYY-MM» with drug · code · cabinet · physical qty · expiry. */
function adcMonthTab_(period, rows, by, at) {
  const ss = SpreadsheetApp.getActive(), name = 'جرد الكابينات ' + period;
  let sh = ss.getSheetByName(name); if (sh) sh.clear(); else sh = ss.insertSheet(name);
  sh.setRightToLeft(true);
  sh.getRange(1, 1).setValue('جرد الكابينات الشهري · Monthly ADC count — ' + period).setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setValue('اعتمده · Approved by: ' + by + ' — ' + at).setFontColor('#64748B');
  const head = ['الدواء · Drug', 'الرمز · Code', 'الكابينة · Cabinet', 'الكمية المتوفرة (العدد الفعلي) · Physical qty', 'تاريخ الانتهاء · Expiry'];
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#0F2942').setWrap(true).setVerticalAlignment('middle');
  if (rows.length) {
    sh.getRange(5, 2, rows.length, 1).setNumberFormat('@'); sh.getRange(5, 5, rows.length, 1).setNumberFormat('@');
    sh.getRange(5, 1, rows.length, head.length).setValues(rows);
  }
  sh.setFrozenRows(4); sh.setColumnWidth(1, 320); sh.setColumnWidths(2, 2, 170); sh.setColumnWidths(4, 2, 150);
  return name;
}

/* ------------------------------ supervisor: executive view ------------------------------ */
/** Supervisor: monthly count results for a chosen range of months (item by item) + progress of the open count. */
function adcCountDetail(token, from, to) {
  adcUser_(token, 'sup');
  const okm = function (m) { return /^\d{4}-\d{2}$/.test(String(m || '')); };
  const ct = adcSheet_(ADC.COUNTS), n = ct.getLastRow() - 1, rows = n > 0 ? ct.getRange(2, 1, n, ADC.COUNT_HEAD.length).getValues() : [];
  const tz = Session.getScriptTimeZone(), months = {};
  rows.forEach(function (r) { months[String(r[0])] = 1; });
  const ms = Object.keys(months).filter(okm).sort();
  if (!okm(from)) from = ms.length ? ms[ms.length - 1] : ''; if (!okm(to)) to = from;
  if (from > to) { const x = from; from = to; to = x; }
  const items = adcItems_(), adc = {}; items.forEach(function (x) { adc[x.code] = x.adc; });
  const list = rows.filter(function (r) { const m = String(r[0]); return m >= from && m <= to; }).map(function (r) {
    return { month: String(r[0]), code: String(r[1]), name: String(r[2]), adc: adc[String(r[1])] || '', opening: adcNum_(r[3]), rec: adcNum_(r[4]), out: adcNum_(r[5]),
      system: adcNum_(r[6]), physical: adcNum_(r[7]), variance: adcNum_(r[8]), sar: r[9] === '' ? '' : adcNum_(r[9]), cause: String(r[10] || ''), by: String(r[11] || ''),
      at: r[12] instanceof Date ? Utilities.formatDate(r[12], tz, 'yyyy-MM-dd HH:mm') : String(r[12] || '') };
  });
  const d = adcDraft_(), prog = ADC.CLUSTERS.map(function (c) { const its = items.filter(function (x) { return x.adc === c[0]; });
    return { adc: c[0], total: its.length, counted: its.filter(function (x) { return d[x.code]; }).length }; });
  return { months: ms, from: from, to: to, rows: list, open: { period: adcPeriod_(), progress: prog } };
}

function adcSummary(token, month, cluster) {
  adcUser_(token, 'sup');
  const tx = adcTx_(), items = adcItems_(), alert = {}, itemAdc = {};
  items.forEach(function (x) { alert[x.code] = x.alert; itemAdc[x.code] = x.adc; });
  const pick = function (r) { return (month === 'all' || String(r[2]) === month) && (cluster === 'all' || String(r[7]) === cluster); };
  const tot = { rec: 0, scan: 0, over: 0, outSAR: 0 }, ward = {}, staff = {}, months = {};
  tx.forEach(function (r) {
    months[String(r[2])] = 1; if (!pick(r)) return;
    const t = String(r[4]), q = adcNum_(r[11]), sar = adcNum_(r[12]), w = String(r[8] || '—'), s = String(r[9] || '—');
    const W = ward[w] || (ward[w] = { ward: w, adc: String(r[7]), rec: 0, scan: 0, over: 0, outSAR: 0 });
    const S = staff[s] || (staff[s] = { user: s, role: String(r[10]), rec: 0, scan: 0, over: 0, ctrl: 0 });
    if (t === 'RECEIVED' || t === 'RETURN') { tot.rec += q; W.rec += q; S.rec += q; }
    else if (t === 'SCANNED') { tot.scan += q; tot.outSAR += sar; W.scan += q; W.outSAR += sar; S.scan += q; }
    else if (t === 'OVERRIDE') { tot.over += q; tot.outSAR += sar; W.over += q; W.outSAR += sar; S.over += q; if (alert[String(r[5])] === 'CONTROLLED') S.ctrl += q; }
  });
  const bc = function (o) { return o.scan + o.over ? Math.round(o.scan / (o.scan + o.over) * 1000) / 10 : 100; };
  const counts = adcSheet_(ADC.COUNTS), cn = counts.getLastRow() - 1, cv = cn > 0 ? counts.getRange(2, 1, cn, ADC.COUNT_HEAD.length).getValues() : [];
  const varByMonth = {};
  cv.forEach(function (r) {
    if (cluster !== 'all' && itemAdc[String(r[1])] !== cluster) return;
    const m = String(r[0]); const o = varByMonth[m] || (varByMonth[m] = { month: m, items: 0, zero: 0, net: 0, sar: 0 });
    o.items++; if (adcNum_(r[8]) === 0) o.zero++; o.net += adcNum_(r[8]); o.sar += adcNum_(r[9]);
  });
  const priced = items.some(function (x) { return x.cost; });
  return { priced: priced, totals: { rec: tot.rec, scan: tot.scan, over: tot.over, outSAR: Math.round(tot.outSAR * 100) / 100, bcma: bc(tot) },
    wards: Object.keys(ward).map(function (k) { const o = ward[k]; o.outSAR = Math.round(o.outSAR * 100) / 100; o.bcma = bc(o); return o; })
      .sort(function (a, b) { return (b.scan + b.over) - (a.scan + a.over); }),
    staff: Object.keys(staff).map(function (k) { const o = staff[k]; o.bcma = bc(o); return o; }).sort(function (a, b) { return b.over - a.over || (b.scan + b.rec) - (a.scan + a.rec); }),
    variance: Object.keys(varByMonth).sort().map(function (k) { const o = varByMonth[k]; o.sar = Math.round(o.sar * 100) / 100; o.accuracy = o.items ? Math.round(o.zero / o.items * 1000) / 10 : 0; return o; }),
    months: Object.keys(months).filter(String).sort(), clusters: ADC.CLUSTERS.map(function (c) { return c[0]; }) };
}

/** Everything the browser needs to rebuild the master Excel workbook with live formulas. */
function adcExport(token) {
  const a = adcUser_(token); if (!(a.p.sup || a.p.count || a.p.refill)) throw new Error('FORBIDDEN');
  const v = adcView_(), counts = adcSheet_(ADC.COUNTS), cn = counts.getLastRow() - 1;
  // staff assigned to cabinets get only their cabinets; supervisors get everything
  const sc = adcScope_(a);
  if (sc) {
    const codes = {}; v.items.forEach(function (x) { if (sc.indexOf(x.adc) > -1) codes[x.code] = 1; });
    return { period: v.period, items: v.items.filter(function (x) { return codes[x.code]; }),
      tx: adcTx_().filter(function (r) { return codes[String(r[5])]; }).map(function (r) { return r.map(function (c) { return c instanceof Date ? Utilities.formatDate(c, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : c; }); }),
      counts: (cn > 0 ? counts.getRange(2, 1, cn, ADC.COUNT_HEAD.length).getValues() : []).filter(function (r) { return codes[String(r[1])]; }) };
  }
  return { period: v.period, items: v.items, tx: adcTx_().map(function (r) { return r.map(function (c) { return c instanceof Date ? Utilities.formatDate(c, Session.getScriptTimeZone(), 'yyyy-MM-dd') : c; }); }),
    counts: cn > 0 ? counts.getRange(2, 1, cn, ADC.COUNT_HEAD.length).getValues() : [] };
}