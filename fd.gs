/** Deletes an item from its list (after the user confirms on the page). */
function formDeleteItem(token, no) {
  const u = formUser_(token, true); no = String(no || '').trim(); if (!no) throw new Error('INVALID_INPUT');
  const ss = SpreadsheetApp.getActive(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let hit = null;
    DRUG_SHEETS.forEach(function (d) { const s = ss.getSheetByName(d.name); if (hit || !s || s.getLastRow() < 2) return;
      s.getRange(2, 1, s.getLastRow() - 1, 4).getDisplayValues().forEach(function (r, i) { if (!hit && String(r[0]).trim() === no) hit = { sh: s, row: i + 2, old: r[1] + ' ' + r[3] }; }); });
    if (!hit) throw new Error('NOT_FOUND');
    hit.sh.deleteRow(hit.row);
    audit_('formulary_item_deleted', '', u.username, '#' + no + ' ' + hit.old, {});
    clearDataCache_();
    formularyBuild_();
    return { ok: true, no: no };
  } finally { lock.releaseLock(); }
}

