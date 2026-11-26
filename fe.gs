/** Edits an existing item's name, strength and dosage form. The serial number never changes, so printed labels stay valid. */
function formEditItem(token, no, item) {
  const u = formUser_(token, true); item = item || {}; no = String(no || '').trim();
  const name = sheetText_(item.name, 120), strength = sheetText_(item.strength, 60), form = sheetText_(item.form, 60);
  if (!no) throw new Error('INVALID_INPUT');
  if (name.length < 2) throw new Error('اسم الصنف مطلوب');
  const ss = SpreadsheetApp.getActive(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let hit = null, dup = false;
    DRUG_SHEETS.forEach(function (d) { const s = ss.getSheetByName(d.name); if (!s || s.getLastRow() < 2) return;
      s.getRange(2, 1, s.getLastRow() - 1, 4).getDisplayValues().forEach(function (r, i) { const n = String(r[0]).trim();
        if (n === no) hit = { sh: s, row: i + 2, old: r[1] + ' ' + r[3] };
        else if (String(r[1]).trim().toLowerCase() === name.toLowerCase() && String(r[3]).trim().toLowerCase() === strength.toLowerCase()) dup = true; }); });
    if (!hit) throw new Error('NOT_FOUND');
    if (dup) throw new Error('الصنف موجود مسبقًا بنفس التركيز');
    hit.sh.getRange(hit.row, 2).setValue(name); hit.sh.getRange(hit.row, 4).setValue(strength); hit.sh.getRange(hit.row, 5).setValue(form);
    audit_('formulary_item_edited', '', u.username, '#' + no + ' ' + hit.old + ' → ' + name + ' ' + strength, {});
    clearDataCache_();
    formularyBuild_();
    return { ok: true, no: no };
  } finally { lock.releaseLock(); }
}

