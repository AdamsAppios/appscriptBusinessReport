/***************************************
 * TalambanUnpaid.js
 * - G1: paste unpaid rows from Talamban within B1:D1
 * - G2: generate output to H2 (paid vs unpaid)
 * - J2: SAVE -> converts selected Unpaid to Paid back in Talamban Expenses
 * - Column D (row 3+): paid checkbox -> refresh output
 * - Column E (row 3+): do-not-display checkbox -> exclude from output + refresh output
 ****************************************/

const TU_CFG = {
  destSheetName: "TalambanUnpaid",
  sourceSheetName: "Talamban",

  // control cells (in TalambanUnpaid)
  fromCell: "B1",
  toCell: "D1",
  pasteBtnCell: "G1",
  genBtnCell: "G2",
  outputCell: "H2",
  saveBtnCell: "J2", // ✅ NEW

  headerRow: 2,
  dataStartRow: 3,

  // expected headers in Talamban
  srcDateHeader: "Date",
  srcExpensesHeader: "Expenses",

  // destination columns (TalambanUnpaid)
  destDateCol: 1,        // A
  destNameCol: 2,        // B
  destAmountCol: 3,      // C
  destPaidCol: 4,        // D  (Paid Nakuha)
  destHideCol: 5         // E  (Do not Display?)
};

/**
 * Called by main onEdit(e) safely (try/catch in OnEdit.js)
 */
function handleTalambanUnpaidOnEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== TU_CFG.destSheetName) return;

  const a1 = e.range.getA1Notation();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // G1 button (Paste)
  if (a1 === TU_CFG.pasteBtnCell && _tu_isTrue_(e.value)) {
    tuPasteUnpaidFromTalamban_();
    tuGenerateOutput_();
    sh.getRange(TU_CFG.pasteBtnCell).setValue(false);
    return;
  }

  // G2 button (Generate Output)
  if (a1 === TU_CFG.genBtnCell && _tu_isTrue_(e.value)) {
    tuGenerateOutput_();
    sh.getRange(TU_CFG.genBtnCell).setValue(false);
    return;
  }

  // ✅ J2 button (Save)
  if (a1 === TU_CFG.saveBtnCell && _tu_isTrue_(e.value)) {
    tuSavePaidBackToTalamban_();
    // refresh list + output so saved rows disappear (now "Paid", not "Unpaid")
    tuPasteUnpaidFromTalamban_();
    tuGenerateOutput_();
    sh.getRange(TU_CFG.saveBtnCell).setValue(false);
    return;
  }

  // Paid checkbox column (D3+) OR Do-not-display column (E3+): auto-refresh output
  if (
    row >= TU_CFG.dataStartRow &&
    (col === TU_CFG.destPaidCol || col === TU_CFG.destHideCol)
  ) {
    tuGenerateOutput_();
    return;
  }
}

/**
 * Reads Talamban rows within date range (B1:D1 in TalambanUnpaid),
 * extracts "Unpaid Name=Amount" from Expenses, and writes to TalambanUnpaid A:C.
 * Preserves old states for Paid(D) and DoNotDisplay(E) by matching Date+Name+Amount.
 */
function tuPasteUnpaidFromTalamban_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dest = ss.getSheetByName(TU_CFG.destSheetName);
  const src = ss.getSheetByName(TU_CFG.sourceSheetName);
  if (!dest || !src) return;

  const fromDate = _tu_parseDate_(dest.getRange(TU_CFG.fromCell).getValue());
  const toDate = _tu_parseDate_(dest.getRange(TU_CFG.toCell).getValue());
  if (!fromDate || !toDate) throw new Error("TalambanUnpaid B1/D1 must be valid dates.");

  const srcDateCol = _tu_findHeaderCol_(src, TU_CFG.srcDateHeader, TU_CFG.headerRow) || 1;
  const srcExpCol  = _tu_findHeaderCol_(src, TU_CFG.srcExpensesHeader, TU_CFG.headerRow) || 24;

  const lastRow = src.getLastRow();
  const numRows = Math.max(0, lastRow - TU_CFG.headerRow);
  if (numRows === 0) return;

  const dateVals = src.getRange(TU_CFG.headerRow + 1, srcDateCol, numRows, 1).getValues();
  const expVals  = src.getRange(TU_CFG.headerRow + 1, srcExpCol,  numRows, 1).getValues();

  // preserve existing paid + hide states
  const stateMap = _tu_buildStateMap_(dest);

  const rowsOut = [];
  const fromD0 = _tu_dayStart_(fromDate);
  const toD0 = _tu_dayStart_(toDate);

  for (let i = 0; i < numRows; i++) {
    const d = _tu_parseDate_(dateVals[i][0]);
    if (!d) continue;

    const d0 = _tu_dayStart_(d);
    if (d0 < fromD0 || d0 > toD0) continue;

    const exp = expVals[i][0];
    const unpaidEntries = _tu_extractUnpaid_(exp); // ONLY Unpaid

    unpaidEntries.forEach(ent => {
      const key = _tu_key_(d0, ent.name, ent.amount);
      const prev = stateMap.get(key) || { paid: false, hide: false };
      rowsOut.push({
        date: d0,
        name: ent.name,
        amount: ent.amount,
        paid: prev.paid === true,
        hide: prev.hide === true
      });
    });
  }

  rowsOut.sort((a, b) => {
    const da = a.date.getTime(), db = b.date.getTime();
    if (da !== db) return da - db;
    return String(a.name).localeCompare(String(b.name));
  });

  // clear old content (A:E from row3 down)
  const destLast = Math.max(dest.getLastRow(), TU_CFG.dataStartRow);
  const clearRows = Math.max(1, destLast - TU_CFG.dataStartRow + 1);
  dest.getRange(TU_CFG.dataStartRow, 1, clearRows, 5).clearContent();

  if (rowsOut.length === 0) {
    dest.getRange(TU_CFG.dataStartRow, TU_CFG.destPaidCol, 1, 1).insertCheckboxes();
    dest.getRange(TU_CFG.dataStartRow, TU_CFG.destHideCol, 1, 1).insertCheckboxes();
    return;
  }

  // write A:C
  dest.getRange(TU_CFG.dataStartRow, TU_CFG.destDateCol, rowsOut.length, 3)
    .setValues(rowsOut.map(r => [r.date, r.name, r.amount]));

  // D Paid
  const paidRange = dest.getRange(TU_CFG.dataStartRow, TU_CFG.destPaidCol, rowsOut.length, 1);
  paidRange.insertCheckboxes();
  paidRange.setValues(rowsOut.map(r => [r.paid === true]));

  // E Do not Display
  const hideRange = dest.getRange(TU_CFG.dataStartRow, TU_CFG.destHideCol, rowsOut.length, 1);
  hideRange.insertCheckboxes();
  hideRange.setValues(rowsOut.map(r => [r.hide === true]));
}

/**
 * Generates the formatted output into H2:
 * - paid nakuha na (D checked)
 * - unpaid wala pa nakuha (D unchecked)
 * Excludes rows where E (Do not Display) is checked.
 */
function tuGenerateOutput_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TU_CFG.destSheetName);
  if (!sh) return;

  const fromDate = _tu_parseDate_(sh.getRange(TU_CFG.fromCell).getValue());
  const toDate = _tu_parseDate_(sh.getRange(TU_CFG.toCell).getValue());
  const fromD = fromDate ? _tu_dayStart_(fromDate) : null;
  const toD = toDate ? _tu_dayStart_(toDate) : null;

  const lastRow = sh.getLastRow();
  const numRows = Math.max(0, lastRow - (TU_CFG.dataStartRow - 1));
  if (numRows === 0) {
    sh.getRange(TU_CFG.outputCell).setValue("");
    return;
  }

  // A:E
  const data = sh.getRange(TU_CFG.dataStartRow, 1, numRows, 5).getValues();

  const paidByDate = new Map();
  const unpdByDate = new Map();

  for (const r of data) {
    const d = _tu_parseDate_(r[0]);
    const name = (r[1] || "").toString().trim();
    const amount = r[2];
    const paid = r[3] === true;
    const hide = r[4] === true;

    if (hide) continue;
    if (!d || !name || amount === "" || amount === null) continue;

    const d0 = _tu_dayStart_(d);
    if (fromD && d0 < fromD) continue;
    if (toD && d0 > toD) continue;

    const tag = paid ? "Paid" : "Unpaid";
    const line = `${tag} ${name}=${_tu_fmtAmount_(amount)};`;

    const key = _tu_dateKey_(d0);
    const map = paid ? paidByDate : unpdByDate;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(line);
  }

  const tz = Session.getScriptTimeZone();
  const out = [];
  out.push("paid nakuha na");
  out.push(_tu_renderGrouped_(paidByDate, tz));
  out.push("");
  out.push("unpaid wala pa nakuha");
  out.push(_tu_renderGrouped_(unpdByDate, tz));

  sh.getRange(TU_CFG.outputCell).setValue(out.join("\n").trim());
}

/**
 * ✅ NEW: Saves back to Talamban Expenses by converting Unpaid -> Paid
 * for every TalambanUnpaid row with Paid checkbox checked (D = TRUE).
 */
function tuSavePaidBackToTalamban_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tu = ss.getSheetByName(TU_CFG.destSheetName);
  const tal = ss.getSheetByName(TU_CFG.sourceSheetName);
  if (!tu || !tal) return;

  const tuLast = tu.getLastRow();
  const tuNum = Math.max(0, tuLast - (TU_CFG.dataStartRow - 1));
  if (tuNum === 0) return;

  // read A:E from TalambanUnpaid
  const tuData = tu.getRange(TU_CFG.dataStartRow, 1, tuNum, 5).getValues();

  // build map dateKey -> [{name, amount}]
  const byDate = new Map();
  for (const r of tuData) {
    const d = _tu_parseDate_(r[0]);
    const name = (r[1] || "").toString().trim();
    const amt = r[2];
    const paidChecked = r[3] === true; // D
    if (!paidChecked) continue;
    if (!d || !name || amt === "" || amt === null) continue;

    const d0 = _tu_dayStart_(d);
    const key = _tu_dateKey_(d0);
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push({ name, amount: amt });
  }

  if (byDate.size === 0) return;

  // locate Talamban columns
  const dateCol = _tu_findHeaderCol_(tal, TU_CFG.srcDateHeader, TU_CFG.headerRow) || 1;
  const expCol  = _tu_findHeaderCol_(tal, TU_CFG.srcExpensesHeader, TU_CFG.headerRow) || 24;

  const talLast = tal.getLastRow();
  const talNum = Math.max(0, talLast - TU_CFG.headerRow);
  if (talNum === 0) return;

  // read Talamban dates + expenses
  const talDates = tal.getRange(TU_CFG.headerRow + 1, dateCol, talNum, 1).getValues();
  const talExps  = tal.getRange(TU_CFG.headerRow + 1, expCol,  talNum, 1).getValues();

  // dateKey -> index
  const talIndexByDate = new Map();
  for (let i = 0; i < talNum; i++) {
    const d = _tu_parseDate_(talDates[i][0]);
    if (!d) continue;
    const key = _tu_dateKey_(_tu_dayStart_(d));
    if (!talIndexByDate.has(key)) talIndexByDate.set(key, i);
  }

  // apply replacements in memory, then write only changed rows
  const updates = [];
  for (const [key, entries] of byDate.entries()) {
    const idx = talIndexByDate.get(key);
    if (idx === undefined) continue;

    const oldText = talExps[idx][0];
    let newText = (oldText === null || oldText === undefined) ? "" : String(oldText);

    for (const ent of entries) {
      newText = _tu_replaceUnpaidWithPaid_(newText, ent.name, ent.amount);
    }

    if (newText !== String(oldText ?? "")) {
      const rowNumber = TU_CFG.headerRow + 1 + idx;
      updates.push({ row: rowNumber, value: newText });
    }
  }

  // write back
  for (const u of updates) {
    tal.getRange(u.row, expCol).setValue(u.value);
  }
}

/* =======================
   Helpers (private)
   ======================= */

function _tu_isTrue_(v) {
  return String(v).toUpperCase() === "TRUE";
}

function _tu_findHeaderCol_(sheet, headerText, headerRow) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const target = String(headerText).trim().toLowerCase();
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim().toLowerCase() === target) return c + 1;
  }
  return null;
}

function _tu_parseDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const mm = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    if (!isNaN(d.getTime())) return d;
  }

  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function _tu_dayStart_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function _tu_extractUnpaid_(expensesCell) {
  if (expensesCell === null || expensesCell === "") return [];
  const text = String(expensesCell).replace(/\n/g, ";");

  // matches: Unpaid Name = 123   (case-insensitive)
  const re = /unpaid\s+([^=;]+?)\s*=\s*([0-9]+(?:\.[0-9]+)?)/ig;

  const out = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    const name = String(m[1]).trim();
    const amt = parseFloat(m[2]);
    if (!name || isNaN(amt)) continue;
    out.push({ name, amount: amt });
  }
  return out;
}

function _tu_dateKey_(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function _tu_key_(dateObj, name, amount) {
  return `${_tu_dateKey_(dateObj)}|${String(name).trim().toLowerCase()}|${_tu_fmtAmount_(amount)}`;
}

function _tu_fmtAmount_(v) {
  const n = typeof v === "number" ? v : parseFloat(v);
  if (isNaN(n)) return String(v);
  return Number.isInteger(n) ? String(n) : String(n);
}

function _tu_buildStateMap_(destSheet) {
  const map = new Map();
  const lastRow = destSheet.getLastRow();
  const numRows = Math.max(0, lastRow - (TU_CFG.dataStartRow - 1));
  if (numRows === 0) return map;

  const data = destSheet.getRange(TU_CFG.dataStartRow, 1, numRows, 5).getValues();
  for (const r of data) {
    const d = _tu_parseDate_(r[0]);
    const name = (r[1] || "").toString().trim();
    const amt = r[2];
    const paid = r[3] === true;
    const hide = r[4] === true;
    if (!d || !name || amt === "" || amt === null) continue;

    const d0 = _tu_dayStart_(d);
    map.set(_tu_key_(d0, name, amt), { paid, hide });
  }
  return map;
}

function _tu_renderGrouped_(mapByDate, tz) {
  const keys = Array.from(mapByDate.keys()).sort(); // yyyy-mm-dd sorts properly
  if (keys.length === 0) return "";

  const chunks = [];
  for (const k of keys) {
    const parts = k.split("-");
    const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    const dateLine = Utilities.formatDate(d, tz, "MM/dd/yyyy");

    chunks.push(dateLine);
    chunks.push(mapByDate.get(k).join(" "));
    chunks.push("");
  }
  return chunks.join("\n").trimEnd();
}

function _tu_escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Replace only the matching amount:
 * "Unpaid Name=440" -> "Paid Name=440"
 */
function _tu_replaceUnpaidWithPaid_(text, name, amount) {
  const nm = String(name).trim();
  const nmEsc = _tu_escapeRegex_(nm);
  const amtNum = (typeof amount === "number") ? amount : parseFloat(amount);

  const re = new RegExp(`\\bunpaid\\s+${nmEsc}\\s*=\\s*([0-9]+(?:\\.[0-9]+)?)`, "ig");

  return String(text).replace(re, (match, numStr) => {
    const n = parseFloat(numStr);
    if (!isNaN(amtNum) && !isNaN(n) && Math.abs(n - amtNum) < 1e-9) {
      return `Paid ${nm}=${numStr}`;
    }
    return match; // not the same amount, don't change
  });
}
