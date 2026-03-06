var DISPLAY_SWELDO_DOC_ID = "1exhAo09Fu_lIkr2bWi6E7I9QcIvimVtKaKs3lP7lhQY"; // Put your "Display Sweldo" Google Doc ID here.

function displaySalary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName("DisplaySalary");
  var showAll = ds.getRange("A2").getValue();

  // Detect script-generated rows only so manual rows (ex: A8:A10) are preserved.
  var lastRow = ds.getLastRow();
  var oldA = lastRow > 2 ? ds.getRange(3, 1, lastRow - 2, 1).getValues() : [];
  var rowsToClear = [];
  for (var i = 0; i < oldA.length; i++) {
    var cellTxt = String(oldA[i][0] || "");
    if (/calculation days\s*:/i.test(cellTxt)) rowsToClear.push(3 + i);
  }

  var hdrRow = ds.getRange(1, 1, 1, ds.getLastColumn()).getValues()[0];
  var totalCol = 8; // default H
  for (var c = 0; c < hdrRow.length; c++) {
    var t = (hdrRow[c] || "").toString().toLowerCase();
    if (t && t.indexOf("total") === 0) {
      totalCol = c + 1;
      break;
    }
  }

  rowsToClear.forEach(function(r) {
    ds.getRange(r, 1).clearContent();
    ds.getRange(r, 6).clearContent();
    ds.getRange(r, totalCol).clearContent();
  });

  if (!showAll) return;

  // Date windows
  var attRange = ds.getRange("B2:C2").getValues()[0]; // attendance & CA window
  var otRange = ds.getRange("D2:E2").getValues()[0]; // OT/UT/HD window
  var attStart = attRange[0];
  var attEnd = attRange[1];
  var otStart = otRange[0];
  var otEnd = otRange[1];
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Manila";

  // Accumulators
  var baseCount = {};
  var contribList = {};
  var contribFraction = {};
  var cashAdvTotal = {};
  var caLogByEmp = {};

  // Scan both stores
  ["Talamban", "Labangon"].forEach(function(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    var data = sh.getDataRange().getValues();
    var hdr = data[1];
    var iDate = hdr.indexOf("Date");
    var iExp = hdr.indexOf("Expenses");
    var iDuty = hdr.indexOf("Duty");
    var maxHStore = sheetName === "Talamban" ? 12 : 11.5;

    data.slice(2).forEach(function(r) {
      var d = r[iDate];
      if (!(d instanceof Date) || isNaN(d)) return;

      var inAtt = attStart && attEnd && d >= attStart && d <= attEnd;
      var inOT = otStart && otEnd && d >= otStart && d <= otEnd;

      // CA capture (attendance window only)
      if (inAtt) {
        var txt = (r[iExp] || "") + "";
        var reCA = /CA\s+([^=;]+)=\s*([\d\.]+)/gi;
        var m;
        while ((m = reCA.exec(txt))) {
          var emp = m[1].trim();
          var amt = parseFloat(m[2]) || 0;
          cashAdvTotal[emp] = (cashAdvTotal[emp] || 0) + amt;
          (caLogByEmp[emp] = caLogByEmp[emp] || []).push({ d: d, amt: amt });
        }
      }

      // Duty parsing
      ("" + r[iDuty]).split(",").forEach(function(part) {
        var t = part.trim();
        if (!t) return;
        var emp = t.replace(/\(.*\)/, "").trim();

        // Base day only in attendance window
        if (inAtt) baseCount[emp] = (baseCount[emp] || 0) + 1;

        // OT/UT/HD only in OT window
        if (inOT) {
          if (/\(HD\)/i.test(t)) {
            (contribList[emp] = contribList[emp] || []).push("- hd 0.5");
            contribFraction[emp] = (contribFraction[emp] || 0) - 0.5;
          }
          var mo = t.match(/\(OT:\s*([\d\.]+)\)/i);
          if (mo) {
            var h = parseFloat(mo[1]) || 0;
            (contribList[emp] = contribList[emp] || []).push("ot " + h + "/" + maxHStore);
            contribFraction[emp] = (contribFraction[emp] || 0) + (h / maxHStore);
          }
          var mu = t.match(/\(UT:\s*([\d\.]+)\)/i);
          if (mu) {
            var h2 = parseFloat(mu[1]) || 0;
            (contribList[emp] = contribList[emp] || []).push("- ut (1-" + h2 + "/" + maxHStore + ")");
            contribFraction[emp] = (contribFraction[emp] || 0) - (1 - h2 / maxHStore);
          }
        }
      });
    });
  });

  // Attendance sheet (rates & fixed add/deduct)
  var at = ss.getSheetByName("Attendance");
  var A = at.getDataRange().getValues();
  var Hh = A[0];
  var idx = function(col) { return Hh.indexOf(col); };
  var cName = idx("Name");
  var cDaily = idx("Daily 1");
  var cAccts = idx("Accts");
  var cCharges = idx("Charges");
  var cShort = idx("Charge Short");
  var cSSS = idx("SSS");
  var cPhil = idx("Philhealth");
  var cExtra = idx("Extra based on gallons sold");

  function roundHalfUp(n) { return n >= 0 ? Math.floor(n + 0.5) : Math.ceil(n - 0.5); }
  function fmtDate(d) { return Utilities.formatDate(d, tz, "MM/dd/yyyy"); }

  // Build and write outputs
  var aVals = [];
  var fVals = [];
  var hVals = [];

  Object.keys(baseCount).forEach(function(emp) {
    var base = baseCount[emp] || 0;
    var frac = contribFraction[emp] || 0;
    var totalDays = parseFloat((base + frac).toFixed(2));

    // Line 1
    var line1 = emp + " calculation days : " + base;
    (contribList[emp] || []).forEach(function(tok) {
      line1 += tok.charAt(0) === "-" ? " " + tok : " + " + tok;
    });
    line1 += " = " + totalDays;

    // Attendance row
    var rIdx = A.findIndex(function(r) { return r[cName] === emp; });
    var rate = rIdx > 0 ? parseFloat(A[rIdx][cDaily]) || 0 : 0;
    var extraG = rIdx > 0 ? parseFloat(A[rIdx][cExtra]) || 0 : 0;
    var accts = rIdx > 0 ? parseFloat(A[rIdx][cAccts]) || 0 : 0;
    var chgs = rIdx > 0 ? parseFloat(A[rIdx][cCharges]) || 0 : 0;
    var shrt = rIdx > 0 ? parseFloat(A[rIdx][cShort]) || 0 : 0;
    var sss = rIdx > 0 ? parseFloat(A[rIdx][cSSS]) || 0 : 0;
    var phil = rIdx > 0 ? parseFloat(A[rIdx][cPhil]) || 0 : 0;
    var caAmt = cashAdvTotal[emp] || 0;

    var basePay = totalDays * rate;
    var foodPay = totalDays * 60;
    var gross = basePay + foodPay;
    var net = gross + extraG - accts - chgs - shrt - sss - phil - caAmt;

    // Line 2: decimal net + rounded whole number
    var line2 = emp + ": " + totalDays + " days : "
      + totalDays + "xdaily rate " + rate
      + " + " + totalDays + "xfood allowance 60 = "
      + basePay.toFixed(0) + "+" + foodPay.toFixed(0)
      + " = " + gross.toFixed(2);
    if (extraG) line2 += " + Extra Based On Gallons " + extraG;
    if (caAmt) line2 += " - Cash Advances " + caAmt;
    if (chgs) line2 += " - Charges " + chgs;
    if (shrt) line2 += " - Charge Short " + shrt;
    if (sss) line2 += " - SSS " + sss;
    if (phil) line2 += " - Philhealth " + phil;
    if (accts) line2 += " - Accts " + accts;
    line2 += " = " + net.toFixed(2) + " (" + roundHalfUp(net) + ")";

    aVals.push([line1 + "\n" + line2]);

    // F-column: CA breakdown (attendance window)
    var caLogStr = "";
    if (caLogByEmp[emp] && caLogByEmp[emp].length) {
      caLogByEmp[emp].sort(function(a, b) { return a.d - b.d; });
      caLogStr = emp + " : " + caLogByEmp[emp]
        .map(function(e) { return fmtDate(e.d) + "=" + (e.amt || 0); })
        .join(" ; ");
    }
    fVals.push([caLogStr]);

    // Totals column (rounded half-up)
    hVals.push([roundHalfUp(net)]);
  });

  if (aVals.length) ds.getRange(3, 1, aVals.length, 1).setValues(aVals);
  if (fVals.length) ds.getRange(3, 6, fVals.length, 1).setValues(fVals);
  if (hVals.length) ds.getRange(3, totalCol, hVals.length, 1).setValues(hVals);

  appendDisplaySalaryToDoc_(ds, tz, attStart, attEnd);
}

function appendDisplaySalaryToDoc_(sheet, tz, startDate, endDate) {
  if (!DISPLAY_SWELDO_DOC_ID) return;
  if (!(startDate instanceof Date) || isNaN(startDate)) return;
  if (!(endDate instanceof Date) || isNaN(endDate)) return;

  var title = "Refilling Salary from "
    + Utilities.formatDate(startDate, tz, "MM/dd/yyyy")
    + " to "
    + Utilities.formatDate(endDate, tz, "MM/dd/yyyy");

  var sectionLines = buildDisplaySalarySectionLines_(sheet);
  var doc = DocumentApp.openById(DISPLAY_SWELDO_DOC_ID);
  var body = doc.getBody();
  var titleIndexes = findParagraphIndexesByText_(body, title);
  var titleIndex = -1;

  if (!titleIndexes.length) {
    if (body.getNumChildren() > 0) {
      var tail = getBodyChildText_(body, body.getNumChildren() - 1).trim();
      if (tail) body.appendParagraph("");
    }
    var newTitlePara = body.appendParagraph(title);
    setParagraphFullyBold_(newTitlePara, true);
    titleIndex = body.getNumChildren() - 1;
  } else {
    for (var d = titleIndexes.length - 1; d >= 1; d--) {
      var dupStart = titleIndexes[d];
      var dupEnd = findSectionEndIndex_(body, dupStart + 1);
      clearBodyRange_(body, dupStart, dupEnd);
    }
    titleIndex = findParagraphIndexByText_(body, title);
    var titlePara = body.getChild(titleIndex).asParagraph();
    titlePara.setText(title);
    setParagraphFullyBold_(titlePara, true);
  }

  var sectionEnd = findSectionEndIndex_(body, titleIndex + 1);
  clearBodyRange_(body, titleIndex + 1, sectionEnd);

  var insertAt = titleIndex + 1;
  if (!sectionLines.length) {
    body.insertParagraph(insertAt, "(No salary rows found)").setBold(false);
  } else {
    sectionLines.forEach(function(line) {
      body.insertParagraph(insertAt, line.text).setBold(!!line.bold);
      insertAt++;
    });
  }

  doc.saveAndClose();
}

function buildDisplaySalarySectionLines_(sheet) {
  var lastRow = sheet.getLastRow();
  var values = lastRow > 2 ? sheet.getRange(3, 1, lastRow - 2, 1).getValues() : [];
  var lines = [];
  var roundedParts = [];
  var roundedTotal = 0;

  values.forEach(function(row) {
    var cellText = String(row[0] || "").trim();
    if (!cellText) return;

    var perEmpLines = cellText
      .split(/\r?\n/)
      .map(function(s) { return s.trim(); })
      .filter(Boolean);
    if (!perEmpLines.length) return;

    perEmpLines.forEach(function(text) {
      lines.push({ text: text, bold: false });
    });
    lines.push({ text: "", bold: false });

    var lastNumber = extractLastNumber_(cellText);
    if (lastNumber === null) return;
    var rounded = roundHalfUp_(lastNumber);
    roundedParts.push(formatWholeWithCommas_(rounded));
    roundedTotal += rounded;
  });

  if (roundedParts.length) {
    lines.push({ text: "Total Salary for house and refilling", bold: false });
    lines.push({
      text: roundedParts.join("+") + "=" + formatWholeWithCommas_(roundedTotal),
      bold: false
    });
  }

  while (lines.length && !lines[lines.length - 1].text) lines.pop();
  return lines;
}

function findParagraphIndexByText_(body, exactText) {
  var all = findParagraphIndexesByText_(body, exactText);
  return all.length ? all[0] : -1;
}

function findParagraphIndexesByText_(body, exactText) {
  var matches = [];
  var normalizedTarget = normalizeHeaderText_(exactText);
  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    if (normalizeHeaderText_(el.asParagraph().getText()) === normalizedTarget) {
      matches.push(i);
    }
  }
  return matches;
}

function normalizeHeaderText_(text) {
  return String(text || "")
    .trim()
    .replace(/:$/, "")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function setParagraphFullyBold_(paragraph, isBold) {
  var makeBold = !!isBold;
  paragraph.setBold(makeBold);
  var text = paragraph.getText();
  if (!text) return;
  paragraph.editAsText().setBold(0, text.length - 1, makeBold);
}

function findSectionEndIndex_(body, fromIndex) {
  for (var i = fromIndex; i < body.getNumChildren(); i++) {
    var txt = getBodyChildText_(body, i).trim();
    if (!txt) continue;
    if (isSalarySectionHeader_(txt)) return i;
  }
  return body.getNumChildren();
}

function clearBodyRange_(body, startIndex, endExclusive) {
  for (var i = endExclusive - 1; i >= startIndex; i--) {
    body.removeChild(body.getChild(i));
  }
}

function getBodyChildText_(body, index) {
  var el = body.getChild(index);
  var t = el.getType();
  if (t === DocumentApp.ElementType.PARAGRAPH) return el.asParagraph().getText();
  if (t === DocumentApp.ElementType.LIST_ITEM) return el.asListItem().getText();
  return "";
}

function isSalarySectionHeader_(text) {
  return /^(Refilling Salary|Moonlit Attendance)\s+from\s+\d{2}\/\d{2}\/\d{4}\s+to\s+\d{2}\/\d{2}\/\d{4}:?$/i.test(String(text || "").trim());
}

function extractLastNumber_(text) {
  var matches = String(text || "").match(/-?\d[\d,]*(?:\.\d+)?/g);
  if (!matches || !matches.length) return null;
  var raw = matches[matches.length - 1].replace(/,/g, "");
  var n = parseFloat(raw);
  return isNaN(n) ? null : n;
}

function roundHalfUp_(n) {
  return n >= 0 ? Math.floor(n + 0.5) : Math.ceil(n - 0.5);
}

function formatWholeWithCommas_(n) {
  return Number(n).toLocaleString("en-US", { maximumFractionDigits: 0 });
}
