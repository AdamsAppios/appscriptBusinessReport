function displaySalary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName("DisplaySalary");
  var showAll = ds.getRange("A2").getValue();

  // ----- Detect which rows we previously filled in column A (so we only clear those in F/H)
  var lastRow = ds.getLastRow();
  var oldA = lastRow > 2 ? ds.getRange(3, 1, lastRow - 2, 1).getValues() : [];
  var rowsToClear = [];
  for (var i = 0; i < oldA.length; i++) {
    if (String(oldA[i][0] || "").trim() !== "") rowsToClear.push(3 + i); // absolute row numbers
  }

  // Clear ONLY the A rows we used before (safe to wipe fully), and ONLY the matching F/H rows.
  // Keep all other blank lines & user-entered values intact.
  if (oldA.length) {
    ds.getRange(3, 1, oldA.length, 1).clearContent(); // A3:A—OK to fully clear our previous output
  }
  var hdrRow = ds.getRange(1, 1, 1, ds.getLastColumn()).getValues()[0];
  var totalCol = 8; // default H
  for (var c = 0; c < hdrRow.length; c++) {
    var t = (hdrRow[c] || "").toString().toLowerCase();
    if (t && t.indexOf("total") === 0) { totalCol = c + 1; break; }
  }
  // Clear Totals (H by default) and CA log (now F) only where we previously printed a line
  rowsToClear.forEach(function(r){
    ds.getRange(r, 6).clearContent();         // F (CA log)
    ds.getRange(r, totalCol).clearContent();  // Total column
  });

  // If the toggle is off, we’re done after cleanup.
  if (!showAll) return;

  // ----- Read date windows
  var attRange = ds.getRange("B2:C2").getValues()[0]; // attendance & CA window
  var otRange  = ds.getRange("D2:E2").getValues()[0]; // OT/UT/HD window
  var attStart = attRange[0], attEnd = attRange[1];
  var otStart  = otRange[0],  otEnd  = otRange[1];
  var tz = ss.getSpreadsheetTimeZone() || "Asia/Manila";

  // ----- Accumulators
  var baseCount       = {};
  var contribList     = {};
  var contribFraction = {};
  var cashAdvTotal    = {};
  var caLogByEmp      = {};

  // ----- Scan both stores
  ["Talamban","Labangon"].forEach(function(sheetName) {
    var sh    = ss.getSheetByName(sheetName);
    var data  = sh.getDataRange().getValues();
    var hdr   = data[1];
    var iDate = hdr.indexOf("Date");
    var iExp  = hdr.indexOf("Expenses");
    var iDuty = hdr.indexOf("Duty");
    var maxHStore = (sheetName === "Talamban" ? 12 : 11.5);

    data.slice(2).forEach(function(r) {
      var d = r[iDate];
      if (!(d instanceof Date) || isNaN(d)) return;

      var inAtt = attStart && attEnd && d >= attStart && d <= attEnd;
      var inOT  = otStart  && otEnd  && d >= otStart  && d <= otEnd;

      // CA capture (attendance window only)
      if (inAtt) {
        var txt = (r[iExp] || "") + "";
        var reCA = /CA\s+([^=;]+)=\s*([\d\.]+)/gi, m;
        while ((m = reCA.exec(txt))) {
          var emp = m[1].trim();
          var amt = parseFloat(m[2]) || 0;
          cashAdvTotal[emp] = (cashAdvTotal[emp] || 0) + amt;
          (caLogByEmp[emp] = caLogByEmp[emp] || []).push({ d: d, amt: amt });
        }
      }

      // Duty parsing
      ("" + r[iDuty]).split(",").forEach(function(part){
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
            contribFraction[emp] = (contribFraction[emp] || 0) - (1 - (h2 / maxHStore));
          }
        }
      });
    });
  });

  // ----- Attendance sheet (rates & fixed add/deduct)
  var at  = ss.getSheetByName("Attendance");
  var A   = at.getDataRange().getValues();
  var Hh  = A[0], idx = col => Hh.indexOf(col);
  var cName = idx("Name"), cDaily = idx("Daily 1");
  var cAccts = idx("Accts"), cCharges = idx("Charges");
  var cShort = idx("Charge Short"), cSSS = idx("SSS"), cPhil = idx("Philhealth");
  var cExtra = idx("Extra based on gallons sold");

  function roundHalfUp(n){ return n >= 0 ? Math.floor(n + 0.5) : Math.ceil(n - 0.5); }
  function fmtDate(d){ return Utilities.formatDate(d, tz, "MM/dd/yyyy"); }

  // ----- Build and write outputs
  var aVals = [], fVals = [], hVals = [];

  Object.keys(baseCount).forEach(function(emp){
    var base = baseCount[emp] || 0;
    var frac = contribFraction[emp] || 0;
    var totalDays = parseFloat((base + frac).toFixed(2));

    // Line 1
    var line1 = emp + " calculation days : " + base;
    (contribList[emp] || []).forEach(function(tok){
      line1 += (tok.charAt(0) === "-" ? " " : " + ") + tok;
    });
    line1 += " = " + totalDays;

    // Attendance row
    var rIdx = A.findIndex(r => r[cName] === emp);
    var rate   = rIdx > 0 ? parseFloat(A[rIdx][cDaily])   || 0 : 0;
    var extraG = rIdx > 0 ? parseFloat(A[rIdx][cExtra])   || 0 : 0;
    var accts  = rIdx > 0 ? parseFloat(A[rIdx][cAccts])   || 0 : 0;
    var chgs   = rIdx > 0 ? parseFloat(A[rIdx][cCharges]) || 0 : 0;
    var shrt   = rIdx > 0 ? parseFloat(A[rIdx][cShort])   || 0 : 0;
    var sss    = rIdx > 0 ? parseFloat(A[rIdx][cSSS])     || 0 : 0;
    var phil   = rIdx > 0 ? parseFloat(A[rIdx][cPhil])    || 0 : 0;
    var caAmt  = cashAdvTotal[emp] || 0;

    var basePay = totalDays * rate;
    var foodPay = totalDays * 60;
    var gross   = basePay + foodPay;
    var net     = gross + extraG - accts - chgs - shrt - sss - phil - caAmt;

    // Line 2 with “3075+900 = …”
    var line2 = emp + ": " + totalDays + " days : "
      + totalDays + "×daily rate " + rate
      + " + " + totalDays + "×food allowance 60 = "
      + basePay.toFixed(0) + "+" + foodPay.toFixed(0)
      + " = " + gross.toFixed(2);
    if (extraG) line2 += " + Extra Based On Gallons " + extraG;
    if (caAmt)  line2 += " - Cash Advances "         + caAmt;
    if (chgs)   line2 += " - Charges "               + chgs;
    if (shrt)   line2 += " - Charge Short "          + shrt;
    if (sss)    line2 += " - SSS "                   + sss;
    if (phil)   line2 += " - Philhealth "            + phil;
    if (accts)  line2 += " - Accts "                 + accts;
    line2 += " = " + net.toFixed(2);

    aVals.push([ line1 + "\n" + line2 ]);

    // F-column: CA breakdown (attendance window)
    var caLogStr = "";
    if (caLogByEmp[emp] && caLogByEmp[emp].length) {
      caLogByEmp[emp].sort(function(a,b){ return a.d - b.d; });
      caLogStr = emp + " : " + caLogByEmp[emp]
        .map(function(e){ return fmtDate(e.d) + "=" + (e.amt || 0); })
        .join(" ; ");
    }
    fVals.push([ caLogStr ]);

    // Totals column (rounded half-up)
    hVals.push([ roundHalfUp(net) ]);
  });

  if (aVals.length) ds.getRange(3, 1, aVals.length, 1).setValues(aVals);     // A
  if (fVals.length) ds.getRange(3, 6, fVals.length, 1).setValues(fVals);     // F (moved from G)
  if (hVals.length) ds.getRange(3, totalCol, hVals.length, 1).setValues(hVals); // H by header
}
