function displaySalary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName("DisplaySalary");
  var showAll = ds.getRange("A2").getValue();
  ds.getRange("A3:A1000").clearContent();
  if (!showAll) return;

  // Read the two ranges
  var [attStart, attEnd] = ds.getRange("B2:C2").getValues()[0];
  var [otStart,  otEnd ] = ds.getRange("D2:E2").getValues()[0];

  // Accumulators
  var baseCount       = {};
  var contribList     = {};
  var contribFraction = {};
  var cashAdv         = {};

  ["Talamban","Labangon"].forEach(function(sheetName) {
    var sh    = ss.getSheetByName(sheetName);
    var data  = sh.getDataRange().getValues();
    var hdr   = data[1];
    var iDate = hdr.indexOf("Date");
    var iExp  = hdr.indexOf("Expenses");
    var iDuty = hdr.indexOf("Duty");
    var maxH  = sheetName==="Talamban"? 12 : 11.5;

    data.slice(2).forEach(function (r) {
      var d = r[iDate];
      if (!(d instanceof Date) || isNaN(d)) return;

      var inAtt = attStart && attEnd && d >= attStart && d <= attEnd;
      var inOT  = otStart  && otEnd  && d >= otStart  && d <= otEnd;

      // --- Cash Advances ONLY within attendance window (B2:C2)
      if (inAtt) {
        var txt = (r[iExp] || "") + "";
        var reCA = /CA\s+([^=;]+)=\s*([\d\.]+)/gi, m;
        while ((m = reCA.exec(txt))) {
          var emp = m[1].trim(), amt = parseFloat(m[2]) || 0;
          cashAdv[emp] = (cashAdv[emp] || 0) + amt;
        }
      }

      // --- Duties parsed once
      ("" + r[iDuty]).split(",").forEach(function (part) {
        var t = part.trim();
        if (!t) return;
        var name = t.replace(/\(.*\)/, "").trim();
        var maxH = (sheetName === "Talamban" ? 12 : 11.5);

        // Base day only inside attendance window
        if (inAtt) baseCount[name] = (baseCount[name] || 0) + 1;

        // OT/UT/HD only inside OT window (D2:E2)
        if (inOT) {
          if (/\(HD\)/i.test(t)) {
            contribList[name] = contribList[name] || [];
            contribList[name].push("- hd 0.5");
            contribFraction[name] = (contribFraction[name] || 0) - 0.5;
          }
          var mo = t.match(/\(OT:\s*([\d\.]+)\)/i);
          if (mo) {
            var h = parseFloat(mo[1]) || 0;
            contribList[name] = contribList[name] || [];
            contribList[name].push("ot " + h + "/" + maxH);
            contribFraction[name] = (contribFraction[name] || 0) + (h / maxH);
          }
          var mu = t.match(/\(UT:\s*([\d\.]+)\)/i);
          if (mu) {
            var h2 = parseFloat(mu[1]) || 0;
            contribList[name] = contribList[name] || [];
            contribList[name].push("- ut (1-" + h2 + "/" + maxH + ")");
            contribFraction[name] = (contribFraction[name] || 0) - (1 - (h2 / maxH));
          }
        }
      });
    });

  });

  // Pull Attendance data
  var at    = ss.getSheetByName("Attendance");
  var A     = at.getDataRange().getValues();
  var H     = A[0];
  var idx   = col=>H.indexOf(col);
  var cName = idx("Name"),   cDaily = idx("Daily 1");
  var cAccts = idx("Accts"), cCharges = idx("Charges");
  var cShort = idx("Charge Short"), cSSS = idx("SSS"), cPhil = idx("Philhealth");
  var cExtra = idx("Extra based on gallons sold");

  // Build output
  var out = [];
  Object.keys(baseCount).forEach(emp=>{
    var base = baseCount[emp]||0;
    var frac = contribFraction[emp]||0;
    var totalDays = parseFloat((base + frac).toFixed(2));

    // —— Line 1: smart “+” vs “-” formatting
    var line1 = emp + " calculation days : " + base;
    (contribList[emp]||[]).forEach(function(c){
      if (c.charAt(0)==="-") {
        line1 += " " + c;
      } else {
        line1 += " + " + c;
      }
    });
    line1 += " = " + totalDays;

    // —— Line 2: salary formula
    var r = A.findIndex(r=>r[cName]===emp);
    var rate   = r>0? parseFloat(A[r][cDaily]) || 0 : 0;
    var extraG = r>0? parseFloat(A[r][cExtra]) || 0 : 0;
    var accts  = r>0? parseFloat(A[r][cAccts]) || 0 : 0;
    var chgs   = r>0? parseFloat(A[r][cCharges])|| 0 : 0;
    var shrt   = r>0? parseFloat(A[r][cShort])   || 0 : 0;
    var sss    = r>0? parseFloat(A[r][cSSS])     || 0 : 0;
    var phil   = r>0? parseFloat(A[r][cPhil])    || 0 : 0;
    var caAmt  = cashAdv[emp]||0;

    var basePay = totalDays * rate;
    var foodPay = totalDays * 60;
    var gross   = basePay + foodPay;
    var net     = gross + extraG
                - accts - chgs - shrt - sss - phil - caAmt;

    var part1 = basePay.toFixed(0);      // e.g. "3075"
    var part2 = foodPay.toFixed(0);      // e.g. "900"
    var line2 = emp + ": " 
      + totalDays + " days : "
      + totalDays + "×daily rate " + rate
      + " + " + totalDays + "×food allowance 60 = "
      + part1 + "+" + part2 
      + " = " + gross.toFixed(2);

    if (extraG) line2 += " + Extra Based On Gallons " + extraG;
    if (caAmt)  line2 += " - Cash Advances "         + caAmt;
    if (chgs)   line2 += " - Charges "               + chgs;
    if (shrt)   line2 += " - Charge Short "          + shrt;
    if (sss)    line2 += " - SSS "                   + sss;
    if (phil)   line2 += " - Philhealth "            + phil;
    if (accts)  line2 += " - Accts "                 + accts;

    line2 += " = " + net.toFixed(2) + "\n";

    out.push([ line1 + "\n" + line2 ]);
  });

  if (out.length) ds.getRange(3,1,out.length,1).setValues(out);
}
