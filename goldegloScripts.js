var customSort = function (a, b) {
  return Number(b.match(/=(\d+)/)[1]) - Number(a.match(/=(\d+)/)[1]);
};

function goldeGloExpensesLineSep(row, column) {
  let golde = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GoldeGlo");
  repeatFilter = (row, column) => {
    filteredExpenses = golde
      .getRange(row, column)
      .getValue()
      .replace(/^Gud evening nong Expenses Jedamsa .* \:/, "")
      .replace(/Total expenses nong .*\d$/, "")
      .replace(/Total Expenses nong: (\d+\.\d+)/,"")
      .replace(/=\s*/g, "=")
      .replace(/-\s*/g, "=")
      .replace(/,/g, "")
      .replace(/\n/g, " ")
      .trim();
    golde.getRange(row, column).setValue(filteredExpenses);
  };
  repeatFilter(row, column);
  let anh = new A1NotationHelper(golde);
  let addStrnColumn = anh.titleToColumnIndex("Display addString");
  let expSortedColumn = anh.titleToColumnIndex("Sorted");
  let sumExpColumn = anh.titleToColumnIndex("Calc Expenses");

  let str = golde.getRange(row, 3).getValue();

  let expArray = str.match(/[A-Za-z0-9\s\.]*\=\d*\.?\d*/g);
  let sumInside = 0;
  let addString = "";
  expArray.map(function (x) {
    addString += "+" + x.match(/\=(\d*\.?\d*)/)[1];
    sumInside += parseFloat(x.match(/\=(\d*\.?\d*)/)[1]);
  });
  expArray = expArray.sort(customSort).join(", ");
  addString = addString.substring(1, addString.length); //Remove first plus
  golde.getRange(row, expSortedColumn).setValue(expArray);
  golde.getRange(row, sumExpColumn).setValue(sumInside);
  golde.getRange(row, addStrnColumn).setValue(addString);
}

function goldeGloExpenses() {
  var dateCell = "B1",
    expensesCol = 3,
    sortCellCol = 4,
    sumExpCol = 7,
    addStringCol = 5;
  var dateEnd = "D1";
  var golde = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GoldeGlo");
  var a = new Date(golde.getRange(dateCell).getValue());
  var b =
    golde.getRange(dateEnd).getValue() == ""
      ? a
      : new Date(golde.getRange(dateEnd).getValue());
  var datesArr = createDateSpan(a, b);
  var addString = "";
  for (var i = 0; i < datesArr.length; i++) {
    var anh = new A1NotationHelper(golde);
    var row = findRowByDateCell(golde, datesArr[i]);
    //SpreadsheetApp.getUi().alert(row); //very useful debugging onedit

    var str = golde.getRange(row, 3).getValue();

    var expArray = str.match(/[A-Za-z0-9\s\.]*\=\d*\.?\d*/g);
    var sumInside = 0;

    expArray.map(function (x) {
      addString += "+" + x.match(/\=(\d*\.?\d*)/)[1];
      sumInside += parseFloat(x.match(/\=(\d*\.?\d*)/)[1]);
    });
    expArray = expArray.sort(customSort).join(", ");
    addString = addString.substring(1, addString.length); //Remove first plus
    golde.getRange(row, sortCellCol).setValue(expArray);
    golde.getRange(row, sumExpCol).setValue(sumInside);
    golde.getRange(row, addStringCol).setValue(addString);
  }
}

/**
 * Processes a multi-line paste into the "Sales Nahalin" column on GoldeGlo.
 * - Extracts "Sales=####" → writes to "Golde Day"
 * - Removes the date line (anything up to the first ":" on line 1) and the "Sales=####" line
 * - Normalizes remaining "label=value" pairs and writes them back as a comma-separated list
 * - Sums all values and writes: Sales Pan = Golde Day + Calc Expenses - sum(items)
 * 
 * To avoid re-trigger loops, Sales Pan is only written when a Sales number is present.
 */
function goldeGloProcessSalesNahalin(row, col) {
  var sh  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GoldeGlo");
  var anh = new A1NotationHelper(sh);

  var salesNahalinCol = anh.titleToColumnIndex("Sales Nahalin");
  var goldeDayCol     = anh.titleToColumnIndex("Golde Day");
  var salesPanCol     = anh.titleToColumnIndex("Sales Pan");
  var calcExpCol      = anh.titleToColumnIndex("Calc Expenses");

  if (col !== salesNahalinCol) return;

  var raw = String(sh.getRange(row, col).getValue() || "").replace(/\r/g, "");

  // --- Extract Sales number (allows commas/decimals)
  var salesMatch = raw.match(/^\s*Sales\s*=\s*([\d,\.]+)/im);
  var hasSales   = !!salesMatch;
  var sales      = hasSales ? Number(String(salesMatch[1]).replace(/,/g, "")) : 0;

  // Write Golde Day only when Sales is present
  if (hasSales && goldeDayCol) {
    sh.getRange(row, goldeDayCol).setValue(sales);
  }

  // --- Strip the leading date line ending with ":" and the Sales line
  var cleaned = raw
    .replace(/^[^\n:]*:\s*\n?/, "")        // remove first line up to ":"
    .replace(/^\s*Sales\s*=\s*.*\n?/im, ""); // remove the "Sales=####" line

  // --- Normalize separators, then parse "label=value" pairs
  cleaned = cleaned
    .replace(/[–-]\s*/g, "=")   // dash to '='
    .replace(/\s*=\s*/g, "=")   // tighten equals
    .replace(/,\s*/g, "\n");    // commas → line breaks for uniform splitting

  var lines = cleaned.split(/\n+/).map(s => s.trim()).filter(Boolean);

  var items    = [];
  var sumItems = 0;

  lines.forEach(line => {
    var m = line.match(/^(.+?)=(\d+(?:\.\d+)?)/);
    if (!m) return;
    var label = m[1].replace(/\s+/g, " ").trim();
    var value = Number(m[2]);
    items.push(label + "=" + value);
    sumItems += value;
  });

  // --- Write normalized CSV back to "Sales Nahalin"
  var normalized = items.join(", ");
  sh.getRange(row, salesNahalinCol).setValue(normalized);

  // --- Corrected formula:
  // Sales Pan = Golde Day + Calc Expenses - sum(items)
  if (hasSales && salesPanCol) {                   // only compute on the first edit (with "Sales=" present)
    var calcExp  = Number(sh.getRange(row, calcExpCol).getValue()) || 0;
    var salesPan = (sales || 0) + calcExp - (sumItems || 0);
    sh.getRange(row, salesPanCol).setValue(salesPan);
  }
}
