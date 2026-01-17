function onEdit(e) {
  var column = e.range.getColumn();
  var row = e.range.getRow();
  var cell = e.range.getA1Notation();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var anh = new A1NotationHelper(sheet);
  var DutyCol = anh.titleToColumnIndex("Duty");
  var CACol = anh.titleToColumnIndex("CA");
  var OtherCol = anh.titleToColumnIndex("Expenses");
  var TotalOtherCol = anh.titleToColumnIndex("TotalExpenses");
  var NetSalesCol = anh.titleToColumnIndex("Net T. Sales Calc");
    var SalesNahalinCol = anh.titleToColumnIndex("Sales Nahalin");

  var lugar = ["Labangon", "Talamban", "Kalimpyo", "Goldswan", "Moonlit"];
  try { handleGoldswanPopulate(e); } catch (_) {}
  if (
    (e.range.getColumn() == CACol || e.range.getColumn() == DutyCol) &&
    row > 2 &&
    lugar.includes(sheet.getSheetName())
  ) {
    addAttendanceToEmployees(3, 16, 2020);
  } else if (sheet.getSheetName() == "GenerateReport") {
    if (cell == "D2")
      genReport();
    else if (cell == "D6")
      genMultReport();
  } else if (sheet.getSheetName() == "GoldeGlo" && cell == "C1") {
    goldeGloExpenses();
  } else if (sheet.getSheetName() == "GoldeGlo" && column == 3 && row > 2) {
    goldeGloExpensesLineSep(row, column);
  } else if (sheet.getSheetName() == "Moonlit Macro") {
    if (cell == "G1") {
      moonlitMacroEraseCells();
    } else if (cell == "C1") {
      executeMoonlit();
    }
  } else if (
    e.range.getColumn() == OtherCol &&
    row > 2 &&
    lugar.includes(sheet.getSheetName())
  ) {
    //SpreadsheetApp.getUi().alert(" Value " + e.value); //very useful debugging onedit
    totalExp = e.value.combineExpenses();
    SpreadsheetApp.getActiveSheet()
      .getRange(row, TotalOtherCol)
      .setValue(totalExp);
    Logger.log(totalExp);
  } else if (
    sheet.getSheetName() === "DisplaySalary" &&
    e.range.getA1Notation() === "A2"
  ) {
    displaySalary();
  }
  else if (sheet.getSheetName() == "GoldeGlo" && column == SalesNahalinCol && row > 2) {
    goldeGloProcessSalesNahalin(row, column);
    }
}




