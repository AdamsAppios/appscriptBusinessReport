class AttendanceEmployee {
  constructor() {
    this.employeesJson = {"Ronel":0, "Richard":0, "Allan":0, "Romel":0, "Angelo":0, "Dominic":0, "Lito":0, "Dion":0,"Charee":0};
    this.locationsArrayAtt = ["Talamban", "Labangon", "Kalimpyo", "Goldswan", "Golde Glo"];
  }

  setUtangs () {
    let atttSheet = SpreadsheetApp.getActive().getSheetByName("Attendance") 
    var values = atttSheet.getDataRange().getValues();
    for (let empj in this.employeesJson) {
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] == empj) {
          atttSheet.getRange(i+1, 2).setValue(this.employeesJson[empj]);
        }
      }
    }
  } 

  displayReports() {
    let caByDateAndLoc = "";
    for (let h=0; h < this.locationsArrayAtt.length; h++) {
    let testSheetClass = SpreadsheetApp.getActive().getSheetByName(this.locationsArrayAtt[h]);
    var dateConverted = new Date();
    var month = parseInt(Utilities.formatDate(dateConverted, 'GMT+8', 'MM'));
    let day = parseInt(Utilities.formatDate(dateConverted, "GMT+8", "dd"));
    let year = parseInt(Utilities.formatDate(dateConverted, "GMT+8", "yyyy"));
    let daysByMonth = parseInt(getDaysInMonth(year, month));
    //get day
    let startDay, endDay;
    Logger.log(day);
    if ( day > 15 && day <= daysByMonth) {
      Logger.log(`Loop day 16 to day ${daysByMonth}`);
      startDay=16, endDay = daysByMonth;
    } else if ( day <= 15 && day >= 1) {
      Logger.log(`Loop day 1 to day 15`);
      startDay = 1, endDay = 15;
    }

    for(let i = startDay; i <= endDay; i++) {
      let dateString = new Date(year, month-1, i);
      let rowDate = findRowByDateCell(testSheetClass, dateString);
      let colDate = titleToColumnIndex(testSheetClass, "Expenses");
      let expeValue = testSheetClass.getRange(rowDate, colDate).getValue();
      let expArray = expeValue.split(";");
      for (let j = 0; j < expArray.length; j++) {
        if (expArray[j].trim().startsWith("CA")) {
          let expArrayjay = expArray[j];
          let eqSeparated =expArrayjay.split("=");
          this.employeesJson[eqSeparated[0].replace(/^\s*CA\s*/,'')] += parseInt(eqSeparated[1]);
          caByDateAndLoc += `${this.locationsArrayAtt[h]}, ${Utilities.formatDate(dateString, 'GMT+8', 'MM/dd/yyyy')}, ${eqSeparated[0]} = ${eqSeparated[1]}\n`;
        }
      }
      Logger.log(` ${dateString} -> Expenses ${expeValue}`);
    }
    Logger.log(this.employeesJson);
    }
    Logger.log(caByDateAndLoc);
    this.setUtangs();
  }
}


function attclasstest() {
  let attEmp = new AttendanceEmployee();
  console.log(attEmp.displayReports());
}
