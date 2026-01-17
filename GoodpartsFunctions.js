//hello my name is trevor coreyson
Function.prototype.method = function (name, func) {
  this.prototype[name] = func;
  return this;
};

//Convert date range to array
var DAY_MILLIS = 24 * 60 * 60 * 1000;

function exampleDateRange() {
  var a = new Date("2016/01/01");
  var b = new Date("2019/01/03");
  var datesArr = createDateSpan(a, b);
  Logger.log(datesArr.length);
}

function createDateSpan(startDate, endDate) {
  if (startDate.getTime() > endDate.getTime()) {
    throw Error("Start is later than end");
  }

  var dates = [];

  var curDate = new Date(startDate.getTime());
  while (!dateCompare(curDate, endDate)) {
    dates.push(curDate);
    curDate = new Date(curDate.getTime() + DAY_MILLIS);
  }
  dates.push(endDate);
  return dates;
}

function dateCompare(a, b) {
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}
//end if

if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, "includes", {
    value: function (searchElement, fromIndex) {
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      // 1. Let O be ? ToObject(this value).
      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If len is 0, return false.
      if (len === 0) {
        return false;
      }

      // 4. Let n be ? ToInteger(fromIndex).
      //    (If fromIndex is undefined, this step produces the value 0.)
      var n = fromIndex | 0;

      var k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);

      function sameValueZero(x, y) {
        return (
          x === y ||
          (typeof x === "number" &&
            typeof y === "number" &&
            isNaN(x) &&
            isNaN(y))
        );
      }

      // 7. Repeat, while k < len
      while (k < len) {
        // a. Let elementK be the result of ? Get(O, ! ToString(k)).
        // b. If SameValueZero(searchElement, elementK) is true, return true.
        if (sameValueZero(o[k], searchElement)) {
          return true;
        }
        // c. Increase k by 1.
        k++;
      }

      // 8. Return false
      return false;
    },
  });
}

String.method("combineExpenses", function () {
  arr1 = this.split(";");
  totalExp = 0;
  for (i = 0; i < arr1.length; i++) {
    totalExp += parseInt(arr1[i].split("=")[1]);
  }
  return totalExp;
});

String.method("sumCellCA", function () {
  return this.split(",").sumCA();
});
function sumCellCA(cell) {
  return cell.split(",").sumCA();
}

Number.method("blankIfZero", function (outputString) {
  if (this > 0) {
    return outputString;
  } else {
    return "";
  }
});

Date.method("daysInMonthNow", function () {
  var monthNow = parseInt((this.getMonth() + 1).toString());
  var yearNow = parseInt(this.getFullYear().toString());
  return new Date(yearNow, monthNow, 0).getDate();
});

if (!String.prototype.endsWith) {
  String.prototype.endsWith = function (search, this_len) {
    if (this_len === undefined || this_len > this.length) {
      this_len = this.length;
    }
    return this.substring(this_len - search.length, this_len) === search;
  };
}

function A1NotationHelper(sheet) {
  this.sheet = sheet;
}

A1NotationHelper.method("CountColA", function () {
  var data = this.sheet.getDataRange().getValues();
  for (var i = data[1].length - 1; i >= 0; i--) {
    if (data[1][i] != null && data[1][i] != "") {
      return i + 1;
    }
  }
});

A1NotationHelper.method("columnToLetter", function (column) {
  var temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
});

//Converts title to column
A1NotationHelper.method("titleToColumnIndex", function (findString) {
  var data = this.sheet.getDataRange().getValues();
  for (var i = 0; data[1].length >= i; i++) {
    if (data[1][i] == findString) {
      return i + 1;
    }
  }
  return -1;
});

//Return row based on cell value 
//EFFECTIVE !!
function findRowByDateCell(sheetClass, dateCell) {
  var values = sheetClass.getDataRange().getValues();
  dateCell = Utilities.formatDate(dateCell, "GMT+8", "MM-dd-yyyy");

  for (var i = 2; i < values.length; i++) {
    var row = "";
    /*SpreadsheetApp.getUi().alert(
      Utilities.formatDate(values[i][0], "GMT+8", "MM-dd-yyyy") +
        " and " +
        dateCell
    );*/
    if (Utilities.formatDate(values[i][0], "GMT+8", "MM-dd-yyyy") == dateCell) {
      row = values[i][0];

      return i + 1;
    }
  }
}

//converts title to column (function version)
function titleToColumnIndex (sheetName , findString) {
  var data = sheetName.getDataRange().getValues();
  for (var i = 0; data[1].length >= i; i++) {
    if (data[1][i] == findString) {
      return i + 1;
    }
  }
  return -1;
};

A1NotationHelper.method("detectRowByCurSheetDay", function (dateEntered) {
  var dataInSheet = this.sheet.getDataRange().getValues();
  var dateEntered = new Date(dateEntered);
  var monthEnt = Utilities.formatDate(dateEntered, "GMT+8", "MM");
  var dayEnt = Utilities.formatDate(dateEntered, "GMT+8", "d");
  let yearEnt = Utilities.formatDate(dateEntered, "GMT+8", "yyyy");
  var RowStarting = 0;
  for (var i = 0; i < dataInSheet.length; i++) {
    var dateCell = new Date(dataInSheet[i][0]); //[0] because column A is the date column
    if (
      Utilities.formatDate(dateCell, "GMT+8", "yyyy") == yearEnt &&
      Utilities.formatDate(dateCell, "GMT+8", "MM") == monthEnt &&
      Utilities.formatDate(dateCell, "GMT+8", "d") == dayEnt
    ) {
      RowStarting = i + 1;
      break;
    }
  }
  return RowStarting;
});

Array.method("sumCA", function () {
  var sum = 0;
  for (var i = 0; i < this.length; i++) {
    sum += parseInt(this[i].replace(/[^\d.-]/g, ""));
  }
  return sum;
});

