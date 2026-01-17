function isNaNzero(str) {
  if (typeof str === "string" && str == "") {
    return 0;
  } else {
    return parseFloat(str);
  }
}
//class experiment
function ReportClass(sLoc, dateEnt) {
  this.dealPrice = 10;
  this.pickPrice = 15;
  this.rndPrice = 150;
  this.sLoc = sLoc;
  //report values available for all three classes
  this.anh = new A1NotationHelper(sLoc);
  this.row = this.anh.detectRowByCurSheetDay(dateEnt); 

  this.dealer = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Dealer")).getValue()
  );
  this.lessplus = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Less")).getValue()
  );
  this.pickup = parseInt(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Pickup")).getValue()
  );
  this.round = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Rnd")).getValue()
  );
  this.totalRefilled = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Gal Cons")).getValue()
  );
  this.badgerMeter = isNaNzero(
    sLoc
      .getRange(this.row, this.anh.titleToColumnIndex("Badger Meter"))
      .getValue()
  );
  this.badgerPrev = isNaNzero(
    sLoc
      .getRange(this.row - 1, this.anh.titleToColumnIndex("Badger Meter"))
      .getValue()
  );
  this.badgerMeterConsump = (this.badgerMeter - this.badgerPrev) * 2;
  this.overShortBadger = this.badgerMeterConsump - this.totalRefilled;
  this.OverCTO = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Over CTO")).getValue()
  );
  this.expenses = sLoc.getRange(this.row, this.anh.titleToColumnIndex("Expenses")).getValue();
  this.totalexpenses = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("TotalExpenses")).getValue());
  this.ctoReport = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("CTO Text")).getValue());
  this.ctoCalc = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("CTO Calc")).getValue());
  this.rndBeg = isNaNzero(sLoc.getRange(this.row-1, this.anh.titleToColumnIndex("Rnd End")).getValue());
  this.rndEnd = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("Rnd End")).getValue());
  this.capsealBeg = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("Cap seal beg rep")).getValue());
  this.capsealEnd = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("Cap seal end rep")).getValue());
  this.capsealOver = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("over Capseal")).getValue());

  this.ngDuty = sLoc.getRange(this.row, this.anh.titleToColumnIndex("Duty")).getValue()
  this.NoteString = sLoc.getRange(this.row, this.anh.titleToColumnIndex("Notes")).getValue()

}
function KalimpReportClass(sLoc, dateEnt) {
  ReportClass.call(this, sLoc, dateEnt);
  this.squarePrice = 15;
  this.smallPrice = 8;
  this.smallSqPrice = 10;
  this.bakeryPrice = 15;
  this.tenlitersPrice = 10;
  this.noSealPrice = 15;

  //get the values of the seal and 
  this.noSeal = isNaNzero (sLoc.getRange(this.row, this.anh.titleToColumnIndex("No Seal")).getValue());
  this.tenliters = parseInt(isNaNzero (sLoc.getRange(this.row, this.anh.titleToColumnIndex("10 Liters")).getValue()));
  this.small = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Small")).getValue()
  );
  this.squareSmall = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("Small Sq.")).getValue())
  this.square = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Square S.")).getValue()
  );
  this.squareSealBeg = isNaNzero(
    sLoc.getRange(this.row-1, this.anh.titleToColumnIndex("Sq End")).getValue()
  );
  this.squareSealEnd = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Sq End")).getValue()
  );
  this.vecoReading = sLoc.getRange(this.row, this.anh.titleToColumnIndex("Veco Reading")).getValue();
  //sales
  this.salesDeal = (this.dealer) * this.dealPrice;
  this.salesPickp = this.pickup * this.pickPrice;
  this.salesRnd = this.round * this.rndPrice;
  this.salesSG = this.small * this.smallPrice;
  this.salesNoSeal = this.noSeal * this.squarePrice;
  this.salesTenLiters = this.tenliters * this.tenlitersPrice;
  this.salesSquare = this.square * this.squarePrice;
  this.salesSquareSmall = this.squareSmall* this.squareSmallPrice;
  this.salesTA = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("T Sales Text")).getValue());

  //Calculating capseal Endings
  this.capsealEndPrev = isNaNzero(
    sLoc.getRange(this.row-1, this.anh.titleToColumnIndex("Cap seal end rep")).getValue()
  );
  this.capsealBeg = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Cap seal beg rep")).getValue()
  );
  this.capsealEnd = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Cap seal end rep")).getValue()
  );
  this.overcapseal = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("over Capseal")).getValue()
  );

  //Display as string:
  this.noteString = (this.NoteString != "") ? `Note: ${this.NoteString}\n\n` : `\n\n`;
  this.dealString = (this.dealer > 0) ? `D${this.dealer}`: ``;
  this.lessplusString = this.lessplus > 0 ? `- less ${this.lessplus} = ${this.dealer-this.lessplus}` : ``; 
  this.pickString = `P${this.pickup}`;
  this.smallString = (this.small > 0) ? `+ sg ${this.small} `: ``;
  this.squareSmallString = (this.squareSmall > 0) ? `+ (SqSmall) ${this.squareSmall} `: ``;
  this.noSealString = (this.noSeal > 0)?`+ (Noseal) ${this.noSeal} `:"";
  this.tenLitersString = (this.tenliters > 0)?`+ (10L) ${this.tenliters} `:"";
  this.rndString = (this.round > 0) ? `+ ${this.round} Rnd`: ``;
  this.squareString = (this.square > 0) ? `+ Sq ${this.square} `: ``;
  this.salesNoSealString = (this.noSeal > 0) ? `+${this.salesNoSeal}` : ``;
  this.salesTenLitersString = (this.salesTenLiters > 0)?`+ ${this.salesTenLiters}`:"";
  this.salesDealString = (this.dealer > 0) ? `${this.salesDeal}/` : ``;
  this.salesRndString = (this.round > 0) ? `+ ${this.salesRnd}` : ``;
  this.salesSGString = (this.small > 0 ) ? `/ ${this.salesSG}` : ``;
  this.salesSquareString = (this.square > 0 )  ?`/ ${this.salesSquare}`:"";
  this.salesSquareSmallString = (this.squareSmall > 0)?`/ ${this.squareSmallString}`:"";
  this.expensesString = (this.expenses != "") ? `- Expenses (${this.expenses})`: ``;
  this.overShortCTOString = (this.OverCTO < 0) ? `Short by ${Math.abs(this.OverCTO)}` : `Over by ${this.OverCTO}`;
  this.overShortRoundString = ((this.rndBeg-this.rndEnd) > this.round) ? `Short by ${(this.rndBeg-this.rndEnd-this.round)} containers` : ``;
  this.overShortCapsealString = (this.capsealOver < 0 ) ? ` Short by ${Math.abs(this.capsealOver)}` : ` Over by ${this.capsealOver}`;
  this.overShortBadgerMeterString = (this.overShortBadger > 0) ? ` Short by ${Math.abs(this.overShortBadger)}` : ` Over by ${Math.abs(this.overShortBadger)}`;


  
}

function LabReportClass(sLoc, dateEnt) {
  ReportClass.call(this, sLoc, dateEnt);
  this.squarePrice = 15;
  this.smallPrice = 5;
  this.squareSmallPrice = 10;
  this.noSeal = isNaNzero (sLoc.getRange(this.row, this.anh.titleToColumnIndex("No Seal")).getValue());
  this.tenliters = parseInt(isNaNzero (sLoc.getRange(this.row, this.anh.titleToColumnIndex("10 Liters")).getValue()));
  this.small = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Small")).getValue()
  );
  this.squareSmall = isNaNzero(sLoc.getRange(this.row, this.anh.titleToColumnIndex("Square Sm")).getValue())
  this.square = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Square")).getValue()
  );
  this.squareSealBeg = isNaNzero(
    sLoc.getRange(this.row-1, this.anh.titleToColumnIndex("Sq End")).getValue()
  );
  this.squareSealEnd = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Sq End")).getValue()
  );
  //sales
  this.salesDeal = (this.dealer) * this.dealPrice;
  this.salesPickp = this.pickup * this.pickPrice;
  this.salesRnd = this.round * this.rndPrice;
  this.salesSG = this.small * this.smallPrice;
  this.salesNoSeal = this.noSeal * this.squarePrice;
  this.salesTenLiters = this.tenliters * this.squareSmallPrice;
  this.salesSquare = this.square * this.squarePrice;
  this.salesSquareSmall = this.squareSmall* this.squareSmallPrice;
  this.salesTA = this.salesDeal + this.salesPickp + this.salesRnd + this.salesSG + this.salesSquareSmall+this.salesNoSeal+this.salesTenLiters+this.salesSquare;

}

function TmbReportClass(sLoc, dateEnt) {

  ReportClass.call(this, sLoc, dateEnt);
  this.dealPrice = 10;
  this.pickPrice = 15;
  this.salesDeal = (this.dealer-this.lessplus) * this.dealPrice;
  this.salesPickp = this.pickup * this.pickPrice;
  this.salesRnd = this.round * this.rndPrice;
  this.salesTA = this.salesDeal + this.salesPickp + this.salesRnd;
  this.small = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Small")).getValue()
  );
  this.agt = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Manong")).getValue()
  );
}

function GSReportClass(sLoc, dateEnt) {
  ReportClass.call(this, sLoc, dateEnt);
  this.pickPrice = 8;
  this.rndPrice = 155;
  this.rectPrice = 175;
  this.dealer = null;
  this.rect = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Rect")).getValue()
  );
  this.collectibles = isNaNzero(
    sLoc
      .getRange(this.row, this.anh.titleToColumnIndex("Collectibles"))
      .getValue()
  );
  this.collectiblesPrev = isNaNzero(
    sLoc
      .getRange(this.row - 1, this.anh.titleToColumnIndex("Collectibles"))
      .getValue()
  );
  this.utangpnp = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("utang pnp")).getValue()
  );
  this.paid = isNaNzero(
    sLoc.getRange(this.row, this.anh.titleToColumnIndex("Paid")).getValue()
  );
}
//inheritance
KalimpReportClass.prototype = Object.create(ReportClass.prototype);
LabReportClass.prototype = Object.create(ReportClass.prototype);
TmbReportClass.prototype = Object.create(ReportClass.prototype);
GSReportClass.prototype = Object.create(ReportClass.prototype);

//Kalimp For Report
KalimpReportClass.prototype.reportDispString = function (dateEnt) {
  let reportString = `Date: ${Utilities.formatDate(
  dateEnt,
  "GMT+8",
  "MMM dd, yyyy . EEE."
  )}:, Duty: ${this.ngDuty} \n`;

  //calculate the capsealstring added based on end

  this.capsealString = "Cap seal beg ";
  if (this.capSealEndPrev != this.capsealBeg) {
    this.capsealString += ` ${this.capsealEndPrev} in ${this.capsealBeg - this.capsealEndPrev} Cap seal End  ${this.capsealEnd}`;
  } else {
    this.capsealString += ` ${this.capsealEndPrev} Cap seal End  ${this.capsealEnd}`;
  }
  //calculate if capseal is short
  this.capsealString += this.overcapseal < 0 ? ` Capseal short by ${this.overcapseal}` : `Capseal over by ${this.overcapseal}`;

  //display strings from parent class
  reportString += `${this.dealString}${this.pickString}${this.squareString}${this.squareSmallString}${this.smallString}${this.noSealString}${this.tenLitersString}${this.rndString}\n`;
  reportString += `=${this.salesDealString}${this.salesPickp}${this.salesSquareString}${this.salesSquareSmallString}${this.salesSGString}${this.salesNoSealString}${this.salesTenLitersString}${this.salesRndString}\n`;
  reportString += `= TA ${this.salesTA}${this.expensesString}\n`;
  reportString += `= ${this.ctoCalc} CTO ${this.ctoReport} ${this.overShortCTOString}\n`;
  reportString += `${this.capsealString}\n`
  reportString += `Square Beg : ${this.squareSealBeg }  End : ${this.squareSealEnd} = ${this.squareSealBeg - this.squareSealEnd}\n`;
  reportString += `Rnd Beg : ${this.rndBeg} Rnd End : ${this.rndEnd} = ${this.rndBeg-this.rndEnd} ${this.overShortRoundString}\n`;
  reportString += `Badger Meter: ${this.badgerMeter} , Meter Consumo: ${this.badgerMeterConsump}${this.overShortBadgerMeterString}\n`;
  reportString += `VECO Reading: ${this.vecoReading}`;
  reportString += `${this.noteString}`;
  return reportString;

}

LabReportClass.prototype.reportDispString = function (dateEnt) {

  let reportString = ""
  let noteString = (this.NoteString != "") ? `Note: ${this.NoteString}\n\n` : `\n\n`;
  let dealString = (this.dealer > 0) ? `D${this.dealer}`: ``;
  let pickString = `P${this.pickup}`;
  let rndString = (this.round > 0) ? `+ ${this.round} Rnd`: ``;
  let noSealString = (this.noSeal > 0)?`+ (Noseal) ${this.noSeal} `:"";
  let salesNoSealString = (this.noSeal > 0) ? `+${this.salesNoSeal}` : ``;
  let tenLitersString = (this.tenliters > 0)?`+ (10L) ${this.tenliters} `:"";
  let salesTenLitersString = (this.salesTenLiters > 0)?`+ ${this.salesTenLiters}`:"";
  let squareString = (this.square > 0) ? `+ Sq ${this.square} `: ``;
  let smallString = (this.small > 0) ? `+ sg ${this.small} `: ``;
  let squareSmallString = (this.squareSmall > 0) ? `+ (SqSmall) ${this.squareSmall} `: ``;
  let salesDealString = (this.dealer > 0) ? `${this.salesDeal}/` : ``;
  let salesRndString = (this.round > 0) ? `+ ${this.salesRnd}` : ``;
  let salesSGString = (this.small > 0 ) ? `/ ${this.salesSG}` : ``;
  let salesSquareString = (this.square > 0 )  ?`/ ${this.salesSquare}`:"";
  let salesSquareSmallString = (this.squareSmall > 0)?`/ ${this.salesSquareSmall}`:"";
  let expensesString = (this.expenses != "") ? `- Expenses (${this.expenses})`: ``;
  let overShortCTOString = (this.OverCTO < 0) ? `Short by ${Math.abs(this.OverCTO)}` : `Over by ${this.OverCTO}`;
  let overShortRoundString = ((this.rndBeg-this.rndEnd) > this.round) ? `Short by ${(this.rndBeg-this.rndEnd-this.round)} containers` : ``;
  let overShortCapsealString = (this.capsealOver < 0 ) ? ` Short by ${Math.abs(this.capsealOver)}` : ` Over by ${this.capsealOver}`;
  let overShortBadgerMeterString = (this.overShortBadger > 0) ? ` Short by ${Math.abs(this.overShortBadger)}` : ` Over by ${Math.abs(this.overShortBadger)}`;
  reportString += `Date: ${Utilities.formatDate(
      dateEnt,
      "GMT+8",
      "MMM dd, yyyy . EEE."
    )}:, Duty: ${this.ngDuty} \n`;
  reportString += `${dealString}${pickString}${squareString}${squareSmallString}${smallString}${noSealString}${tenLitersString}${rndString}\n`;
  reportString += `=${salesDealString}${this.salesPickp}${salesSquareString}${salesSquareSmallString}${salesSGString}${salesNoSealString}${salesTenLitersString}${salesRndString}\n`;
  reportString += `= TA ${this.salesTA}${expensesString}\n`;
  reportString += `= ${this.ctoCalc} CTO ${this.ctoReport} ${overShortCTOString}\n`;
  reportString += `Capseal Beg : ${this.capsealBeg} Capseal End : ${this.capsealEnd} = ${this.capsealBeg-this.capsealEnd} ${overShortCapsealString}\n`
  reportString += `Square Beg : ${this.squareSealBeg }  End : ${this.squareSealEnd} = ${this.squareSealBeg - this.squareSealEnd}\n`;
  reportString += `Rnd Beg : ${this.rndBeg} Rnd End : ${this.rndEnd} = ${this.rndBeg-this.rndEnd} ${overShortRoundString}\n`;
  reportString += `Badger Meter: ${this.badgerMeter} , Meter Consumo: ${this.badgerMeterConsump}${overShortBadgerMeterString}\n`;
  reportString += `${noteString}`;
  return reportString; 
}

//Talamban For Report
TmbReportClass.prototype.reportDispString = function(dateEnt) {
  let reportString = ""
  let dealString = `D${this.dealer}`;
  let lessplusString = this.lessplus > 0 ? `- less ${this.lessplus} = ${this.dealer-this.lessplus}` : ``; 
  let pickString = `P${this.pickup}`;
  let rndString = (this.round > 0) ? `+ ${this.round} Rnd`: ``;
  let salesRndString = (this.round > 0) ? `+ ${this.salesRnd}` : ``;
  let expensesString = (this.expenses != "") ? `- Expenses (${this.expenses})`: ``;
  let overShortCTOString = (this.OverCTO < 0) ? `Short by ${Math.abs(this.OverCTO)}` : `Over by ${this.OverCTO}`;
  let overShortRoundString = ((this.rndBeg-this.rndEnd) > this.round) ? `Short by ${(this.rndBeg-this.rndEnd-this.round)} containers` : ``;
  let overShortCapsealString = (this.capsealOver < 0 ) ? ` Short by ${Math.abs(this.capsealOver)}` : ` Over by ${this.capsealOver}`;

  let overShortBadgerMeterString = (this.overShortBadger > 0) ? ` Short by ${Math.abs(this.overShortBadger)}` : ` Over by ${Math.abs(this.overShortBadger)}`;
  let noteString = (this.NoteString != "") ? `Note: ${this.NoteString}` : ``;
  reportString += `Date: ${Utilities.formatDate(
    dateEnt,
    "GMT+8",
    "MMM dd, yyyy . EEE."
  )}:, Duty: ${this.ngDuty} \n`;
  reportString += `${dealString}${lessplusString} ${pickString}${rndString}\n`;
  reportString += `=${this.salesDeal}/${this.salesPickp} ${salesRndString}\n`;
  reportString += `= TA ${this.salesTA}${expensesString}\n`;
  reportString += `= ${this.ctoCalc} CTO ${this.ctoReport} ${overShortCTOString}\n`;
  reportString += `Capseal Beg : ${this.capsealBeg} Capseal End : ${this.capsealEnd} = ${this.capsealBeg-this.capsealEnd} ${overShortCapsealString}\n`
  reportString += `Rnd Beg : ${this.rndBeg} Rnd End : ${this.rndEnd} = ${this.rndBeg-this.rndEnd} ${overShortRoundString}\n`;  
  reportString += `Badger Meter: ${this.badgerMeter} , Meter Consumo: ${this.badgerMeterConsump}${overShortBadgerMeterString} \n`;

  reportString += `${noteString}\n\n`;
  return reportString; 
}

function genMultReport() {
  let locCell = "C3",
    begDateCell = "B6",
    endDateCell = "C6",
    reportCell = 7; //A7 Sugod
  var attendS =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GenerateReport");
  let a = new Date(attendS.getRange(begDateCell).getValue());
  let b = attendS.getRange(endDateCell).getValue() == ""
      ? a
      : new Date(attendS.getRange(endDateCell).getValue());
  var datesArr = createDateSpan(a, b);
  for (let i = 0; i < datesArr.length; i++) {
    let reportString = "";
    var sLoc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      attendS.getRange(locCell).getValue()
    );
    if (sLoc.getSheetName() == "Talamban") {
      let tmbClass = new TmbReportClass(sLoc, datesArr[i]);
      reportString = tmbClass.reportDispString(datesArr[i]);
      attendS.getRange(`A${reportCell++}`).setValue(reportString);
    } else if(sLoc.getSheetName() == "Labangon") {
      let labClass = new LabReportClass(sLoc, datesArr[i]);
      reportString = labClass.reportDispString(datesArr[i]);
      attendS.getRange(`A${reportCell++}`).setValue(reportString);
    } else if(sLoc.getSheetName() == "Kalimpyo") {
      let kmpClass = new KalimpReportClass(sLoc, datesArr[i]);
      reportString = kmpClass.reportDispString(datesArr[i]);
      attendS.getRange(`A${reportCell++}`).setValue(reportString);
    }
    if (sLoc === null) {
      return;
    }
  }
}



