function detectDateByRow(sLoc, row) {
  let dateEnt = simplifyDate(convertDateToUTC(sLoc.getRange(row, 1).getValue(), 3))
  return dateEnt;
}

function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}


//To avoid dailight savings time (date is decremented one day before)
function convertDateToUTC(date, timezoneOffset) { 
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours() + timezoneOffset, date.getUTCMinutes(), date.getUTCSeconds()); 
}

function simplifyDate(dateEnt) {
  return Utilities.formatDate(dateEnt, `GMT +8:00`, "MMMM dd, yyyy")
}

function simplifyTime(timeEnt) { 
  return Utilities.formatDate(timeEnt, `GMT +8:00`, "hh a"); 
}

function showCurrentTime() {
  Logger.log(Utilities.formatDate(convertDateToUTC(new Date(),3), "GMT+0800", "yyyy-MM-dd HH:mm:ss"));
}

function getCurrentDateFunc() {
  return `Date: ${simplifyDate(convertDateToUTC(new Date(), 3))}`;
}

function greetFunc() {
  let curHour = parseInt(Utilities.formatDate(convertDateToUTC(new Date(),3), "GMT +8:00", "HH"));
  let strGreet = "";
  if (curHour >= 0 && curHour < 12 ) {
    strGreet = ` Good AM maam/sir `;
  } else if (curHour >= 12 && curHour < 24)
    strGreet = ` Good PM maam/sir `;
  return strGreet;
}

