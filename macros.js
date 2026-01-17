function coloroftextincell() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E4').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('16 - overtime 3 Hours,17,18 - overtime 3 Hours,19 - overtime 3 Hours,20,21,22,24,25,26,27,28,29,')
  .setTextStyle(30, 46, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#ff0000')
  .build())
  .build());
  spreadsheet.getRange('E5').activate();
};