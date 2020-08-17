function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E10').activate();
  spreadsheet.getRange('E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E13').activate()
  .setValue('GL Description');
  spreadsheet.getRange('E15').activate();
};

function UntitledMacro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E7').activate();
  spreadsheet.getActiveRangeList().setFontColor('#fbe4d5');
  spreadsheet.getRange('F6').activate();
  spreadsheet.getActiveRangeList().setBackground('#fbe4d5');
};

function UntitledMacro2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getActiveRange().getDataRegion().activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A1:H24').applyRowBanding(SpreadsheetApp.BandingTheme.ORANGE);
};