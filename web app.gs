function doGet(e) {
  var app = UiApp.createApplication();
  var mode = e.parameter.mode;
  var owner = Session.getEffectiveUser().getEmail().toLowerCase();
  var user = Session.getActiveUser().getEmail().toLowerCase();
  var sheetName = e.parameter.activeSheetName;
  var ssId = e.parameter.ssId;
  var message = "The Grouper web app is only designed to work from links provided within the spreadsheet context.";
  var ss = SpreadsheetApp.openById(ssId);
  if ((ssId)&&(!ss)) {
     message = "You do not have access to the spreadsheet for which this web app requires access.";
     return ContentService.createTextOutput(message);
  }
  var sheet = ss.getSheetByName(sheetName);
  var sheetEditors = sheet.getSheetProtection().getUsers();
  if ((ss)&&(sheetEditors.indexOf(user)==-1)) {
    message = "You do not have the right to manage this group. Contact " + owner + " if you believe this is a mistake.";
    return ContentService.createTextOutput(message);  
  }
  if (mode == 'update') {
    message = loadUpdates(sheetName, mode);
  }
  if (mode == 'refresh') {
    message = refreshGroupSheets(sheetName, mode);
  }
  return ContentService.createTextOutput(message);
}
