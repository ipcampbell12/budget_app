//Global variables
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 


function initMenu(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Budget App");
  menu.addItem("Show Budget Form", "showUserForm");
 
  menu.addItem("Clear Logs","clearContents");
  menu.addItem("New Month","dupName");

  menu.addToUi();
}

function onOpen(e){

  initMenu();
}


function showUserForm() {
  var template = HtmlService.createTemplateFromFile("log");
  var html = template.evaluate(); 
  html.setTitle("Budget App");
  SpreadsheetApp.getUi().showSidebar(html);
}

function userClicked(expense,amount,category,type,formattedDate){

const monthNames = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

const d = new Date();
var month = monthNames[d.getMonth()]

  var ws = ss.getSheetByName("Comprehensive Expense Tracker");
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "MM/dd/yy");
  //apparently upper and lower case for the date format make a difference
  
 var data = [expense,amount,category,type,formattedDate,month];
 ws.getRange(3,2).setNumberFormat("$0.00" )
  ws.insertRowBefore(3).getRange(3,1,1,data.length).setValues([data]);
  //has to be an array 

}





function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent()

}

function clearContents(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var ws = ss.getSheetByName("Comprehensive Expense Tracker");
  ws.getRange(3,1,ws.getLastRow()-1,5).clearContent();
}



function dupName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var name = Browser.inputBox('Enter new sheet name');
  var newSheet = ss.insertSheet(name, {template: sheet}); 
  var template = ss.getSheetByName("Template");
  var formulas = template.getRange(3,4,template.getLastRow()-4,1).getFormulasR1C1();


  var expensesRange = newSheet.getRange(3,4,newSheet.getLastRow()-4); 
  var lastMonthRange = newSheet.getRange(3,2,newSheet.getLastRow()-2);
  var expenses = newSheet.getRange(3,4,newSheet.getLastRow()-2).getValues();
  lastMonthRange.setValues(expenses);
  expensesRange.setFormulasR1C1(formulas);
  newSheet.getRange("j11:j24").uncheck();
  newSheet.getRange("j2:j7").uncheck();

  var month = newSheet.getSheetName();
  newSheet.getRange("A1").setValue(month)

//Logger.log(month)
}



