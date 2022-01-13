function getDataFromGoogleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FormData");
  const [header, ...data] = sheet.getDataRange().getDisplayValues();
  const choices = {}
  header.forEach(function(title,index) {
    choices[title] = data.map(row => row[index]).filter(e  => e !="");
  });
  return choices;
}

function populateGoogleForm() {
 const GOOGLE_FORM_ID = "1UoAMMvtUuw9HWnshJktmpoTVAUeFEs04XjVxW1xmjYw";
  const googleForm = FormApp.openById(GOOGLE_FORM_ID);
  const items = googleForm.getItems();
  const choices = getDataFromGoogleSheets();
  items.forEach(function(item) {
    const itemTitle = item.getTitle();
    if(itemTitle in choices) {
      const itemType = item.getType();
      switch (itemType) {
        case FormApp.ItemType.CHECKBOX: 
        item.asCheckboxItem().setChoiceValues(choices[itemTitle]);
        break;
        case FormApp.ItemType.LIST: 
        item.asListItem().setChoiceValues(choices[itemTitle]);
        break;
        case FormApp.ItemType.MULTIPLE_CHOICE :
        item.asMultipleChoiceItem().setChoiceValues(choices[itemTitle]);
        break;
        default:
       
      }
    }
  })
}

function sortColumn() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName("CookieOrders");
 var range = sheet.getRange("A2:Z500");

 // Sorts by the values in the first column (A)
 range.sort(1);
}

function getDataFromCookieOrders(row) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CookieOrders");
var range = sheet.getRange(2,1,85,26);
var values = range.getValues();

return values[row];
}


function updateHistorical () {
const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("HistoricalCustomerData");
var range = sheet.getRange(2,1,85,26);
var emptyRow = 0;
  for(let i = 1; i<85;i++) {
sheet.appendRow(getDataFromCookieOrders(i));
console.log(getDataFromCookieOrders(i).join());
if(getDataFromCookieOrders(i).join() === ",,,,,,,,,,,,,,,,,,,,,,,,,"){
  emptyRow++;
}
if(emptyRow >= 4){
  break;
}
  }
deleteDuplicates();
}

function sortByDate(sheetName) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 var range = sheet.getRange("A2:Z500");

 // Sorts by the values in the first column (A)
 range.sort(1);
}

function deleteDuplicates() {
  var sheetName = "HistoricalCustomerData";
  sortByDate(sheetName);
  reuseDeleteDuplicates(sheetName);
  sortByDate(sheetName);
 } 


 function deleteQuestionNames() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName("CookieOrders");
 var range = sheet.getRange("A1:x500");

for (let i = 0; i<50; i++) {
  if(range.getValues()[i+1].join() === range.getValues()[0].join()) {
   sheet.getRange(i+2,1,1,26).setValues([["","","","","","","","","","","","","","","","","","","","","","","","","","",]]);
   break;
  }
}

  }


function saveWeekly() {
var ss1 = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = ss1.getSheetByName("CookieOrders");
var range1 = sheet1.getRange("F2:F500");
var thisWeek = SpreadsheetApp.getActiveSpreadsheet();
var sheet2 = thisWeek.getSheetByName("ThisWeek");
var range2 = sheet2.getRange("A2:z500");
var oneWeek = SpreadsheetApp.getActiveSpreadsheet();
var sheet3 = oneWeek.getSheetByName("OneWeek");
var range3 = sheet3.getRange("A2:z500");
var twoWeeks = SpreadsheetApp.getActiveSpreadsheet();
var sheet4 = twoWeeks.getSheetByName("TwoWeeks");
var range4 = sheet4.getRange("A2:z500");
var threeWeeks = SpreadsheetApp.getActiveSpreadsheet();
var sheet5 = threeWeeks.getSheetByName("ThreeWeeks");
var range5 = sheet5.getRange("A2:Z500");
var dates = SpreadsheetApp.getActiveSpreadsheet();
var sheet6 = dates.getSheetByName("Dates");
var sunday = Date.parse(sheet6.getRange("B1").getValues()[0]);
var nextSunday = Date.parse(sheet6.getRange("B2").getValues()[0]);
var twoSundays = Date.parse(sheet6.getRange("B3").getValues()[0]);
var threeSundays = Date.parse(sheet6.getRange("B4").getValues()[0]);

let referenceDate;
sortColumn();
for (let i = 0; i < 100; i++) {
referenceDate = Date.parse(range1.getValues()[i]);
Logger.log('referenceDate is ' + range1.getValues()[i]);
if(referenceDate >= threeSundays) {
  Logger.log('three weeks!');
  sheet5.appendRow(getDataFromCookieOrders(i));
}
else if(referenceDate >= twoSundays) {
  Logger.log('two weeks!');
  sheet4.appendRow(getDataFromCookieOrders(i));
}
else if(referenceDate >= nextSunday) {
  Logger.log('next week!');
  sheet3.appendRow(getDataFromCookieOrders(i));
}
else if(referenceDate >= sunday) {
  Logger.log('this week?');
  sheet2.appendRow(getDataFromCookieOrders(i));
}
else {
  Logger.log('smaller!');
}

}
everyWeekDeleteDuplicates();
}


function testReuseDeleteDuplicates (sheetName) {
  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange("A2:Z500");
    let emptyRowi = 0;
    let emptyRowj = 0;
   

  for (let i = 0; i < 300; i++) {
    sortByDate(sheetName);
     let emptyRowj = 0;
   let referenceRow = range.getValues()[i];  
   let referenceDateCell = sheet.getRange("A"+i).getValue();
   Logger.log("i is: " + i);
   Logger.log("emptyRowi: " + emptyRowi);
   Logger.log("emptyRowj: " + emptyRowj);
   Logger.log("ReferenceRow: " + referenceRow);
   

  if (referenceRow.join() === ",,,,,,,,,,,,,,,,,,,,,,,,,") {
    emptyRowi++;
    Logger.log("Empty Reference");
    if (emptyRowi >= 10) {
      emptyRowi = 0;
      break;
    }
  }
for(let j = i + 1; j < 400; j++) {
  let checkedRow = range.getValues()[j];
  let checkedDateCell = sheet.getRange("A"+j).getValue();
  Logger.log("CheckedRow: " + checkedRow);
   if (referenceRow.join() === checkedRow.join() && checkedRow.join() !== ",,,,,,,,,,,,,,,,,,,,,,,,,") {
Logger.log("Same!");
emptyRowi = 0;

sheet.getRange(j+2,1,1,26).setValues([["","","","","","","","","","","","","","","","","","","","","","","","","","",]]);
  }
if (checkedRow.join() === ",,,,,,,,,,,,,,,,,,,,,,,,,") {
  emptyRowj++;
  if(emptyRowj >= 10) {
    emptyRowj = 0;
    break;
  }
}
else {
  emptyRowi = 0;
}
 
 
}
  } 
 
}

function reuseDeleteDuplicates (sheetName) {
  
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange("A2:Z500");
    let emptyRowi = 0;
    let emptyRowj = 0;
   

  for (let i = 0; i < 300; i++) {
    sortByDate(sheetName);
     let emptyRowj = 0;
   let referenceRow = range.getValues()[i];  
   Logger.log("i is: " + i);
  //  Logger.log("emptyRowi: " + emptyRowi);
  //  Logger.log("emptyRowj: " + emptyRowj);
  //  Logger.log("ReferenceRow: " + referenceRow);
   

  if (referenceRow.join() === ",,,,,,,,,,,,,,,,,,,,,,,,,") {
    emptyRowi++;
    // Logger.log("Empty Reference");
    if (emptyRowi >= 3) {
      emptyRowi = 0;
      break;
    }
  }
for(let j = i + 1; j < 400; j++) {
  let checkedRow = range.getValues()[j];
  // Logger.log("CheckedRow: " + checkedRow);
   if (referenceRow.join() === checkedRow.join() && checkedRow.join() !== ",,,,,,,,,,,,,,,,,,,,,,,,,") {
Logger.log("Same!");
emptyRowi = 0;

sheet.getRange(j+2,1,1,26).setValues([["","","","","","","","","","","","","","","","","","","","","","","","","","",]]);
  }
if (checkedRow.join() === ",,,,,,,,,,,,,,,,,,,,,,,,,") {
  emptyRowj++;
  if(emptyRowj >= 3) {
    emptyRowj = 0;
    break;
  }
}
else {
  emptyRowi = 0;
}
 
 
}
  } 
 
}

function everyWeekDeleteDuplicates() {
  var thisWeek = "ThisWeek";
  var oneWeek = "OneWeek";
  var twoWeeks = "TwoWeeks";
  var threeWeeks = "ThreeWeeks";
  sortByDate(thisWeek);
  sortByDate(oneWeek);
  sortByDate(twoWeeks);
  sortByDate(threeWeeks);
  reuseDeleteDuplicates(thisWeek);
  reuseDeleteDuplicates(oneWeek);
  reuseDeleteDuplicates(twoWeeks);
  reuseDeleteDuplicates(threeWeeks);
  sortByDate(thisWeek);
  sortByDate(oneWeek);
  sortByDate(twoWeeks);
  sortByDate(threeWeeks);
}

function thisWeekDeleteDuplicates () {

  var sheetName = "ThisWeek";
    sortByDate(thisWeek);
  reuseDeleteDuplicates(sheetName);
  sortByDate(thisWeek);
}

function oneWeekDeleteDuplicates () {
  var sheetName = "OneWeek";
    sortByDate(oneWeek);
  reuseDeleteDuplicates(sheetName);
  sortByDate(oneWeek);
  
}

function twoWeeksDeleteDuplicates () {
  var sheetName = "TwoWeeks";
  sortByDate(twoWeeks);
  reuseDeleteDuplicates(sheetName);
  sortByDate(twoWeeks);
 
}

function threeWeeksDeleteDuplicates () {
  var sheetName = "ThreeWeeks";
     sortByDate(threeWeeks);
  reuseDeleteDuplicates(sheetName);
   sortByDate(threeWeeks);
}
function deleteEverything(sheetName) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetName);
var range = sheet.getRange("A2:Z500");
for(let i = 0; i<100; i++) {
  sheet.getRange(i+2,1,1,24).setValues([["","","","","","","","","","","","","","","","","","","","","","","","","","",]]);
}

}
function migrateData(giving, receiving) {
  var ss1 = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGiving = ss1.getSheetByName(giving);
   var rangeGiving = sheetGiving.getRange("A2:Z500");
   var ss2 = SpreadsheetApp.getActiveSpreadsheet();
  var sheetReceiving = ss2.getSheetByName(receiving);
  var rangeReceiving = sheetReceiving.getRange("A2:Z500");
  var emptyRow = 0;

  for(let i = 0; i<50; i++) {
    Logger.log("Replaced");
 var givenData = rangeGiving.getValues()[i];
  sheetReceiving.getRange(i+2,1,1,26).setValues([givenData]);
  if(givenData.join() === ",,,,,,,,,,,,,,,,,,,,,,,,,,,") {
    emptyRow++;
    if(emptyRow >= 10) {
      break;
    }
  }
  }
 

}

function weeklyMigration() {
  var thisWeek = "ThisWeek";
  var oneWeek = "OneWeek";
  var twoWeeks = "TwoWeeks";
  var threeWeeks = "ThreeWeeks";

everyWeekDeleteDuplicates();
  deleteEverything(thisWeek);
  migrateData(oneWeek, thisWeek);
  deleteEverything(oneWeek);
  migrateData(twoWeeks, oneWeek);
  deleteEverything(twoWeeks);
  migrateData(threeWeeks, twoWeeks);
  deleteEverything(threeWeeks);
  everyWeekDeleteDuplicates();

}

function sortHistorical () {
  var sheetName = "HistoricalCustomerData";
  sortByDate(sheetName);
}

function importHistoricalNames () {
  var giving = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HistoricalCustomerData');
  var givingRangeEmail = giving.getRange("B2:B500");
  var givingRangeName = giving.getRange("S2:S500");
  var givingRangePhoneNumber = giving.getRange("T2:T500");
  var givingRangeAddress = giving.getRange("U2:U500");
  var recieving = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerNames');
  
var data = true;
var index = 0;
var givenEmailData;
var givenNameData;
var givenPhoneData;
var givenAddressData;

while (data) {
var concatData = [];

concatData[0] = givingRangeName.getValues()[index].toString();
concatData[1] = givingRangeEmail.getValues()[index].toString();
concatData[2] = givingRangePhoneNumber.getValues()[index].toString();
concatData[3] = givingRangeAddress.getValues()[index].toString();
recieving.appendRow(concatData);
index++;
Logger.log(concatData);
if(concatData.join() === ",,,") {
  data = false;
}
}
}

function deleteCustomerOrders () {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerNames');
 var index = 0;
 var data = true;
  while (data) {
    ss.getRange(index+2,5,1,19).setValues([["","","","","","","","","","","","","","","","","","","","",]])[index];
    index++;
    Logger.log(ss.getRange("A2:Z500").getValues()[index].join());
    if (ss.getRange("A2:Z500").getValues()[index].join() === ",,,,,,,,,,,,,,,,,,,,,,,,,") {
      data = false;
    }
  }
}

function treatCustomerNames () {
  var sheetName = "CustomerNames";
  sortByDate(sheetName);
  importHistoricalNames();
  sortByDate(sheetName);
  reuseDeleteDuplicates(sheetName);
  sortByDate(sheetName);
}

function thursdays () {
  var cookieOrders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieOrders");
  var lastRow = cookieOrders.getLastRow();
  var formRow = cookieOrders.getRange(lastRow, 1, 1, 26);
  var deliveryDate = cookieOrders.getRange(lastRow, 26).getValues();
  // var aVals = thursday.getRange("A1:A").getValues();
  // var totRows = aVals.filter(String).length;
switch (deliveryDate.join()) {
  case "Thursday, 9 December 2021":
  var thursday9 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("9");
  thursday9.appendRow(formRow.getValues()[0]);
  Logger.log("Moved: " + formRow.getValues()[0]);
  reuseDeleteDuplicates("9");
  break;
  case "Thursday, 16 December 2021":
  var thursday16 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("16");
  thursday16.appendRow(formRow.getValues()[0]);
  Logger.log("Moved: " + formRow.getValues()[0]);
  reuseDeleteDuplicates("16");
  break;
  case "Thursday, 23 December 2021":
  var thursday23 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("23");
  thursday23.appendRow(formRow.getValues()[0]);
  Logger.log("Moved: " + formRow.getValues()[0]);
  reuseDeleteDuplicates("23");
  break;
  case "Thursday, 30 December 2021":
  var thursday30 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("30");
  thursday30.appendRow(formRow.getValues()[0]);
  Logger.log("Moved: " + formRow.getValues()[0]);
  reuseDeleteDuplicates("30");
  break;
  case "":
  Logger.log("EMPTY");
  break;
  default: 
  Logger.log("NONE");
  break;
}



}

function sendOrderEmail() {
  sortByDate("CookieOrders");
  var cookieOrders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieOrders");
  var lastRow = cookieOrders.getLastRow();
  var emailRange = cookieOrders.getRange(lastRow, 2);
  var nameRange = cookieOrders.getRange(lastRow, 19);
  var dateRange = cookieOrders.getRange(lastRow, 26);
  var martha_email = "marthamrave@gmail.com";
  var email = String(emailRange.getValues()[0]);
  var name = String(nameRange.getValues()[0]);
  var first_name = name.split(" ")[0];
  Logger.log(first_name);

  var date = String(dateRange.getValues()[0]);
  var emailContents = "";
  Logger.log("Last Row: " + lastRow);
  // Logger.log("Email: " + email);
  Logger.log("Name: " + name);
  Logger.log("Date: " + date);

  var cookieData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieData");
  var offsetColumns = 5;
  var dataRange = cookieData.getRange(31, lastRow + offsetColumns);
  var totalCost = dataRange.getValues()[0];

  let cookiesOrdered = cookieOrderList(lastRow);
  let emailCookies = "";

  for (var prop in cookiesOrdered){
emailCookies = emailCookies + "<br>" + prop + ": " + cookiesOrdered[prop];
  }



  
  emailContents = "Thank you so much " + first_name + ", your order has been received!" + "<br> Your cookies and quantities ordered are:<br>" + emailCookies + "<br>The total for your cookie order comes to: $" + totalCost + "*<br><br>If you have chosen to make an e-transfer, please direct payment to martharave@yahoo.com. If you're paying by cash please pay at time of delivery.<br>" + "<br>Your cookie delivery date is scheduled for " + date + ".<br><br>Please feel free to email me if you have any questions at martharave@yahoo.com." + "<br><br>I look forward to baking holiday treats for you!<br><br>Sincerely,<br><br>Martha<br><br>*if there are any additional fees, I will email you directly with the correct total";
  var body = '<p style = "font-family:georgia,garamond,serif;font-size:16px;font-style:default;"> '+ emailContents +' </p> ';
  
  var additionalInfo = "";

if(isAdditionalInfo(lastRow)){
  additionalInfo = "<p style = 'font-family:georgia,garamond,serif;font-size:40px;font-style:bold;'> CHECK ADDITIONAL INFORMATION!!!! </p>"
}
var martha_body = additionalInfo + body

MailApp.sendEmail(martha_email, "Your Cookie Order", "", {htmlBody: martha_body});
MailApp.sendEmail(email, "Your Cookie Order", "", {htmlBody: body});

  
}

function cookieOrderList(rowIndex){
  let cookieList = ["Gingerbread Snowflake",	"Sugar Cookie Snowflake",	"Toblerone Shortbread",	"Lemon & Ginger Shortbread",	"Pecan Dusters",	"Graham Toffee Bark",	"Cherry Shortbread",	"Chocolate Toffee Shortbread","Brown Sugar Shortbread", "Jar of Toblerone","Elf Assortment",	"Santa Assortment",	"Gingerbread Boys",	"Gingerbread Girls", "Gingerbread Christmas Tree"];
 
  let newCookieList = [];
  let cookieVals = [];
  var cookieOrders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieOrders");
  for (i = 0; i < 15; i++){
var range = cookieOrders.getRange("R" + rowIndex + "C" + (i + 3));
var cellData = range.getValues()[0];
if(cellData.join() !== ''){
cookieVals.push(cellData.join());
newCookieList.push(cookieList[i]);
}

  }
  

  var cookieAndVal = {}
  newCookieList.forEach((key, i) => cookieAndVal[key] = cookieVals[i]);
  return cookieAndVal;
  
}

function isAdditionalInfo(rowIndex){
var cookieOrders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieOrders");
var range = cookieOrders.getRange("R" + rowIndex + "C18");
var cellData = range.getValues()[0];
if(cellData.join() !== ''){
  return true;
}
else {
  return false;
}
}

function portChanges () {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CookieOrders");
  var ui = SpreadsheetApp.getUi();
  var row = parseInt(ui.prompt("What row did you edit?").getResponseText());
  Logger.log(row);
  var editRow = sheet.getRange(row, 1, 1, 26).getValues();
  var timestamp = sheet.getRange(row, 1).getValues().join();
  var deliveryDate = sheet.getRange(row, 26).getValues().join();
Logger.log(timestamp);
Logger.log(deliveryDate);
  switch (deliveryDate) {

    case "Thursday, 9 December 2021":
      var thursday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("9");
      var totRows = thursday.getMaxRows().toString();
      var aVals = thursday.getRange("A1:A").getValues();
      var totRows = aVals.filter(String).length;
      Logger.log("totRows: " + totRows);
      Logger.log("Thursday, 9 December");
      for (var i = 2; i <= totRows; i++) {
        Logger.log(thursday.getRange(i,1).getValues().join());
        if (timestamp == thursday.getRange(i,1).getValues().join()){
          thursday.getRange(i, 1, 1, 26).setValues(editRow);
          Logger.log("Set values to: " + editRow);
          return;
        }
      } 
      ui.alert("Matching Order Not Found");
      break;
    case "Thursday, 16 December 2021":
      var thursday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("16");
      var totRows = thursday.getMaxRows().toString();
      var aVals = thursday.getRange("A1:A").getValues();
      var totRows = aVals.filter(String).length;
      Logger.log("totRows: " + totRows);
      Logger.log("Thursday, 16 December");
      for (var i = 2; i <= totRows; i++) {
        Logger.log(thursday.getRange(i,1).getValues().join());
        if (timestamp == thursday.getRange(i,1).getValues().join()){
          thursday.getRange(i, 1, 1, 26).setValues(editRow);
          Logger.log("Set values to: " + editRow);
          return;
        }
      } 
      ui.alert("Matching Order Not Found");
      break;
    case "Thursday, 23 December 2021":
      var thursday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("23");
      var totRows = thursday.getMaxRows().toString();
      var aVals = thursday.getRange("A1:A").getValues();
      var totRows = aVals.filter(String).length;
      Logger.log("totRows: " + totRows);
      Logger.log("Thursday, 23 December");
      for (var i = 2; i <= totRows; i++) {
        Logger.log(thursday.getRange(i,1).getValues().join());
        if (timestamp == thursday.getRange(i,1).getValues().join()){
          thursday.getRange(i, 1, 1, 26).setValues(editRow);
          Logger.log("Set values to: " + editRow);
          return;
        }
      } 
      ui.alert("Matching Order Not Found");
      break;
    case "Thursday, 30 December 2021":
      var thursday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("30");
      var totRows = thursday.getMaxRows().toString();
      var aVals = thursday.getRange("A1:A").getValues();
      var totRows = aVals.filter(String).length;
      Logger.log("totRows: " + totRows);
      Logger.log("Thursday, 30 December");
      for (var i = 2; i <= totRows; i++) {
        Logger.log(thursday.getRange(i,1).getValues().join());
        if (timestamp == thursday.getRange(i,1).getValues().join()){
          thursday.getRange(i, 1, 1, 26).setValues(editRow);
          Logger.log("Set values to: " + editRow);
          return;
        }
      }
      ui.alert("Matching Order Not Found"); 
      break;
    default:
    ui.alert("Matching Order Not Found"); 
    break;
  } 
} 
function sortThursdays(){
  sortByDate('9');
  sortByDate('16');
  sortByDate('23');
  sortByDate('30');
  sortByDate('Corporate')
}

function thursdayDelete(){
  reuseDeleteDuplicates("9");
reuseDeleteDuplicates("16");
reuseDeleteDuplicates("23");
}

function nineDelete(){
  reuseDeleteDuplicates("9");
}
function sixteenDelete(){
  reuseDeleteDuplicates("16");
}
function twentythreeDelete(){
  reuseDeleteDuplicates("23");
}


