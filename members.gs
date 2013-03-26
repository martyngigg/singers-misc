//----------------------------------------------------------------------------------
// Functions to support register
//----------------------------------------------------------------------------------

EDITING_SHEET_NAME="Workspace";
MASTER_LIST_SHEET_NAME="Members List";

NEW_SURNAME_COL=1;
NEW_FORENAME_COL=2;
NEW_EMAIL_COL=3;


DEL_SURNAME_COL=5;
DEL_FORENAME_COL=6;
DEL_EMAIL_COL=7;

START_EDIT_ROW=4;
START_LIST_ROW=2;

INPUT_DATE_CELL="B1";
OUTPUT_DATE_COL=4;

//
// Add/remove requested members from first sheet
//
function updateMembersList() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  editSheet = ss.getSheetByName(EDITING_SHEET_NAME);
  masterListSheet = ss.getSheetByName(MASTER_LIST_SHEET_NAME);
  
  // Is there anything to try
  var nMaxSrcRows=editSheet.getDataRange().getNumRows() - (START_EDIT_ROW - 1);
  if(nMaxSrcRows > 0) {
    addNewMembers(editSheet, masterListSheet);
  }
  // Data has been removed
  nMaxSrcRows=editSheet.getDataRange().getNumRows() - (START_EDIT_ROW - 1);
  if(nMaxSrcRows > 0) {  
    deleteRequestedMembers(editSheet, masterListSheet);
  }
  
  // Organise final list
  removeDuplicates(masterListSheet);
  sortBySurname(masterListSheet);
}

// Add requested members from first sheet
// @param srcSheet Sheet containing new members to add
// @param destSheet Sheet where members are to be added
function addNewMembers(srcSheet, destSheet) {
  var nMaxSrcRows=srcSheet.getDataRange().getNumRows() - 2;
  var newMemberData = srcSheet.getRange(START_EDIT_ROW, NEW_SURNAME_COL, nMaxSrcRows, NEW_EMAIL_COL-NEW_SURNAME_COL+1);
  // None to add
  if(newMemberData.getValue() == "") return;
  
  var nCurrentMembers = destSheet.getDataRange().getNumRows() - 1; //Take off title row  
  var destRange = destSheet.getRange(START_LIST_ROW+nCurrentMembers,1);
  newMemberData.moveTo(destRange);

  // Put the date in
  var dateCell = srcSheet.getRange(INPUT_DATE_CELL);
  var startRow=START_LIST_ROW+nCurrentMembers;
  var endRow=startRow+newMemberData.getNumRows() - 2; //Take off title row & 1 for difference
  dateCell.copyValuesToRange(destSheet, OUTPUT_DATE_COL, OUTPUT_DATE_COL, startRow, endRow);
    
}

// Delete requested members from first sheet
// @param srcSheet Sheet containing members to be deleted
// @param destSheet Sheet where members are to be deleted
function deleteRequestedMembers(srcSheet, destSheet) {
  var nMaxSrcRows=srcSheet.getDataRange().getNumRows() - 2;
  var delMemberData = srcSheet.getRange(START_EDIT_ROW, DEL_SURNAME_COL, nMaxSrcRows, DEL_EMAIL_COL-DEL_SURNAME_COL+1);
  // None to delete
  if(delMemberData.getValue() == "") return;

  var delMemberDataValues=delMemberData.getValues();
  var masterListLength=destSheet.getDataRange().getNumRows();
  var memberListData = destSheet.getRange(START_LIST_ROW,NEW_SURNAME_COL,masterListLength-1,NEW_EMAIL_COL-NEW_SURNAME_COL+1).getValues();
  for(i in delMemberDataValues) {
    var delRow=delMemberDataValues[i];
    var rowNo=START_LIST_ROW;
    for(j in memberListData) {
      var memberRow=memberListData[j];
      if(delRow.join() == memberRow.join()) {
        masterListSheet.deleteRow(rowNo);
        break;
      }
      rowNo += 1;
    }
  }
  delMemberData.clear();
}

// Remove any duplicate members. Duplicates have to
// have the same surname,firstname & email
// @param sheet The sheet that contains the members list
function removeDuplicates(sheet) {
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
        break;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

// Sort the sheet by surname
// @param sheet The sheet containing the members list
function sortBySurname(sheet) { 
  var numRowsInSheet = sheet.getMaxRows();
  var wholeRange = sheet.getDataRange();
  var data = sheet.getRange(2, 1, wholeRange.getNumRows(), wholeRange.getNumColumns());
  data.sort([1]);
}

function onOpen() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var menuEntries = [{name: "DemoUI", functionName: "updateEmailContactsList"}];
//  ss.addMenu("Tutorial", menuEntries);
}

// Name of group that contains singers contacts
SINGERS_CONTACTS_GROUP="MembersList";

// Update the email contacts list
function updateEmailContactsList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var editSheet = ss.getSheetByName(EDITING_SHEET_NAME);
  var s = editSheet.getRange(DATE_CELL).getValue();
  Browser.msgBox(s);
  
  
  // First trash current list then recreate. Easiest way to deal with changed emails etc
  //group = ContactsApp.getContactGroup(SINGERS_CONTACTS_GROUP);
  
//
//  ss = SpreadsheetApp.getActiveSpreadsheet();
//  masterListSheet = ss.getSheetByName("Members List");
//  nMembers = masterListSheet.getDataRange().getNumRows() - 1;
//  var memberListData = masterListSheet.getRange(START_LIST_ROW,NEW_SURNAME_COL,nMembers,NEW_EMAIL_COL-NEW_SURNAME_COL+1).getValues();
//
//  // i is the index from 0 into 2D javascript array memberListData
//  for(i in memberListData) { 
//    emailAddress = memberListData[i][NEW_EMAIL_COL-1];
//    contacts = ContactsApp.getContactsByEmailAddress(emailAddress);
//    if(contacts.length == 0) {
//      
//    }
//    else {
//      contact
//    }
//  }
//  
//  Browser.msgBox(s);
//  
////  var membersGroup = ContactsApp.getContactGroup(SINGERS_CONTACTS_GROUP);
////  if(!membersGroup) {
////    membersGroup = ContactsApp.createContactGroup(SINGERS_CONTACTS_GROUP);
////  }
  
  
  
    //membersGroup = ContactsApp.createContactGroup("Members");
  //membersGroup.addContact(ContactsApp.createContact("Martyn", "Gigg", "test@email.com"));
  
  //   var app = UiApp.createApplication();  
//   var form = app.createFormPanel();
//   var flow = app.createFlowPanel();
//   flow.add(app.createRichTextArea().setStyleName("textBox").setStyleAttribute("font-family", "Arial"));
//   flow.add(app.createSubmitButton("Submit"));
//   form.add(flow);
//   app.add(form);
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   spreadsheet.show(app);
}