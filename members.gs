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

// Name of group that contains singers contacts
CONTACT_GROUP_NAME="MembersList";

// Add/remove requested members from first sheet
function updateMembersList() {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  editSheet = ss.getSheetByName(EDITING_SHEET_NAME);
  masterListSheet = ss.getSheetByName(MASTER_LIST_SHEET_NAME);

  // Do the deletion first so that the an email can be changed by
  // putting the member in both sides with the different email address
  
  // Data has been removed
  nMaxSrcRows=editSheet.getDataRange().getNumRows() - (START_EDIT_ROW - 1);
  if(nMaxSrcRows > 0) {  
    deleteRequestedMembers(editSheet, masterListSheet);
  }
  
  // Is there anything to try
  var nMaxSrcRows=editSheet.getDataRange().getNumRows() - (START_EDIT_ROW - 1);
  if(nMaxSrcRows > 0) {
    addNewMembers(editSheet, masterListSheet);
  }
  
  // Organise final list
  removeDuplicates(masterListSheet);
  sortBySurname(masterListSheet);
  // Check for any email updates
  updateContactInfo(masterListSheet);
}

// Add requested members from first sheet
// @param srcSheet Sheet containing new members to add
// @param destSheet Sheet where members are to be added
function addNewMembers(srcSheet, destSheet) {
  var nMaxSrcRows=srcSheet.getDataRange().getNumRows() - 2;
  var newMemberData = srcSheet.getRange(START_EDIT_ROW, NEW_SURNAME_COL, nMaxSrcRows, NEW_EMAIL_COL-NEW_SURNAME_COL+1);
  // None to add
  if(newMemberData.getValue() == "") return;
  
  // First add the new members to the contact group
  newMemberValues = newMemberData.getValues();
  addContacts(newMemberValues);
  
  var nCurrentMembers = destSheet.getDataRange().getNumRows() - 1; //Take off title row  
  var destRange = destSheet.getRange(START_LIST_ROW+nCurrentMembers,1);
  newMemberData.moveTo(destRange);

  // Put the date in
  var dateCell = srcSheet.getRange(INPUT_DATE_CELL);
  var startRow=START_LIST_ROW+nCurrentMembers;
  var endRow=startRow+newMemberData.getNumRows() - 2; //Take off title row & 1 for difference
  dateCell.copyValuesToRange(destSheet, OUTPUT_DATE_COL, OUTPUT_DATE_COL, startRow, endRow);
}

/// Adds new contacts to the email group
/// @param info 2D array of surname, forename, email
/// @param group The group to update
function addContacts(info) {
  
  // The getContact function has to load all contacts each time so has terrible performance in a loop
  // for looking up if a contact exists so load them once
  var group = ContactsApp.getContactGroup(CONTACT_GROUP_NAME);
  var currentContacts = ContactsApp.getContactsByGroup(group);
  
  for(i in info) {
    surname = info[i][NEW_SURNAME_COL-1];
    forename = info[i][NEW_FORENAME_COL-1];
    emailAddress = info[i][NEW_EMAIL_COL-1];
    if(emailAddress == "n/a" || emailAddress == "" || surname == "" || forename == "") continue;

    // Do we have this email address already
    var found=false; 
    for(j in currentContacts) {
      contact = currentContacts[j];
      if(contact.getPrimaryEmail() == emailAddress)
      {
        found=true;
        Browser.msgBox("A contact with email \"" + emailAddress 
                        + "\" already exists. Name=" + contact.getFullName() 
                        + "\\n Email details not updated.");
        break;
      }
    }
    if(!found) {
      group.addContact(ContactsApp.createContact(forename, surname, emailAddress));
    }
  }
}

// Delete requested members from first sheet
// @param srcSheet Sheet containing members to be deleted
// @param destSheet Sheet where members are to be deleted
function deleteRequestedMembers(srcSheet, destSheet) {  
  var nMaxSrcRows=srcSheet.getDataRange().getNumRows() - 2;
  var delMemberData = srcSheet.getRange(START_EDIT_ROW, DEL_SURNAME_COL, nMaxSrcRows, DEL_EMAIL_COL-DEL_SURNAME_COL+1);
  // None to delete
  if(delMemberData.getValue() == "") return;

  
  // First remove the  members from the contact group & the contacts
  var delMemberDataValues=delMemberData.getValues();
  contactGroup = getContactGroup();
  removeContacts(delMemberDataValues, contactGroup);
  
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

/// Remove contacts from the email group
/// @param info 2D array of surname, forename, email
/// @param group The group to update
function removeContacts(info, group) {
  for(i in info) {
    emailAddress = info[i][NEW_EMAIL_COL-1];
    contact = ContactsApp.getContact(emailAddress);
    if(contact) {
      group.removeContact(contact);
      ContactsApp.deleteContact(contact);
    }
  }
}

// Create or return the contact group
// @return The members contact group
function getContactGroup() {
  var membersGroup = ContactsApp.getContactGroup(CONTACT_GROUP_NAME);
  if(!membersGroup) {
    membersGroup = ContactsApp.createContactGroup(CONTACT_GROUP_NAME);
  }
  return membersGroup;
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

// Update an email addresses that may have changed
// @param sheet The sheet containing the members list
function updateContactInfo(sheet) {
  var wholeRange = sheet.getDataRange();
  var dataValues = sheet.getRange(2, 1, wholeRange.getNumRows(), 3).getValues();
  // The getContact function has to load all contacts each time so has terrible performance in a loop
  // for looking up if a contact exists so load them once
  var group = ContactsApp.getContactGroup(CONTACT_GROUP_NAME);
  var currentContacts = ContactsApp.getContactsByGroup(group);

  for(i in dataValues) {
    row = dataValues[i];
    Logger.log(row);
    surname = row[0];
    forename = row[1];
    listEmail = row[2];
    for(j in currentContacts) {
      contact = currentContacts[j];
      contactEmail = contact.getEmails()[0].getAddress();
      contactName = contact.getFamilyName();
      if(surname == contactName && listEmail != contactEmail) {
        contact.removeFromGroup(group);
        contact.deleteContact();
        group.addContact(ContactsApp.createContact(forename, surname, listEmail));
        break;
      }
    }
  }
}

//-----------------------------------------------------------------------------------------
// Register functions
//-----------------------------------------------------------------------------------------


REGISTER_NAME="Register";
REGISTER_INDEX=2;
ROWS_PER_BLOCK=67;
REGISTER_COL1=1;
REGISTER_COL2=5;
TICK_COL_WIDTH=30; //pixels
  
// Create a register from the current
// members list
function createRegister() {
  membersSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  registerSheet = membersSpreadSheet.getSheetByName(REGISTER_NAME);
  if(registerSheet) {
    SpreadsheetApp.setActiveSheet(registerSheet);
    membersSpreadSheet.deleteActiveSheet();
  }
  registerSheet = membersSpreadSheet.insertSheet(REGISTER_INDEX); //makes sheet active
  registerSheet.setName(REGISTER_NAME);
  // Make sure its big enough or copyValuesToRange fails
  registerSheet.insertRowsAfter(100,150); //Add 400 rows to end
  
  var membersSheet = membersSpreadSheet.getSheetByName(MASTER_LIST_SHEET_NAME);
  var nMembers = membersSheet.getDataRange().getNumRows()-1; //take off top row
  
  var ntotalBlocks=Math.floor(nMembers / ROWS_PER_BLOCK);
  var ntailValues=nMembers % (ROWS_PER_BLOCK);
  Logger.log("Number of blocks: " + ntotalBlocks);
  Logger.log("Number of tail values: " + ntailValues);

  var destSurnameCol=REGISTER_COL1;
  var srcStartRow=START_LIST_ROW;
  var destStartRow=1;
  var nrows=ROWS_PER_BLOCK;
  for(var i=0; i < ntotalBlocks; ++i) {    
    var memberListRange = membersSheet.getRange(srcStartRow,NEW_SURNAME_COL, nrows, NEW_FORENAME_COL-NEW_SURNAME_COL+1);
    var rowEnd=destStartRow+nrows;
    var colEnd=destSurnameCol+1;
    Logger.log("Copying " + nrows + " from " + srcStartRow + "," + NEW_SURNAME_COL 
               + " to start=" + destStartRow + "," + destSurnameCol + " end="+rowEnd + "," + colEnd);
    memberListRange.copyValuesToRange(registerSheet, destSurnameCol, colEnd,destStartRow,rowEnd);
    // Activate borders
    registerSheet.getRange(destStartRow,destSurnameCol,nrows,3).setBorder(true,true, false, true, true, true);
    
    // Next block
    srcStartRow += nrows;
    if(destSurnameCol == REGISTER_COL1) {
      destSurnameCol=REGISTER_COL2;
    }
    else {
      destSurnameCol=REGISTER_COL1;
      destStartRow=i*ROWS_PER_BLOCK+1;
      //if(i > 0) destStartRow += 2; //Add block separators
    }
  }
  if(ntailValues > 0) {
    Logger.log("Copying " + ntailValues + " from " + srcStartRow + "," + NEW_SURNAME_COL + " to " + destStartRow + "," + destSurnameCol);
    var memberListRange = membersSheet.getRange(srcStartRow,NEW_SURNAME_COL, ntailValues,NEW_FORENAME_COL-NEW_SURNAME_COL+1);
    memberListRange.copyValuesToRange(registerSheet, destSurnameCol, destSurnameCol+1,destStartRow,destStartRow+ntailValues);  
    // Activate borders
    registerSheet.getRange(destStartRow,destSurnameCol,ntailValues,3).setBorder(true,true, true, true, true, true);
  }
  
  // Boldify surnames
  var registerData = registerSheet.getDataRange();
  var totalDataRows = registerData.getNumRows();
  registerSheet.getRange(1, REGISTER_COL1, totalDataRows, 1).setFontWeight("bold");
  registerSheet.getRange(1, REGISTER_COL2, totalDataRows, 1).setFontWeight("bold");
  
  // Resize the "tick" column
  registerSheet.setColumnWidth(REGISTER_COL1+2, TICK_COL_WIDTH);
  registerSheet.setColumnWidth(REGISTER_COL2+2, TICK_COL_WIDTH);
    
}
