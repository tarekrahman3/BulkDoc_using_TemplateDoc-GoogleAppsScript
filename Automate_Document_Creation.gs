/**
 * Bound Spreadsheet: https://docs.google.com/spreadsheets/d/1sicfypeLACVJFB75GjUu-9_MOSpD-PtHIth9r9uyG_8/
 * Template Document: https://docs.google.com/document/d/1jXkpC-nhAVtDbKI7k1glIshOw1Qx4uUJaAgicz9Ie_g/
**/

/**
 * * List of functions with parameters:
 * main()
 * getInputData()
 * docCreate()
 * createDocForEach(templateID,newFolder,FOLDER_NAME,n)
 * constructAddressLine1(body, address_1stLine)
 * constructAddressLine2(body, address_2ndLine)
**/

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Create Document')
      .addItem('Run', 'main')
      .addToUi();
}

function main() {
  fileManagement()
  var array = docCreate();
  var sourceSheet = SpreadsheetApp.getActiveSheet();
  targetRange = sourceSheet.getRange(2,4,array.length,2);
  targetRange.setValues(array);
  SpreadsheetApp.getUi().alert(newFolderLink);
}

// getting addresses from spreadsheet
function getInputData() {
  dictList = [];
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Address List');
  range = sheet.getRange(2,2,(sheet.getLastRow())-1,2).getValues();
  for(var i=0; i<range.length;i++){
   data = {
     'a':range[i][0],
     'b':range[i][1]
     }
    dictList.push(data);
  }
  return dictList;
}


function prepareResources(){
  var formattedDate = Utilities.formatDate(new Date(), "GMT"+6, "yyyy-MM-dd'T'HH:mm:ss'Z'");
  templateID = '1jXkpC-nhAVtDbKI7k1glIshOw1Qx4uUJaAgicz9Ie_g';
  FOLDER_NAME = "Doc Batch Update Exports " + formattedDate;
  newFolder = DriveApp.getFolderById('1i34BTh21lg_vYr8tu1YpJdfZKTs1YOAJ').createFolder(FOLDER_NAME);
  newFolderLink = 'Folder Liink: ' + 'https://drive.google.com/drive/folders/' + (newFolder.getId())  + '/';
  return prepareResources;
}

// change the values of newly created document by one using the inputData() dictionary
function docCreate(){
  
  // heart of this fuction, the dictionary 
  addressDictionary = getInputData();
  console.log(addressDictionary)
  new_docs_list = [];
  for (var n = 0; n < addressDictionary.length; n++) {
    var newDoc = createDocForEach(templateID,newFolder,FOLDER_NAME,n);
    newDocId = newDoc.getId();
    body = DocumentApp.openById(newDocId).getBody();
    console.log(n+1 + ' ' + String(addressDictionary[n].a))
    constructAddressLine1(body, String(addressDictionary[n].a));
    constructAddressLine2(body, String(addressDictionary[n].b));
    newDocId = newDoc.getId()
    new_docs_list.push(
      ['https://docs.google.com/document/d/' + newDocId,
      'https://docs.google.com/document/d/' + newDocId + '/export?format=pdf']);
  }
  console.log(new_docs_list)
  return new_docs_list;
}

// creating new document by copying the template document
function createDocForEach(templateID,newFolder,FOLDER_NAME,n){
  addressDictionary = getInputData();
  template = DocumentApp.getActiveDocument();
  folder = DriveApp.getFolderById(newFolder.getId());
  newDoc =  DriveApp.getFileById(templateID).makeCopy((addressDictionary[n].a), folder);
  return newDoc;
}

// find and replace AddressLine1
function constructAddressLine1(body, address_1stLine){
  get_item1 = body.findText('AddressLine1').getElement().editAsText().replaceText('AddressLine1', address_1stLine);
  return constructAddressLine1;
}

// find and replace AddressLine2
function constructAddressLine2(body, address_2ndLine){
  get_item2 = body.findText('AddressLine2').getElement().editAsText().replaceText('AddressLine2', address_2ndLine);
  return constructAddressLine2;
}


