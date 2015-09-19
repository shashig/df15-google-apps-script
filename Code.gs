function printActiveSheetContents() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    Logger.log('First Name: ' + data[i][0] + ', Last Name: ' + data[i][1]);
  }
}

function copySheet1ToSheet3() {
  var inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet3');
  var data = inputSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
     outputSheet.appendRow([data[i][0], data[i][1]]);
  }
}

function normalizePracticeAttachments() {
  var inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PracticeAttachments-Denormalized");
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PracticeAttachments-Normalized");
  var attachmentsFolder = '';
  
  if (inputSheet != null) {
    //Clear data and write header row
    outputSheet.clear();
    outputSheet.appendRow(['Practice Name', 'ParentId', 'Name', 'Body']);
    var data = inputSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      outputSheet.appendRow([data[i][1], data[i][0], 'FactSheet', attachmentsFolder + data[i][2]]);
      outputSheet.appendRow([data[i][1], data[i][0], 'QuickFact1', attachmentsFolder + data[i][3]]);
      outputSheet.appendRow([data[i][1], data[i][0], 'QuickFact2', attachmentsFolder + data[i][4]]);
      outputSheet.appendRow([data[i][1], data[i][0], 'QuickFact3', attachmentsFolder + data[i][5]]);
      outputSheet.appendRow([data[i][1], data[i][0], 'QuickFact4', attachmentsFolder + data[i][6]]);
    }
  }
}

function getBioAttachments() {
  var inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bios2");
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FilesToDownload");
  var urls = [];
  outputSheet.clear();
  //Add URLs to be downloaded
  if (inputSheet != null) {
    var data = inputSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] != '' && urls.indexOf('http://www.serviceshub.com/wp-content/uploads/' + data[i][2]) == -1) {
        urls.push('http://www.serviceshub.com/wp-content/uploads/' + data[i][2]);
      }
      if (urls.indexOf(data[i][3]) == -1) {
        urls.push(data[i][3]);
      }
    }
    for (var url in urls) {
      outputSheet.appendRow([urls[url]]);
    }
  }
}

function getSFDCData() {
  var instanceUrl = PropertiesService.getUserProperties().getProperty('SFDC_INSTANCE_URL');
  var queryUrlSegment = PropertiesService.getScriptProperties().getProperty('SFDC_QUERY_URL_SEGMENT');
  var fieldNames = [];
  
  //Use first two rows of data as input
  //Row 1/Column 1: Object Name
  //Row 2: Field Names
  var sheet = SpreadsheetApp.getActiveSheet();
  var inputRange = sheet.getDataRange().offset(0, 0, 2);
  var data = inputRange.getValues();
  
  //Construct the SOQL from the retrieved data
  var object = data[0][0];
  var soql = 'select ';
  for (var i = 0; i < data[1].length; i++) {
    if (i > 0) {
      soql += ', ';
    }
    soql += data[1][i];
    
    //Save the field names for later use while reading data
    fieldNames.push(data[1][i]);
  }
  soql += ' from ' + object;
  
  Logger.log(soql);
  
  //Construct the REST Call URL
  var getDataURL = instanceUrl + queryUrlSegment + soql;
  
  //Invoke the REST Call
  var response = UrlFetchApp.fetch(getDataURL,getUrlFetchOptions()).getContentText();  
  Logger.log(response);
  
  var dataResponse = JSON.parse(response);
  
  //Get reference to output range and clear it. 
  var outputRange = sheet.getDataRange().offset(2, 0);
  outputRange.clear();

  //Write total rows retrieved count and boolean indicating 
  //whether all rows have been retrieved
  var totalRowsCell = sheet.getRange('C1');
  var allRowsDoneCell = sheet.getRange('E1');  
  totalRowsCell.setValue(dataResponse.totalSize);
  allRowsDoneCell.setValue(dataResponse.done);
  
  var objects = dataResponse.records;

  //Process the response records  
  for (var i = 0; i < objects.length; i++) {
    var object = objects[i];
    var dataRow = [];

    //Process the fields
    for (var j = 0; j < fieldNames.length; j++) {
      //If the field is on a parent object, we need to get the parent 
      //object first and then get the field within the parent. 
      if (fieldNames[j].indexOf('.') > -1) {
        var dotPos = fieldNames[j].indexOf('.');
        var parentObjVar = fieldNames[j].substring(0, dotPos);
        var parentObjFieldName = fieldNames[j].substring(dotPos + 1);
        
        var parentObj = object[parentObjVar];
        
        //Check if parent is null. 
        if (parentObj) {
          dataRow.push(parentObj[parentObjFieldName]);
        } else {
          dataRow.push(null);
        }
      } else {
        //Field is on the object
        dataRow.push(object[fieldNames[j]]);
      }
    }
    sheet.appendRow(dataRow);
  }
}

function insertParentObject() {
  var instanceUrl = PropertiesService.getUserProperties().getProperty('SFDC_INSTANCE_URL');
  var parentObjectResourceUrl = PropertiesService.getScriptProperties().getProperty('SFDC_PARENT_OBJECT_RESOURCE_URL');
  var postURL = instanceUrl + parentObjectResourceUrl;

  var srcOrgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SourceOrg-Parent');
  var inputRange = srcOrgSheet.getDataRange().offset(2, 0);
  var data = inputRange.getValues();

  var targetOrgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TargetOrg-Parent');
  
  for (var i = 0; i < data.length - 2; i++) {
    var targetRow = [];
    targetRow.push(data[i][0]);
    targetRow.push(data[i][1]);
    
    var payload =  JSON.stringify(
      {"Name" : data[i][1]}
    );
    var response = UrlFetchApp.fetch(postURL, getUrlFetchPOSTOptions(payload)).getContentText();  
    Logger.log(response);
    var dataResponse = JSON.parse(response);

    targetRow.push(dataResponse.id);
    targetOrgSheet.appendRow(targetRow);
  }
}

function insertChildObject() {
  var instanceUrl = PropertiesService.getUserProperties().getProperty('SFDC_INSTANCE_URL');
  var childObjectResourceUrl = PropertiesService.getScriptProperties().getProperty('SFDC_CHILD_OBJECT_RESOURCE_URL');
  var postURL = instanceUrl + childObjectResourceUrl;

  var srcOrgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SourceOrg-Child');
  var inputRange = srcOrgSheet.getDataRange().offset(2, 0);
  var data = inputRange.getValues();

  var targetOrgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TargetOrg-Child');
  
  var parentMapping = getParentMapping();
  
  for (var i = 0; i < data.length - 2; i++) {
    var targetRow = [];
    targetRow.push(data[i][0]);
    targetRow.push(data[i][1]);
    targetRow.push(data[i][2]);
    targetRow.push(data[i][3]);
    targetRow.push(parentMapping[data[i][3]]);
    
    var payload =  JSON.stringify(
      {"Name" : data[i][1], "Parent_Object__c" : parentMapping[data[i][3]]}
    );
    Logger.log(payload);
    
    var response = UrlFetchApp.fetch(postURL, getUrlFetchPOSTOptions(payload)).getContentText();  
    Logger.log(response);
    var dataResponse = JSON.parse(response);

    targetRow.push(dataResponse.id);
    targetOrgSheet.appendRow(targetRow);
  }
}

function getParentMapping() {
  var mapping = {};
  var targetOrgParentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TargetOrg-Parent');
  var inputRange = targetOrgParentSheet.getDataRange().offset(2, 0);
  var data = inputRange.getValues();

  for (var i = 0; i < data.length - 2; i++) {
    mapping[data[i][0]] = data[i][2];
  }
  
  Logger.log(JSON.stringify(mapping));
  return (mapping);
}

function setupMenu() {
  Logger.log('Within setupMenu');
  var menuEntries = [];
  menuEntries.push({name: "Print to Log", functionName: "printActiveSheetContents"});
  menuEntries.push({name: "Copy Data", functionName: "copySheet1ToSheet3"});
  menuEntries.push({name: "Normalize Practice Attachments", functionName: "normalizePracticeAttachments"});
  menuEntries.push({name: "Get Bio Attachments", functionName: "getBioAttachments"});
  SpreadsheetApp.getActive().addMenu('Custom Sheet Menu', menuEntries);
  
   //Logger.log(ScriptApp.getService().getUrl());
  menuEntries = [];
  menuEntries.push({name: "Authorize", functionName: "initiateOAuth"});
  menuEntries.push({name: "Get Data", functionName: "getSFDCData"});
  menuEntries.push({name: "Insert Parent", functionName: "insertParentObject"});
  menuEntries.push({name: "Insert Child", functionName: "insertChildObject"});
  SpreadsheetApp.getActive().addMenu('Salesforce', menuEntries);
  Logger.log('Exiting setupMenu');
}

function onOpen() {
  setupMenu();
}

function debugFunc() {
  var payload =  JSON.stringify(
      {"FirstName" : "Shashi",
       "LastName" : "Guru",
       "Email" : "shashi@hfs.com",
       "Phone" : "98456"
      }
    );
  Logger.log(payload);
}
