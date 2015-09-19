function initiateOAuth() {
  Logger.log('**** In initiateOAuth() ***'); 
  
  HTMLToOutput = "<html><h1>You need to login</h1><a href='"+getURLForAuthorization()+"'>click here to start</a><br>Re-open this window when you return.</html>";
  SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutput(HTMLToOutput));
}

function getURLForAuthorization() {
   var scriptProperties = PropertiesService.getScriptProperties();
   return scriptProperties.getProperty('SFDC_AUTHORIZE_URL') + '?response_type=code&client_id=' + 
          scriptProperties.getProperty('SFDC_CLIENT_ID') + '&redirect_uri=' + 
          ScriptApp.getService().getUrl();
}

function doGet(e) {
  var HTMLToOutput;
  //If we get "code" as a parameter in, then this is a callback
  if(e.parameters.code){
    getAndStoreAccessToken(e.parameters.code);
    HTMLToOutput = '<html><h1>Finished with oAuth</h1>You can close this window.</html>';
  }
  return HtmlService.createHtmlOutput(HTMLToOutput);
}

function getAndStoreAccessToken(code){
   var scriptProperties = PropertiesService.getScriptProperties();
  //Construct Token URL with client id, client secret, redirect uri, retrieved short lived code 
  //and grant_type = authorization_code
  var authURL = scriptProperties.getProperty('SFDC_TOKEN_URL') + '?client_id=' + 
          scriptProperties.getProperty('SFDC_CLIENT_ID') + '&client_secret=' + 
          scriptProperties.getProperty('SFDC_CLIENT_SECRET') + '&grant_type=authorization_code&redirect_uri=' + 
          ScriptApp.getService().getUrl() + '&code=' + code;
  
  //Use URLFetch service
  var response = UrlFetchApp.fetch(authURL).getContentText();   
  Logger.log(response);
  
  var tokenResponse = JSON.parse(response);
  
  var userProperties = PropertiesService.getUserProperties();
  
  Logger.log('\n\nAccess Token: ' + tokenResponse.access_token);
  Logger.log('\nRefresh Token: ' + tokenResponse.refresh_token);
  Logger.log('\nInstance URL: ' + tokenResponse.instance_url);
  
  //Set access token, instance url & refresh token for subsequent use
  userProperties.setProperty('SFDC_AUTH_TOKEN', tokenResponse.access_token);
  userProperties.setProperty('SFDC_REFRESH_TOKEN', tokenResponse.refresh_token);
  userProperties.setProperty('SFDC_INSTANCE_URL', tokenResponse.instance_url);  
}

function getUrlFetchOptions() {
  var token = PropertiesService.getUserProperties().getProperty('SFDC_AUTH_TOKEN');
  return {
    "contentType" : "application/json",
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json"
    }
  };
}

function getUrlFetchPOSTOptions(payload) {
  var token = PropertiesService.getUserProperties().getProperty('SFDC_AUTH_TOKEN');
  return {
    "contentType" : "application/json",
    "method": "post",
    "payload" : payload,
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json"
    }
  };
}

function showUserProperties() {
  var userProps = PropertiesService.getUserProperties().getProperties();
  
  var i = 1;
  for (var key in userProps) {
    Logger.log('\n%s. %s: %s', i++, key, userProps[key]);
  }
  

  if (i == 1) {
    Logger.log('No user properties available');
  }
}
