/**
* @author      Bineet Mishra
* @date        10/01/14
* @description Google Spreadsheets plugin to query/edit SFDC data
*/

var USERNAME_PROPERTY_NAME = "username";
var PASSWORD_PROPERTY_NAME = "password";
var SECURITY_TOKEN_PROPERTY_NAME = "securityToken";
var SESSION_ID_PROPERTY_NAME = "sessionId";
var SERVICE_URL_PROPERTY_NAME = "serviceUrl";
var INSTANCE_URL_PROPERTY_NAME = "instanceUrl";
var IS_SANDBOX_PROPERTY_NAME = "isSandbox";
var NEXT_RECORDS_URL_PROPERTY_NAME = "nextRecordsUrl";
var SANDBOX_SOAP_URL = "https://test.salesforce.com/services/Soap/u/27.0";
var PRODUCTION_SOAP_URL = "https://www.salesforce.com/services/Soap/u/27.0"


/* Defaults for this particular spreadsheet, change as desired */
var DEFAULT_FORMAT = 'Pretty';
var DEFAULT_LANGUAGE = 'JavaScript';
var DEFAULT_STRUCTURE = 'List';

/*Added */
var SOQL = 'soql';
var SOBJECT = 'sObject';

var OPERATION_CREATE = 'create';
var OPERATION_UPDATE = 'update';

/**
 * @return String Username.
 */
function getUsername() {
  var key = ScriptProperties.getProperty(USERNAME_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}
 
/**
 * @param String Username.
 */
function setUsername(key) {
  ScriptProperties.setProperty(USERNAME_PROPERTY_NAME, key);
}
 
/**
 * @return String Password.
 */
function getPassword() {
  var key = ScriptProperties.getProperty(PASSWORD_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}
 
/**
 * @param String Password.
 */
function setPassword(key) {
  ScriptProperties.setProperty(PASSWORD_PROPERTY_NAME, key);
}

/**
 * @return String Security Token.
 */
function getSecurityToken() {
  var key = ScriptProperties.getProperty(SECURITY_TOKEN_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}
 
/**
 * @param String Security Token.
 */
function setSecurityToken(key) {
  ScriptProperties.setProperty(SECURITY_TOKEN_PROPERTY_NAME, key);
}

/**
 * @return String Session Id.
 */
function getSessionId() {
  var key = ScriptProperties.getProperty(SESSION_ID_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}

/**
 * @param String Session Id.
 */
function setSessionId(key) {
  ScriptProperties.setProperty(SESSION_ID_PROPERTY_NAME, key);
}

/**
 * @return String Instance URL.
 */
function getInstanceUrl() {
  var key = ScriptProperties.getProperty(INSTANCE_URL_PROPERTY_NAME);
  if (key == null) {
    key = "";
  }
  return key;
}

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
  ScriptProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
}


/**
 * @param String use sandbox url.
 */
function setUseSandbox(key) {
  ScriptProperties.setProperty(IS_SANDBOX_PROPERTY_NAME, key);
}

/**
 * @return bool if using sandbox.
 */
function getUseSandbox() {
  var key = ScriptProperties.getProperty(IS_SANDBOX_PROPERTY_NAME);
  if (key == null) {
    key = false;
  }
  return key;
}


/**
 * @param String url for next records url.
 */
function setNextRecordsUrl(key) {
  
  if(key==undefined)
    key = "";
    
  ScriptProperties.setProperty(NEXT_RECORDS_URL_PROPERTY_NAME, key);
}

/**
 * @return String url for next records url (querymore).
 */
function getNextRecordsUrl() {
  var key = ScriptProperties.getProperty(NEXT_RECORDS_URL_PROPERTY_NAME);
  if (key == null || key == undefined) {
    key = "";
  }
  return key;
}

/**
 * @param String Instance URL.
 */
function setInstanceUrl(key) {
  ScriptProperties.setProperty(INSTANCE_URL_PROPERTY_NAME, key);
}

/**
 * @return bool if using sandbox.
 */
function getSfdcSoapEndpoint(){
  var isSandbox = getUseSandbox() == "true" ? true: false;
  if(isSandbox)
    return SANDBOX_SOAP_URL;
  else 
    return PRODUCTION_SOAP_URL;
}

function getRestEndpoint(){
  //Move this logic to the property
  var queryEndpoint = ".salesforce.com";
  
  var endpoint = getInstanceUrl().replace("api-","").match("https://[a-z0-9]*");
  
  return endpoint+queryEndpoint;
}

/**
 * @return String SOQL.
 */
function getSoql() {
  var key = ScriptProperties.getProperty(SOQL);
  if (key == null) {
    key = "";
  }
  return key;
}
 
/**
 * @param String SOQL.
 */
function setSoql(key) {
  ScriptProperties.setProperty(SOQL, key);
}

/**
 * @return String sObject.
 */
function getSObject() {
  var key = ScriptProperties.getProperty(SOBJECT);
  if (key == null) {
    key = "";
  }
  return key;
}
 
/**
 * @param String sObject.
 */
function setSObject(key) {
  ScriptProperties.setProperty(SOBJECT, key);
}

                   

function onInstall(){
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Settings", functionName: "renderSettingsDialog"},
                     {name: "Login", functionName: "login"},
                     {name: "Query", functionName: "renderQueryDialog"},
                     {name: "Query More", functionName: "sendQueryMore"},
                     {name: "Refresh", functionName: "refreshSheet"},
                     {name: "Insert All", functionName: "createNew"},
                     {name: "Save All", functionName: "saveUpdates"}
                    ];
  ss.addMenu("SFDC Connector", menuEntries);
}

/** Retrieve config params from the UI and store them. */
function saveConfiguration(e) {
  setUsername(e.parameter.username);
  setPassword(e.parameter.password);
  setSecurityToken(e.parameter.securityToken);
  setUseSandbox(e.parameter.sandbox);
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function renderSettingsDialog(){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("Salesforce Configuration");
  app.setStyleAttribute("padding", "10px");
  
  var helpLabel = app.createLabel(
      "Enter your Username, Password, and Security Token");
  helpLabel.setStyleAttribute("text-align", "justify");
 
  var usernameLabel = app.createLabel("Username:");
  var username = app.createTextBox();
  username.setName("username");
  username.setWidth("75%");
  username.setText(getUsername());
  
  var passwordLabel = app.createLabel("Password:");
  var password = app.createPasswordTextBox();
  password.setName("password");
  password.setWidth("75%");
  password.setText(getPassword());
  
  var securityTokenLabel = app.createLabel("Security Token:");
  var securityToken = app.createTextBox();
  securityToken.setName("securityToken");
  securityToken.setWidth("75%");
  securityToken.setText(getSecurityToken());
  
  var sandboxLabel = app.createLabel("Sandbox:");
  var sandbox = app.createCheckBox();
  sandbox.setName("sandbox");
  sandbox.setValue(getUseSandbox() == "true" ? true: false);
  
  var saveHandler = app.createServerClickHandler("saveConfiguration");
  var saveButton = app.createButton("Save Configuration", saveHandler);
  
  var listPanel = app.createGrid(4, 2);
  listPanel.setStyleAttribute("margin-top", "10px")
  listPanel.setWidth("100%");
  listPanel.setWidget(0, 0, usernameLabel);
  listPanel.setWidget(0, 1, username);
  listPanel.setWidget(1, 0, passwordLabel);
  listPanel.setWidget(1, 1, password);
  listPanel.setWidget(2, 0, securityTokenLabel);
  listPanel.setWidget(2, 1, securityToken);
  listPanel.setWidget(3, 0, sandboxLabel);
  listPanel.setWidget(3, 1, sandbox);
  
  // Ensure that all form fields get sent along to the handler
  saveHandler.addCallbackElement(listPanel);
  
  var dialogPanel = app.createFlowPanel();
  dialogPanel.add(helpLabel);
  dialogPanel.add(listPanel);
  dialogPanel.add(saveButton);
  app.add(dialogPanel);
  doc.show(app);
}

function login() {
  
  var message="<?xml version='1.0' encoding='utf-8'?>" 
    +"<soap:Envelope xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' " 
    +   "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://"
    +   "www.w3.org/2001/XMLSchema'>" 
    +  "<soap:Body>" 
    +     "<login xmlns='urn:partner.soap.sforce.com'>" 
    +        "<username>" + getUsername() + "</username>"
    +        "<password>"+ getPassword() + getSecurityToken() + "</password>"
    +     "</login>" 
    +  "</soap:Body>" 
    + "</soap:Envelope>";
  
   var httpheaders = {SOAPAction: "login"};
   var parameters = {
     method : "POST",
     contentType: "text/xml",
     headers: httpheaders,
     payload : message};

    try{
      var result = UrlFetchApp.fetch(getSfdcSoapEndpoint(), parameters).getContentText();
      var soapResult = Xml.parse(result, false);
            
      setSessionId(soapResult.Envelope.Body.loginResponse.result.sessionId.getText());
      setInstanceUrl(soapResult.Envelope.Body.loginResponse.result.serverUrl.getText());
      
    } catch(e){
      Logger.log("EXCEPTION!!!");
      Logger.log(e);
      Browser.msgBox(e);
    }
}

function renderGridData(object, renderHeaders){
  var sheet = SpreadsheetApp.getActiveSheet();
 
  var data = [];
  var sObjectAttributes = {};
  
  //Need to always build headers for row length/rendering
  var headers = buildHeaders(object.records);
  
  if(renderHeaders){  
    data.push(headers);
  }
  
  for (var i in object.records) {
    var values = [];
    for(var j in object.records[i]){
      if(j!="attributes"){
        values.push(object.records[i][j]);
      } else {
        var id = object.records[i][j].url.substr(object.records[i][j].url.length-18,18);
        //Logger.log(id);
        sObjectAttributes[id] = object.records[i][j].type;
      }
    }
    data.push(values);
  }
  if(data.length >1){
    Logger.log('Last Row: ' + sheet.getLastRow());
    Logger.log('Data Length: ' + data.length);
    var destinationRange = sheet.getRange(sheet.getLastRow()+1, 1, data.length, headers.length);
    destinationRange.setValues(data);
  }
  else{
    Browser.msgBox('No records to display');
    renderQueryDialog();
  }
}


function buildHeaders(records){
  var headers = [];
  for(var i in records[0]){
    if(i!="attributes")
      headers.push(i);
  }
  //Logger.log(headers);
  return headers;
}

function sendSoqlQuery(e){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  if(e.parameter.soql != ""){
    setSoql(e.parameter.soql);
  }
  if(e.parameter.sobjects != ""){
    setSObject(e.parameter.sobjects);
  }
  
  var results = query(encodeURIComponent(getSoql()));
  renderGridData(processResults(results), true);  
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function refreshSheet(){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  
  var results = query(encodeURIComponent(getSoql()));
  renderGridData(processResults(results), true);  
}

function sendQueryMore(){
  var results = queryMore(); 
  if(results != null)
    renderGridData(processResults(results), false);
}

function createNew(){
  var rowsData = getRows(OPERATION_CREATE);//For rest of the coulmns excluding Id
  for (var i = 0; i < rowsData.length; i++) {
    var message = getRowAsJson(rowsData[i], OPERATION_CREATE);
    Logger.log(message);
    createRecord(message);    
  }
  Browser.msgBox("Records created successfully");
}

function createRecord(message){
  return fetch(getRestEndpoint()+"/services/data/v22.0/sobjects/" + getSObject() + "/", message);
}


function saveUpdates(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();// For Id column
  
  var rowsData = getRows(OPERATION_UPDATE);//For rest of the coulmns excluding Id
  for (var i = 0; i < rowsData.length; i++) {
    var id = data[i+1][sheet.getLastColumn()-1];
    var message = getRowAsJson(rowsData[i], OPERATION_UPDATE);
    Logger.log(message);
    Logger.log(id);
    saveRecord(id, message);    
  }
  Browser.msgBox("Records saved successfully");
}


function saveRecord(id, message){
  return fetch(getRestEndpoint()+"/services/data/v22.0/sobjects/" + getSObject() + "/" + id + "?_HttpMethod=PATCH", message);
}

function processResults(results){
  var object = Utilities.jsonParse(results);
  setNextRecordsUrl(object.nextRecordsUrl);
  
  return object;
}

function renderQueryDialog(){  
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("SQOL Query");
  app.setStyleAttribute("padding", "10px");
  app.setHeight(250);
  
  var helpLabel = app.createLabel("Enter your SOQL query below:");
  helpLabel.setStyleAttribute("text-align", "justify");
  
  //Set available options for objects
  var json = getSObjects();
  var object = processResults(json);
  var items = [];
  //items.push('---Select One---');
  for(var i=0; i < object['sobjects'].length; i++){
    items.push(object['sobjects'][i]['name']);
  }
  Logger.log(object['sobjects'][0]['name']);
  var sObjects = makeListBox(app, 'sobjects', items);
  
  //var sObject = app.createTextArea().setId("sobject").setName("sobject").setWidth("40%").setText(getSObject());
  var soql = app.createTextArea().setId('soql').setName("soql").setWidth("80%").setText(getSoql());
  var queryAll = app.createCheckBox().setText("Query All?");
  
  var sendHandler = app.createServerClickHandler("sendSoqlQuery");
  var sendButton = app.createButton("Query", sendHandler);
  
  var listHandler = app.createServerClickHandler('changeSObject');
  listHandler.addCallbackElement(sObjects);  
  sObjects.addChangeHandler(listHandler);
  
  var dialogPanel = app.createFlowPanel();
  dialogPanel.add(helpLabel);
  dialogPanel.add(sObjects);
  //dialogPanel.add(sObject);
  dialogPanel.add(soql);
  dialogPanel.add(queryAll);
  dialogPanel.add(sendButton);
  sendHandler.addCallbackElement(dialogPanel);
  app.add(dialogPanel);
  doc.show(app);
}

//This is the sObject onChange function
function changeSObject(e){
  var app = UiApp.getActiveApplication();
  setSObject(e.parameter.sobjects);
  app.getElementById('soql').setText('SELECT id FROM ' + getSObject());
  return app;
}

/**
 * @param String SQOL query
 */
function query(soql){
  return fetch(getRestEndpoint()+"/services/data/v27.0/"+"query?q="+soql);
}

function getSObjects(){
  return fetch(getRestEndpoint()+"/services/data/v27.0/sobjects/");
}

/**
 * @param String nextrecords Url
 */
function queryMore(){
  Logger.log("Next Url:" + getNextRecordsUrl());
  
  var nextRecordsUrl = getNextRecordsUrl();
  
  if(nextRecordsUrl !="")
    return fetch(getRestEndpoint()+getNextRecordsUrl());
  else {
    Browser.msgBox("No more records to query.");
    return null;
  }
}

/**
 * @param String url to fetch from SFDC via REST API
 */
function fetch(url){
  
  var httpheaders = {Authorization: "OAuth " + getSessionId()};
  var parameters = {headers: httpheaders}; 
  
  //Logger.log(parameters);
  try{
    return UrlFetchApp.fetch(url, parameters).getContentText();
  } catch(e){
    Logger.log(e);
    Browser.msgBox(e);
  }  
}

/**
 * @param String url to fetch from SFDC via REST API
 */
function fetch(url, message){
  
  var httpheaders = {Authorization: "OAuth " + getSessionId(), "Content-Type": "application/json"};
  var parameters = 
  {
    "headers": httpheaders,
    "payload": message
  }; 
  
  //Logger.log(parameters);
  try{
    return UrlFetchApp.fetch(url, parameters).getContentText();
  } catch(e){
    Logger.log(e);
    Browser.msgBox(e);
  }
}

//Methods to create UI elements
function makeListBox(app, name, items) {
  var listBox = app.createListBox().setId(name).setName(name);
  listBox.setVisibleItemCount(1);
  
  var cache = CacheService.getPublicCache();
  //var selectedValue = cache.get(name);
  var selectedValue = getSObject();
  Logger.log(selectedValue);
  for (var i = 0; i < items.length; i++) {
    listBox.addItem(items[i]);
    if (items[i] == selectedValue) {
      listBox.setSelectedIndex(i);
    }
  }
  return listBox;
}

/*Convert to JSON methods*/

function getRows(operation) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowsData = getRowsData(sheet, getConvertOptions(operation));
  return rowsData;
}

function getRowAsJson(rowData, operation) {
  var json = makeJSON(rowData, getConvertOptions(operation));
  return json;
}
  
function getConvertOptions(operation) {
  var options = {};
  
  options.operation = operation;
  
  var cache = CacheService.getPublicCache();
  cache.put('language', options.language);
  cache.put('format',   options.format);
  cache.put('structure',   options.structure);
  cache.put('operation',   options.operation);
  
  Logger.log(options);
  return options;
}

function makeJSON(object, options) {
  var jsonString = JSON.stringify(object, null, 4);
  return jsonString;
}


function getRowsData(sheet, options) {
  var lastColumn;
  if(options.operation == OPERATION_UPDATE){
    lastColumn = sheet.getLastColumn()-1;
  }
  else{
    lastColumn = sheet.getLastColumn();
  }
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), lastColumn);
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows(), lastColumn);
  var objects = getObjects(dataRange.getValues(), headers);
  return objects;
}


function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

//Utility methods
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

function isDigit(char) {
  return char >= '0' && char <= '9';
}


function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }
  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }
  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }
  return ret;
}
