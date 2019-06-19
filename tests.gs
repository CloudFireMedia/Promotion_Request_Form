// Logging
// -------

function Log_(message) {

  if (!TEST_ENABLE_LOGGING) {
    return;
  }

  var authRequired = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus()
  
  if (authRequired === ScriptApp.AuthorizationStatus.REQUIRED) {
    return
  }

  if (typeof message === 'object') {
    message = JSON.stringify(message)
  }

  message = new Date() + ' - ' + message

  SpreadsheetApp
    .openById(Config.get('PROMOTION_FORM_RESPONSES_GSHEET_ID'))
    .getSheetByName('Log')
    .appendRow([message])
    
  console.log(message)
}

// Tests
// -----

function test(a) {
  if (a === undefined) {
    return
  }
  return
}

function test_onFormSubmit() {
  onFormSubmit({triggerUid: '111'})
}

function test_pollStaffSpreadsheet() {
  pollStaffSpreadsheet_()
}

function test_checkPromotionCalendar() {
  checkPromotionCalendar_({triggerUid: '111'})
}

function test_Todoist_onFormSubmit() {

    var todoistConfig = {
    
      spreadsheetId:     '1JEqPQJSiBliliqw1y-wrrdP6ikU11DPuIF72l-rN84g',
      rowNumber:         15,
      
//      token:             'd380194cdf40c2a384bdfac3a7bdaa59e23cb85b',     // todoist1@ajrcomputing.com
      token:      '441bfeb4a104be0333a8190f94703068e01e94ff',
      
//      taskTemplateId:    '0BzM8_MjdRURAOHFOYnFINDdlS2M', // Chad
      taskTemplateId:    '1pUfm-4huDlWqU3QHmN4odytDLk2u09BA', // Andrew
      
      commentTemplateId: '1oIdPamSeUWPDQgb5F6NXbjKWcY3gvOVuw6LicegiSwg',       
      designerEmail:     'jamielascher@gmail.com',            
      properties:        PropertiesService.getScriptProperties(), 
      lock:              LockService.getScriptLock()
    }

    Todoist.onFormSubmit(todoistConfig)
}

// Config
// ------

function test_dumpConfig() {

  Logger.log('UserProperties:')
  Logger.log(PropertiesService.getUserProperties().getProperties())
  
  var properties = PropertiesService.getDocumentProperties()
  
  if (properties !== null) {
    Logger.log('DocumentProperties:')
    Logger.log(properties.getProperties())  
  }

  Logger.log('ScriptProperties:')
  Logger.log(PropertiesService.getScriptProperties().getProperties())   

  Logger.log('Triggers:')  
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    Logger.log('Trigger: ' + trigger.getHandlerFunction() + ' (' + trigger.getUniqueId() + ')')
  })
}

function test_clearConfig() {

  PropertiesService.getUserProperties().deleteAllProperties()
  
  var properties = PropertiesService.getDocumentProperties()
  
  if (properties !== null) {
    properties.deleteAllProperties() 
  }

  PropertiesService.getScriptProperties().deleteAllProperties()   
  
  test_dumpConfig()
}