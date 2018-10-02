/** Trigger implementations and related callbacks */

/** Force enable drive scope */
/** Drive.Files.list(); */  

/**
 * onOpen trigger.
 *
 * @param {Event} e - Event object.
 */
//function onOpen(e) {
//
//    var menu = FormApp.getUi().createAddonMenu();
//    
//    if (e.authMode === ScriptApp.AuthMode.NONE) {
//    
//        menu.addItem("Run Setup", "startSetup");
//        
//    } else {
//        
//        menu.addItem("Run Setup", "startSetup")
//            .addSeparator()
//            .addItem("Map Staff Spreadsheet",           "mapToStaffSpreadsheet")
//            .addItem("Map Responses Spreadsheet",       "mapToResponsesSpreadsheet")
//            .addItem("Set Target Response Sheet",       "setTargetSheetForResponses")
//            .addItem("Set Pull Range for Names",        "setPullRangeForUserNames")
//            .addItem("Set Pull Range for Emails",       "setPullRangeForUserEmails")
//            .addItem("Map HootSuite Spreadsheet",       "mapToHootSuiteSpreadsheet")
//            .addItem("Set Target HootSuite Sheet",      "setTargetSheetForHootSuite")
//            .addItem("Set Target Calendar",             "setTargetCalendar")
//            .addItem("Set Todoist Comment Template",    "setTodoistCommentTemplate")
//            .addItem("Set Todoist Tasks Template",      "setTodoistTasksTemplate")
//            .addItem("Set Todoist Auth Token",          "setTodoistAuthToken")
//            .addItem("Show settings",                   "showCache")            
//            //.addItem("Reset to defaults",               "resetToDefaults")
//    }
//    
//    menu.addToUi();
//   
//}
//
///** Creates prompt to map form to staff spreadsheet id */
//function mapToStaffSpreadsheet() {
//    mapToSpreadsheet("MAP STAFF SPREADSHEET", "Please enter (or paste) the id of the staff spreadsheet.", "STAFF_SPREADSHEET_ID");
//}
//
///** Creates prompt to map form to responses spreadsheet id */
//function mapToResponsesSpreadsheet() {
//    mapToSpreadsheet("MAP RESPONSES SPREADSHEET", "Please enter (or paste) the id of the responses spreadsheet.", "RESPONSES_SPREADSHEET_ID");
//}
//
///** Create prompt to map form to HootSuite spreadsheet id */
//function mapToHootSuiteSpreadsheet() {
//    mapToSpreadsheet("MAP HOOTSUITE SPREADSHEET", "Please enter (or paste) the id of the Hootsuite spreadsheet,", "HOOTSUITE_SPREADSHEET_ID");
//}
//
///**
// * Begin setup flow, triggering a series of prompts
// * instructing the user to set requested parameters.
// *
// */
//function startSetup() {
//
//  throw new Error('All setup is now done at CloudFire installation')
//
//    var ui = FormApp.getUi(),
//        result = ui.alert(
//            "SETUP",
//            "This form must be configured before use. You'll be \npresented with a series of prompts to guide you.\n",
//            ui.ButtonSet.OK
//        );
//  
//    if (result === ui.Button.OK) {
//        
//        clearCache_();
//        
//        // chained function invocations per setup step
//        
//        mapToSpreadsheet("MAP STAFF SPREADSHEET", "Please enter (or paste) the id of the staff spreadsheet.", "STAFF_SPREADSHEET_ID", function() {
//        
//            setPullRangeForUserNames(function() {
//            
//                setPullRangeForUserEmails(function() {
//                
//                    mapToSpreadsheet("MAP RESPONSES SPREADSHEET", "Please enter (or paste) the id of the responses spreadsheet.", "RESPONSES_SPREADSHEET_ID", function() {
//                    
//                        setTargetSheetForResponses(function() {
//                        
//                            mapToSpreadsheet("MAP HOOTSUITE SPREADSHEET", "Please enter (or paste) the id of the Hootsuite spreadsheet.", "HOOTSUITE_SPREADSHEET_ID", function(){
//                            
//                                setTargetSheetForHootSuite(function() {
//                                
//                                    setTargetCalendar(function() {
//                                    
//                                        setTodoistAuthToken(function(){
//
//                                            setTodoistTasksTemplate(function(){
//
//                                                setTodoistCommentTemplate(function(){
//                                                
//                                                    initialize(function(){
//                                                  
//                                                        FormApp.getUi().alert("SETUP COMPLETE!");
//                                                        return;
//                                                    })                  
//                                                })
//                                            })
//                                        })
//                                    })                    
//                                })
//                            })
//                        })
//                    })
//                })
//            })
//        });
//        
//    }
//    
//    return;
//    
//    
//}
//
///** Clears all cache settings */
//function clearCache_() {
//    var propertyCache = new PropertyCache(),
//        staffSSID = Config.get("STAFF_DATA_GSHEET_ID");
//    
//    if (staffSSID) {
//        propertyCache.remove(staffSSID, true);
//    }
//    
//    propertyCache.remove("STAFF_SPREADSHEET_ID", true);
//    propertyCache.remove("RESPONSES_SPREADSHEET_ID", true);
//    propertyCache.remove("RESPONSE_SHEET_NAME", true);
//    propertyCache.remove("PULL_RANGE_FOR_USER_NAMES", true);
//    propertyCache.remove("PULL_RANGE_FOR_USER_EMAILS", true);
//    propertyCache.remove("HOOTSUITE_SPREADSHEET_ID", true);
//    propertyCache.remove("HOOTSUITE_SHEET_NAME", true);
//    propertyCache.remove("CALENDAR_NAME", true);
//    propertyCache.remove("TODOIST_AUTH_TOKEN", true);
//    propertyCache.remove("TODOIST_TASKS_TEMPLATE_ID", true);
//    propertyCache.remove("TODOIST_COMMENT_TEMPLATE_ID", true);
//    
//}

///** Shows all of the cache settings */
//function showCache() {
//    var propertyCache = new PropertyCache(),
//
//        prompt =
//            "Staff Spreadsheet ID: "             + propertyCache.get("STAFF_SPREADSHEET_ID") + "\n" +
//            "Responses Spreadsheet ID: "         + propertyCache.get("RESPONSES_SPREADSHEET_ID") + "\n" +
//            "Responses Sheet name: "             + (propertyCache.get("RESPONSE_SHEET_NAME") || DEFAULT_RESPONSE_SHEET_NAME) + "\n" +
//            "Pull range for users names: "       + (propertyCache.get("PULL_RANGE_FOR_USER_NAMES") || DEFAULT_PULL_RANGE_FOR_USER_NAMES) + "\n" +
//            "Pull range for users emails: "      + (propertyCache.get("PULL_RANGE_FOR_USER_EMAILS") || DEFAULT_PULL_RANGE_FOR_USER_EMAILS) + "\n" +
//            "Hootsuite Spreadsheet ID: "         + propertyCache.get("HOOTSUITE_SPREADSHEET_ID") + "\n" +
//            "Hootsuite Sheet Name: "             + (propertyCache.get("HOOTSUITE_SHEET_NAME") || DEFAULT_HOOTSUITE_SHEET_NAME) + "\n" +
//            "Calendar Name: "                    + propertyCache.get("CALENDAR_NAME") + "\n" +
//            "Todoist Auth token: "               + propertyCache.get("TODOIST_AUTH_TOKEN") + "\n" +
//            "Todoist Tasks Template CSV ID: "    + propertyCache.get("TODOIST_TASKS_TEMPLATE_ID") + "\n" +
//            "Todoist Comment Template GDoc ID: " + propertyCache.get("TODOIST_COMMENT_TEMPLATE_ID")
//
//    FormApp.getUi().alert(prompt)
//  
//}

/**
 * Maps the form to a spreadsheet; stores the spreadsheet
 * id as a property accessible to the scripts bound to 
 * the form.
 *
 * @param {String} title       - Title of the prompt.
 * @param {String} prompt      - Prompt statement.
 * @param {String} key         - Key to store spreadsheet id.
 * @param {Function} callback  - An optional callback to be invoked when the 
 *                               dialog recieves confirmation. Used to chain
 *                               multiple invocations.
 */
//function mapToSpreadsheet(title, prompt, key, callback) {
//    var propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            title,
//            prompt,
//            ui.ButtonSet.OK
//        );
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        try {
//          
//            var id = result.getResponseText().trim()
//            var ss = SpreadsheetApp.openById(id);
//            
//            propertyCache.put(key, ss.getId(), true);
//            callback && callback();
//            
//        } catch(e) {
//        
//            if (e.message.indexOf("Bad value ") !== -1 || e.message.indexOf("is missing (perhaps it was deleted?)") !== -1) {
//                        
//                result = ui.alert(
//                    "Error accessing Spreadsheet (" +  title + ")! \nTry again.",
//                    ui.ButtonSet.OK
//                );
//                
//                if (result === ui.Button.OK) {
//                    mapToSpreadsheet(title, prompt, key, callback);
//                }
//                
//            } else {
//            
//              throw e
//            }
//            
//        }
//        
//    }
//    
//}
//
///** 
// * Create prompt to map form to Todoist comment template GDoc id 
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations
// */
//function setTodoistCommentTemplate(callback) {
//    var propertyCache = new PropertyCache(),
//        key = "TODOIST_COMMENT_TEMPLATE_ID",
//        title = "Todoist Comment Template GDoc ID",
//        prompt = "Please enter (or paste) the id of the Todoist Comment Template GDoc.",
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            title,
//            Utilities.formatString(prompt),
//            ui.ButtonSet.OK
//        );
//        
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//
//        try {
//          
//            var id = result.getResponseText().trim()
//            var gdoc = DocumentApp.openById(id);
//            
//            propertyCache.put(key, gdoc.getId(), true);
//            callback && callback();
//            
//        } catch(e) {
//        
//            if (e.message.indexOf("Bad value ") !== -1 || e.message.indexOf("is missing (perhaps it was deleted?)") !== -1) {
//                        
//                var result = ui.alert(
//                    "Error accessing " +  title + "! \nTry again.",
//                    ui.ButtonSet.OK
//                );
//                
//                if (result === ui.Button.OK) {
//                    setTodoistCommentTemplate(callback);
//                }
//                
//            } else {
//            
//              throw e
//            }
//            
//        }
//        
//    }
//    
//}
//
///** 
// * Create prompt to set Todoist tasks template CSV ID 
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations
// */
//function setTodoistTasksTemplate(callback) {
//    var propertyCache = new PropertyCache(),
//        key = "TODOIST_TASKS_TEMPLATE_ID",
//        title = "MAP TODOIST TASK TEMPLATE",
//        prompt = "Please enter (or paste) the id of the Todoist tasks template CSV.",
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            title,
//            Utilities.formatString(prompt),
//            ui.ButtonSet.OK
//        );
//        
//    if (result.getSelectedButton() === ui.Button.OK) {
//
//        try {
//          
//            var id = result.getResponseText().trim()
//            var file = DriveApp.getFileById(id);
//            
//            propertyCache.put(key, file.getId(), true);
//            callback && callback();
//            
//        } catch(e) {
//        
//            if (e.message.indexOf("Bad value ") !== -1 || e.message.indexOf("is missing (perhaps it was deleted?)") !== -1) {
//                        
//                var result = ui.alert(
//                    "Error accessing " +  title + "! \nTry again.",
//                    ui.ButtonSet.OK
//                );
//                
//                if (result === ui.Button.OK) {
//                    setTodoistTasksTemplate(callback);
//                }
//                
//            } else {
//            
//              throw e
//            }
//            
//        }
//        
//    }
//    
//}
//    
///**
// * Sets the desired pull range for user names.
// * Defaults to the constant DEFAULT_PULL_RANGE_FOR_USER_NAMES.
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setPullRangeForUserNames(callback) {
//
//    var range,
//        propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET PULL RANGE FOR USER NAMES",
//            Utilities.formatString("Enter (or paste) the desired range to\npull user names.\n\nOr leave it empty to use the\ndefault range (%s)", DEFAULT_PULL_RANGE_FOR_USER_NAMES),
//            ui.ButtonSet.OK
//        );
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        range = result.getResponseText().trim();
//        
//        if (range !== "") propertyCache.put("PULL_RANGE_FOR_USER_NAMES",range, true);
//        
//        if (range === "" && propertyCache.get("PULL_RANGE_FOR_USER_NAMES")) {
//            propertyCache.remove("PULL_RANGE_FOR_USER_NAMES", true);
//        }
//                
//        callback && callback();
//        
//    }
//    
//}
//
///**
// * Sets the desired pull range for user emails.
// * Defaults to the constant DEFAULT_PULL_RANGE_FOR_USER_EMAILS.
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setPullRangeForUserEmails(callback) {
//
//    var range,
//        propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET PULL RANGE FOR USER EMAILS",
//            Utilities.formatString("Enter (or paste) the desired range to\npull user emails.\n\nOr leave it empty to use the\ndefault range (%s)", DEFAULT_PULL_RANGE_FOR_USER_EMAILS),
//            ui.ButtonSet.OK
//        );
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        var range = result.getResponseText().trim();
//        
//        if (range !== "") propertyCache.put("PULL_RANGE_FOR_USER_EMAILS", range, true);
//        
//        if (range === "" && propertyCache.get("PULL_RANGE_FOR_USER_EMAILS")) {
//            propertyCache.remove("PULL_RANGE_FOR_USER_EMAILS", true);
//        }
//                
//        callback && callback();
//        
//    }
//
//
//}
//
///**
// * Sets the response sheet name.
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setTargetSheetForResponses(callback) {
//    var sheetName,
//        propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET TARGET SHEET FOR RESPONSES",
//            Utilities.formatString("Enter (or paste) the name of the target sheet.\n\n Or leave it empty to use the \ndefault response sheet (%s)", DEFAULT_RESPONSE_SHEET_NAME),
//            ui.ButtonSet.OK
//        )
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        var sheetName = result.getResponseText().trim();
//        
//        if (sheetName !== "") propertyCache.put("RESPONSE_SHEET_NAME", sheetName, true);
//        
//        if (sheetName === "" && propertyCache.get("RESPONSE_SHEET_NAME")) {
//            propertyCache.remove("RESPONSE_SHEET_NAME", true);
//        }
//                
//        
//        callback && callback();
//    }
//    
//}
//
///**
// * Sets the hootsuite sheet name.
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setTargetSheetForHootSuite(callback) {
//    var sheetName,
//        propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET TARGET SHEET FOR HOOTSUITE",
//            Utilities.formatString("Enter (or paste) the name of the target sheet.\n\n Or leave it empty to use the \ndefault sheet name (%s)", DEFAULT_HOOTSUITE_SHEET_NAME),
//            ui.ButtonSet.OK
//        )
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        var sheetName = result.getResponseText().trim();
//        
//        if (sheetName !== "") propertyCache.put("HOOTSUITE_SHEET_NAME", sheetName, true);
//        
//        if (sheetName === "" && propertyCache.get("HOOTSUITE_SHEET_NAME")) {
//            propertyCache.remove("HOOTSUITE_SHEET_NAME", true);
//        }
//                
//        
//        callback && callback();
//    }
//    
//}
//
///**
// * Sets target Calendar
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setTargetCalendar(callback) {
//    var sheetName,
//        propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET TARGET CALENDAR",
//            Utilities.formatString("Enter (or paste) the name of the target calendar.\n\n Or leave it empty to use your \ndefault calendar."),
//            ui.ButtonSet.OK
//        )
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        var sheetName = result.getResponseText().trim();
//        
//        if (sheetName !== "") propertyCache.put("CALENDAR_NAME", sheetName, true);
//        
//        if (sheetName === "" && propertyCache.get("CALENDAR_NAME")) {
//            propertyCache.remove("CALENDAR_NAME", true);
//        }
//                
//        
//        callback && callback();
//    }
//    
//}
//
///**
// * Sets Todoist Auth Token
// *
// * @param {Function} callback  - An optional callback to be invoked when the 
// *                               dialog recieves confirmation. Used to chain
// *                               multiple invocations.
// */
//function setTodoistAuthToken(callback) {
//    var propertyCache = new PropertyCache(),
//        ui = FormApp.getUi(),
//        result = ui.prompt(
//            "SET TODOIST AUTH TOKEN",
//            Utilities.formatString("Enter (or paste) the authorisation token for the Todoist account."),
//            ui.ButtonSet.OK
//        )
//    
//    if (result.getSelectedButton() === ui.Button.OK) {
//        
//        var token = result.getResponseText().trim();
//        
//        if (token !== "") propertyCache.put("TODOIST_AUTH_TOKEN", token, true);
//        
//        if (token === "" && propertyCache.get("TODOIST_AUTH_TOKEN")) {
//            propertyCache.remove("TODOIST_AUTH_TOKEN", true);
//        }
//                
//        
//        callback && callback();
//    }
//    
//    
//}

/** 
 * Initialize triggers.
 *
 * @param {Function} callback  - An optional callback to be invoked when the 
 *                               dialog recieves confirmation. Used to chain
 *                               multiple invocations.
 */
function initialize(callback) {
    
    var trigger,
        response,
        result,
        propertyCache = new PropertyCache(),
        triggerIDs = [
            propertyCache.get(UPDATE_FORM_CONTEXT_ID),
            propertyCache.get(FORM_SUBMIT_CONTEXT_ID)
        ].filter(function(id) {
            return id !== null;
        });
    
    // Fetch metadata info on form to determine if current user owns the form
    
    var url = Utilities.formatString("https://www.googleapis.com/drive/v3/files/%s?fields=ownedByMe", FormApp.getActiveForm().getId())
    var options = {
        "headers":{
            "Authorization":"Bearer " + ScriptApp.getOAuthToken()
        },
        "muteHttpExceptions": true
    }

    response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(response.getContentText())
    }
    
    result = JSON.parse(response);
    
    // Only allow owner of the form to create/destroy these triggers
    if (result.ownedByMe) {
        
        // delete existing "pollStaffSpreadsheet" and "onFormSubmit" triggers
        ScriptApp.getProjectTriggers().forEach(function(trigger) {
            //var trigger = ScriptApp.getProjectTriggers()[0];
            if(trigger.getHandlerFunction() === "updateForm" || trigger.getHandlerFunction() === "onFormSubmit") {
                
                ScriptApp.deleteTrigger(trigger);
                                        
            }
                
        });
                
        // creates a time-based trigger; trigger handler invokes polling on staff spreadsheet to
        // update names in form as well as updates date patterns
        // Note: Time intervals cannot be lower that one hour for add-ons. Script will throw an error otherwise.
        //       Alternative is to use an external cron service to trigger updates via web app.
        
        trigger = ScriptApp.newTrigger("updateForm").timeBased().everyHours(1).create();
        propertyCache.put(UPDATE_FORM_CONTEXT_ID, trigger.getUniqueId(), true);
        
        // create an event-based trigger for onFormSubmit
        trigger = ScriptApp.newTrigger("onFormSubmit").forForm(FormApp.getActiveForm().getId()).onFormSubmit().create();
        propertyCache.put(FORM_SUBMIT_CONTEXT_ID, trigger.getUniqueId(), true);
        
        updateForm(null);
    
    } else {
    
      FormApp.getUi().alert('Only the owner of the form can start the triggers required to process form submissions')
    }
        
    callback && callback();
}

/**
 * Updates the name and date range items on 
 * promotion request form.
 *
 * @param {Event} e - An event object. Set when the function is invoked from a trigger.
 */
function updateForm(e) {
    pollStaffSpreadsheet_();
    updateDatePatternsForTiers_();
}

/**
 * When setup is first run a time-based trigger is created that invokes this
 * function every hour. Note: The lowest interval allowed to add-ons is an hour.
 * Lower than that and will cause an error.
 */
function pollStaffSpreadsheet_() {

    // fetch metadata for spreadsheet
    var cached,
        propertyCache = new PropertyCache(),
        ssID = propertyCache.get("STAFF_SPREADSHEET_ID"),
        response = UrlFetchApp.fetch(
            Utilities.formatString("https://www.googleapis.com/drive/v3/files/%s?fields=modifiedTime,version", ssID),
            {
                "headers":{
                    "Authorization":"Bearer " + ScriptApp.getOAuthToken()
                }
            }
        ),
        result;
    
    result = JSON.parse(response);
    
    cached = propertyCache.get(ssID);
    
    // check if sheet was modified and if modified update form's name item.
    if (!cached || cached.modifiedTime !== result.modifiedTime || cached.version !== result.version ) {
    
        propertyCache.put(ssID, result);
        updateNamePatternForSponsors_();
        
    }
    
}

/**
 * Removes the custom properties set for pull ranges.
 * System will fallback on defaults.
 *
 */
function resetToDefaults() {
    var propertyCache = new PropertyCache(),
        staffSSID = propertyCache.get("STAFF_SPREADSHEET_ID");
    
    propertyCache.remove(staffSSID, true);
    propertyCache.remove("PULL_RANGE_FOR_USER_NAMES", true);
    propertyCache.remove("PULL_RANGE_FOR_USER_EMAILS", true);
    
}

/**
 * Updates the regex pattern for sponsor names in the 
 * corresponding form item.
 *
 */
function updateNamePatternForSponsors_() {
    var propertyCache = new PropertyCache(),
        ss = SpreadsheetApp.openById(propertyCache.get("STAFF_SPREADSHEET_ID")),
        nameTextItem = FormApp.getActiveForm().getItemById(FORM_NAME_ITEM_ID).asTextItem(),
        pattern = "",
        validationBuilder,
        values;
    
    nameTextItem.clearValidation();
    
    values = ss.getRange(propertyCache.get("PULL_RANGE_FOR_USER_NAMES") || DEFAULT_PULL_RANGE_FOR_USER_NAMES).getValues();
    
    pattern = "(" 
        + values.map(function(row) {
              row[0] = row[0].trim();    // trim first name of leading and trailing spaces
              row[1] = row[1].trim();    // trim last name of leading and trailing spaces
              
              return row.join(" ");              
          }).filter(function(item) {      
              
              return item !== " ";
          }).join("|") 
        + ")\\s*";
    
    validationBuilder = FormApp.createTextValidation().requireTextMatchesPattern(pattern);
    validationBuilder.setHelpText("Please enter your first and last name (case sensitive). You must be a CCN staff member to use this form.");
    nameTextItem.setValidation(validationBuilder.build());
}

/**
 * Updates start/end date regex patterns for each tier.
 *
 */
function updateDatePatternsForTiers_() {

    var form = FormApp.getActiveForm(),
        validationBuilder,
        dateTextItem,
        items = [
            {"id":FORM_BRONZE_START_DATE_ITEM_ID, "days":21, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm. Note: the deadline for a Bronze form is 3 weeks prior to the event start date. If your event does not begin within this 3 week window, please contact the Communications Director to discuss promotion."},
            {"id":FORM_BRONZE_END_DATE_ITEM_ID, "days":21, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm."},
            {"id":FORM_SILVER_START_DATE_ITEM_ID, "days":42, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm. Note: the deadline for a Silver form is 6 weeks prior to the event start date. If your event does not begin within this 6 week window, please click 'Back' and choose a form with a shorter deadline."},
            {"id":FORM_SILVER_END_DATE_ITEM_ID, "days":42, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm."},
            {"id":FORM_GOLD_START_DATE_ITEM_ID, "days":70, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm. Note: the deadline for a Gold form is 10 weeks prior to the event start date. If your event does not begin within this 10 week window, please click 'Back' and choose a form with a shorter deadline."},
            {"id":FORM_GOLD_END_DATE_ITEM_ID, "days":70, "helpText":"Use M/D/YY HH:MM am/pm format, i.e. 1/1/18 1pm."}
        ];
     
    items.forEach(function(item, index) {
         var pattern = getRegexFor_(item.days),
             dateTextItem = form.getItemById(item.id).asTextItem(),
             validationBuilder = FormApp.createTextValidation().requireTextMatchesPattern(pattern);
         
         dateTextItem.clearValidation();
         
         validationBuilder.setHelpText(item.helpText);
           
         dateTextItem.setValidation(validationBuilder.build());
    });

}

/**
 * Produces a regex pattern to match a span of days from today. 
 *
 * @param {Number} days - Span of days to offset from
 *
 * @return {String} A regex pattern to match valid dates
 */
function getRegexFor_(days) {

    var struct = {},
    
        date = new Date(Date.now() + days * MILLISECONDS_IN_A_DAY),
        
        year = date.getFullYear(),
        month = date.getMonth(),
        date = date.getDate(),
        
        nonLeapYears = [],
        
        patterns = [],
        
        endYear = year + UPPER_BOUND_YEAR_OFFSET;
    
    // build struct
    struct[year] = {};
    buildSubStruct_.call(struct[year], year, month, date, 11, 31);
    
    patterns.push("(" + generateDateRegexFromStruct_(struct) + ")");
    
    year++;
    
                     
    for(;year <= endYear; year++) {
       
       if (!isLeapYear(year)) {
       
           nonLeapYears.push(year);
           nonLeapYears.push(year.toString().slice(-2));
           
       } else {
       
           patterns.push(
               "(((0?2\/(0?[1-9]|1[0-9]|2[0-9]))|((0?[13578]|1[02])\/(0?[1-9]|1[0-9]|2[0-9]|3[01]))|((0?[469]|11)\/(0?[1-9]|1[0-9]|2[0-9]|30)))\/(" + year + "|" + year.toString().slice(-2) + "))"
           );
       
       }
        
    }
    
    patterns.push(
        "(((0?2\/(0?[1-9]|1[0-9]|2[0-8]))|((0?[13578]|1[02])\/(0?[1-9]|1[0-9]|2[0-9]|3[01]))|((0?[469]|11)\/(0?[1-9]|1[0-9]|2[0-9]|30)))\/(" + nonLeapYears.join("|") + "))"    
    );
    
    return "(" + patterns.join("|") + ") " + TIME_PATTERN_12_HOUR;   
}

/**
 * Helper function. 
 * Builds sub structures on month boundary dates
 * within a given year.
 *
 * @param {Number} year        - 4-digit year
 * @param {Number} _startMonth - Start month zero-indexed (0 - 11)
 * @param {Number} _startDate  - Start month date (1 - 31)
 * @param {Number} _endMonth   - End month zero-indexed (0 - 11)
 * @param {Number} _endDate    - End month date (1 - 31)
 */
function buildSubStruct_(year, _startMonth, _startDate, _endMonth, _endDate) {

	var obj,
		stack = new Array(12),
		
		self,
		
		lastMonth = _endMonth,
		lastDateOfMonth = new Date(year, _endMonth, _endDate); 
	
	this[_startMonth] = {};
	this[_startMonth]["start"] = _startDate;
	
	
	while (_startMonth !== lastMonth) {
	
		obj = {};
		obj.start = 1;
		obj.end = lastDateOfMonth.getDate();
		
		stack[lastMonth] = obj;
		
		lastDateOfMonth = new Date((new Date(year, lastMonth, 1)).getTime() - MILLISECONDS_IN_A_DAY);
		lastMonth = lastDateOfMonth.getMonth();
	}
	
	self = this;
	
	// use stack to populate struct; this maintains insertion order
	stack.forEach(function(item, month){
	
		if (item) {
		
		  self[month] = item;
		  
		}
		
	});
	
	this[_startMonth]["end"] = lastDateOfMonth.getDate();
		
}

/**
 * Helper function.
 * Generates a regular expression (from a structured source)
 * to validate date inputs.
 *
 * @param {Object} struct - An object containing dates in a tree-like data structure
 */
function generateDateRegexFromStruct_(struct) {
    var yearSubPattern,
        yearSubPatternArray,
        
        monthSubPattern,
        monthSubPatternArray,
        
        dateSubPatternArray,
        
        range,
        sub;
    
    yearSubPatternArray = [];
    
    for(var year in struct) {
        
        yearSubPattern = "((";
        monthSubPatternArray = [];
        
        for( var month in struct[year]) {
        
            sub = "" + (parseInt(month) + 1);
            
            monthSubPattern = "(" + (sub.length === 1 ? ("(" + sub +"|0" + sub + ")") : sub) + "\/";
            dateSubPatternArray = [];
            
            range = struct[year][month];
                
            for (var i = range.start; i <= range.end; i++) {
                
                dateSubPatternArray.push((i > 9 ? i : ("(" + i + "|0" + i + ")")));
                
            }
            
            monthSubPattern += "(" + dateSubPatternArray.join("|") +"))";
            monthSubPatternArray.push(monthSubPattern);
        }
        
        yearSubPattern += monthSubPatternArray.join("|") + ")\/("+ year + "|" + year.toString().slice(-2) + "))";
        yearSubPatternArray.push(yearSubPattern);
        
    }
    
    return yearSubPatternArray.join("|");
    
}

/**
 * Determines if a given year is a leap year.
 *
 * @returns {Boolean} True if year parameter is a leap year; false otherwise.
 */
function isLeapYear(year) {
    return year % 4 === 0;
}

/**
 * OnFormSubmit trigger handler. 
 * 
 * @param {Event} e - An event object.
 */
function onFormSubmit(e) {

    Log_(e);
    
//    var propertyCache = new PropertyCache(),
        
    var hootsuiteSSID = Config.get("HOOTSUITE_SPREADSHEET_ID"),
//        hootsuiteSheetName = (propertyCache.get("HOOTSUITE_SHEET_NAME") || DEFAULT_HOOTSUITE_SHEET_NAME),    
        hootsuiteSheetName = DEFAULT_HOOTSUITE_SHEET_NAME,
        
        hootsuiteSheet = SpreadsheetApp.openById(hootsuiteSSID).getSheetByName(hootsuiteSheetName),
        hootsuiteSheetID = hootsuiteSheet.getSheetId(),
        
        responseSSID = Config.get("RESPONSES_SPREADSHEET_ID"),
        responseSheetName = (propertyCache.get("RESPONSE_SHEET_NAME") || DEFAULT_RESPONSE_SHEET_NAME),
        
        responseSheet = SpreadsheetApp.openById(responseSSID).getSheetByName(responseSheetName),
        responseSheetID = responseSheet.getSheetId(),
        
        staffSSID = Config.get("STAFF_SPREADSHEET_ID"),
        staffNameRange = (propertyCache.get("PULL_RANGE_FOR_USER_NAMES") || DEFAULT_PULL_RANGE_FOR_USER_NAMES),
        staffEmailRange = (propertyCache.get("PULL_RANGE_FOR_USER_EMAILS") || DEFAULT_PULL_RANGE_FOR_USER_EMAILS),
                
        formResponse = e.response,
        form = e.source,
        
        rowData,
        
        schema = [
			{"id":FORM_TIER_SELECTION_ITEM_ID, "colIndex":6},
			{"id":FORM_NAME_ITEM_ID, "colIndex":4},
			{"colIndex":5}, // email
			{"colIndex":7, "postfix":"_START_DATE_ITEM_ID"}, // event start
			//{"colIndex":8}, // timestamp
			{"colIndex":9, "postfix":"_EVENT_TITLE_ITEM_ID"},
			{"colIndex":10, "postfix":"_END_DATE_ITEM_ID"}, // event end
			{"colIndex":11, "postfix":"_EVENT_LOCATION_ITEM_ID"},
			{"colIndex":12, "postfix":"_EVENT_COST_ITEM_ID"},
			{"colIndex":13, "postfix":"_EVENT_REGISTRATION_TYPE_ITEM_ID"},// registration type
			{"colIndex":14}, // registration location
            {"colIndex":15}, // registration deadline          
			{"colIndex":16, "postfix":"_EVENT_FOR_ITEM_ID"},
			{"colIndex":17, "postfix":"_EVENT_ABOUT_ITEM_ID"},
			{"colIndex":18, "postfix":"_EVENT_REQUESTED_ASSETS_ITEM_ID"},
			{"colIndex":19, "postfix":"_OTHER_NOTES_ITEM_ID"}
        ],
        
        validationResult,
        responseTimestamp = toExcelSerialNumberFormat(formResponse.getTimestamp()),
        response,
        name,
        type,
        tier,
        
        sponsorEmail,
        
        templateData = {
            "event":{
                "mediaRequests":[]
            },
            "sponsor":{},
            "registration":{}
        },
        
        eventData = {
            "options":{}
        },
        
        registrationURL,
        registrationLocation,
        registrationDeadline,
        
        timezoneOffset = -(new Date().getTimezoneOffset()),
        timezoneOffsetFormatted = (timezoneOffset < 0 ? "-" : "+")
                                + ("00" + Math.abs(Math.floor(timezoneOffset/60)).toFixed()).slice(-2)
                                + ":"
                                + ("00" + Math.abs(timezoneOffset % 60)).slice(-2),
        
        registrationDeadlinePostfix = "T00:00:00" + timezoneOffsetFormatted;
        
        
    schema.forEach(function(item) {
        
        if (item.postfix) item.id = eval("FORM_" + tier + item.postfix);
        
        if (item.id) {
            
            var nextItem = form.getItemById(item.id)
            var itemResponse = formResponse.getResponseForItem(nextItem)
            
            if (itemResponse !== null) {
            
              response = itemResponse.getResponse();
              type = nextItem.getType(); 
              
            } else {
            
              return;
            }
        }
        
        switch(item.colIndex) {
            case 6:    // Set Tier
                tier = item.value = response.toUpperCase().match(/(GOLD|SILVER|BRONZE)/i)[0];
                templateData.event.tier = response;
                templateData.event.confirmationWeekCount = tier === "GOLD" ? "Ten" : tier === "SILVER" ? "Six" : "Three"; 
                break;
                
            case 4:    // Set Name
                name = item.value = response.trim();
                templateData.sponsor.name = name;
                break;
                
            case 5:    // Set Email
                var rowIndex = Sheets.Spreadsheets.Values.get(staffSSID, staffNameRange).values.map(function(row){
                    return row[0] + " " + row[1];
                }).indexOf(name);
                                
                item.value = Sheets.Spreadsheets.Values.get(staffSSID, staffEmailRange).values[rowIndex].join();
                
                templateData.sponsor.email = item.value;
                sponsorEmail = item.value;
                break;
                
            case 7:   // Start date
            case 10:  // End date                
                var match = response.match(/(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})(?: at| At| AT)? (\d{1,2})(\:(\d{2}))?/),
                    month = match[1],
                    day = match[2],
                    year = (match[3].length === 2) ? "20" + match[3] : match[3],
                      hours = ((response.slice(-2).toLowerCase() === "pm") ? "" + ((parseInt(match[4]) % 12) + 12) : (match[4] % 12)),
                    minutes = (match.length == 7 && match[6]) ? match[6] : "00",
                    date;
                
                date = new Date(
                        Utilities.formatString(
                                "%s-%s-%sT%s:%s:00", 
                                year, 
                                ("0" + month).slice(-2), 
                                ("0" + day).slice(-2), 
                                ("0" + hours).slice(-2), 
                                ("0" + minutes).slice(-2)
                        )
                );
                
                item.value = toExcelSerialNumberFormat(date); // Date must be in UTC time (no offsets)
                
                if (item.colIndex === 7) {
                    eventData.startTime = date;
                    templateData.event.startTime = Utilities.formatDate(date, Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT);
                    templateData.event.startTimeShort = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM.dd");
                }
                
                if (item.colIndex === 10) {
                    eventData.endTime = date;
                    templateData.event.endTime = Utilities.formatDate(date, Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT);
                }
                
                break;
            
            //case 8: // Timestamp
            //    item.value = toExcelSerialNumberFormat(formResponse.getTimestamp());
            //    break;
            
            case 9: // title
                item.value = response;
                eventData.title = response;
                templateData.event.title = response;
                break;
            
            case 11: // location
                item.value = response;
                eventData.options.location = response;
                templateData.event.location = response;
                break;
            
            case 12:// cost
                item.value = response;
                templateData.event.cost = response;
                break;
            
            case 13:
                item.value = response;
                
                switch(response) {
                    
                    case "Register online":
                        // Grab registration url and deadline                        
                        registrationURL = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_ONLINE_REGISTRATION_URL_ITEM_ID"))).getResponse();
                        registrationDeadline = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_ONLINE_REGISTRATION_DEADLINE_ITEM_ID"))).getResponse();                       
                        
                        templateData.registration.url = registrationURL;
                        templateData.registration.deadlineDate = (registrationDeadline === "") ? null :new Date(registrationDeadline + registrationDeadlinePostfix);
                        templateData.registration.deadline = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT);
                        templateData.registration.deadlineShort = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT_SHORT);
                        templateData.registration.type = "ONLINE";
                        break;
                        
                    case "Register in person":
                        // Grab registration location and deadline
                        registrationLocation = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_OFFLINE_WHERE_CAN_ATTENDEES_REGISTER_ITEM_ID"))).getResponse();
                        registrationDeadline = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_OFFLINE_REGISTRATION_DEADLINE_ITEM_ID"))).getResponse();
                        
                        templateData.registration.deadlineDate = (registrationDeadline === "") ? null :new Date(registrationDeadline + registrationDeadlinePostfix);
                        templateData.registration.deadline = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT);
                        templateData.registration.deadlineShort = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT_SHORT);
                        templateData.registration.type = "OFFLINE";
                        break;
                        
                    case "Register online - OR - in person":
                        // Grab registration url, location and deadline
                        registrationURL = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_ONLINEOFFLINE_REGISTRATION_URL_ITEM_ID"))).getResponse();
                        registrationLocation = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_ONLINEOFFLINE_WHERE_CAN_ATTENDEES_REGISTER_ITEM_ID"))).getResponse();
                        registrationDeadline = formResponse.getResponseForItem(form.getItemById(eval("FORM_" + tier + "_EVENT_ONLINEOFFLINE_REGISTRATION_DEADLINE_ITEM_ID"))).getResponse();
                        
                        templateData.registration.url = registrationURL;
                        templateData.registration.deadlineDate = (registrationDeadline === "") ? null :new Date(registrationDeadline + registrationDeadlinePostfix);
                        templateData.registration.deadline = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT);
                        templateData.registration.deadlineShort = (registrationDeadline === "") ? null : Utilities.formatDate(new Date(registrationDeadline + registrationDeadlinePostfix), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT_SHORT);
                        templateData.registration.type = "BOTH";
                        break;
                    
                    case "No registration required":
                        templateData.registration.type = "NONE";
                        break;
                }
                
                break;
            
            case 14: // registration location
                item.value = registrationLocation || "N/A";
                templateData.registration.location = item.value;
                break;
                
            case 15: // registration deadline
                item.value = (registrationDeadline === undefined || registrationDeadline === "") ? "NONE" : toExcelSerialNumberFormat(new Date(registrationDeadline + registrationDeadlinePostfix));
                break;
                
            case 16: // event for
                item.value = response;
                templateData.event.targetAudience = response;
                break;
                
            case 17: // event about
                item.value = response;
                templateData.event.about = response;
                break;
                         
            default:
                item.value = (type === FormApp.ItemType.CHECKBOX) ? response.map(function(option) { return option.replace(/\[.+\]/g, "").trim(); }).join(", ") : response;
                
                if (item.colIndex === 18 ) {
                    templateData.event.mediaRequests.push(item.value);
                }
                
                break;
               
        }
        
    });
    
    // Validate response; sends an error report if response has invalid answers
    validationResult = validateFormResponse(Object.assign({}, templateData, eventData));
    
    if (validationResult.errors.length) {
    
        sendErrorReportEmailToRespondent(
            sponsorEmail,
            Object.assign(
                validationResult, 
                templateData,
                eventData,
                {
                    "submitTime":Utilities.formatDate(formResponse.getTimestamp(), Session.getScriptTimeZone(), REGISTRATION_DEADLINE_FORMAT),
                    "editResponseUrl":formResponse.getEditResponseUrl()
                }
            )
        );
        
        return;
    }
    
    // sort on column index
    schema = schema.sort(function(a, b){
        return a.colIndex - b.colIndex;
    });
    
    // Response sheet update
    
    rowData = schema.reduce(
        function(rowData, item) {
            var cellData = Sheets.newCellData();
            
            rowData.values = rowData.values || [];
            
            cellData.userEnteredValue = Sheets.newExtendedValue();
            cellData.userEnteredValue[(typeof item.value) + "Value"] = item.value;
            
            rowData.values.push(cellData);
            
            return rowData;
        }, 
        Sheets.newRowData()
    );
    
    var responseSheetMaxRows = responseSheet.getMaxRows();
    
    PropertiesService.getDocumentProperties().setProperty("rowData", JSON.stringify(rowData));
    
    Sheets.Spreadsheets.batchUpdate(
        {
            "requests":[
                {
                    "insertDimension": {
                        "inheritFromBefore":true,
                        "range":{
                            "dimension":"ROWS",
                            "sheetId":responseSheetID,
                            "startIndex": responseSheetMaxRows - 1,
                            "endIndex": responseSheetMaxRows
                        }
                    }
                }, 
                {
                    "updateCells": {
                        "fields":"userEnteredValue.stringValue,userEnteredValue.numberValue",
                        "range": {
                            "sheetId":responseSheetID,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                            "startRowIndex": responseSheetMaxRows - 1,
                            "endRowIndex": responseSheetMaxRows
                        },
                        "rows":[
                            {
                                "values":[
                                    {
                                        "userEnteredValue":{
                                            "numberValue":responseTimestamp
                                        }
                                    }
                                ]
                            }
                        ]
                    }
                },
                {
                    "updateCells": {
                        "fields":"userEnteredValue.stringValue,userEnteredValue.numberValue",
                        "range": {
                            "sheetId":responseSheetID,
                            "startColumnIndex": 5,
                            "endColumnIndex": 20,
                            "startRowIndex": responseSheetMaxRows - 1,
                            "endRowIndex": responseSheetMaxRows
                        },
                        "rows":[
                            rowData
                        ]
                    }
                }
            ]
        },
        responseSSID
    );
    
    
    // HootSuite sheet Update
    var hootSuiteRows = [],
        numEntries = tier === "GOLD" ? 3 :
                     tier === "SILVER" ? 2 :
                     1;
        
    for(var i = 1; i <= numEntries; i++) {
        
        var rowData = Sheets.newRowData(),
            cellData = Sheets.newCellData();
        
        rowData.values = [];
        
        cellData.userEnteredValue = Sheets.newExtendedValue();
        cellData.userEnteredValue.numberValue = toExcelSerialNumberFormat(new Date(eventData.startTime.getTime() - (7 * MILLISECONDS_IN_A_DAY * i)));
            
        rowData.values.push(cellData);
        
        cellData = Sheets.newCellData();
        cellData.userEnteredValue = Sheets.newExtendedValue();
        cellData.userEnteredValue.stringValue = templateData.event.about; // + " - " + templateData.event.startTime;
        
        rowData.values.push(cellData);
        
        cellData = Sheets.newCellData();
        cellData.userEnteredValue = Sheets.newExtendedValue();
        cellData.userEnteredValue.stringValue = (registrationURL)? registrationURL : "";
        
        rowData.values.push(cellData);
        
        hootSuiteRows.push(rowData);
    }
    
    var hootsuiteSheetMaxRows = hootsuiteSheet.getMaxRows();
    
    Sheets.Spreadsheets.batchUpdate(
        {
            "requests":[
                {
                    "insertDimension":{
                        "inheritFromBefore":true,
                        "range":{
                            "dimension":"ROWS",
                            "sheetId":hootsuiteSheetID,
                            "startIndex":hootsuiteSheetMaxRows,
                            "endIndex":hootsuiteSheetMaxRows + hootSuiteRows.length
                        }
                    }
                },
                {
                    "updateCells": {
                        "fields":"userEnteredValue.stringValue,userEnteredValue.numberValue",
                        "range": {
                            "sheetId":hootsuiteSheetID,
                            "startColumnIndex": 0,
                            "endColumnIndex": 3,
                            "startRowIndex": hootsuiteSheetMaxRows,
                            "endRowIndex": hootsuiteSheetMaxRows + hootSuiteRows.length
                        },
                        "rows": hootSuiteRows
                    }
                }
            ]
        }, 
        hootsuiteSSID
    );
       
    sendConfirmationEmail(sponsorEmail, templateData);
    
    eventData.calendarName = Config.get("GOOGLE_CALENDAR_NAME");
    createCalendarEvent(eventData, templateData);
    
    emailResourceTeamLeader(templateData);

    if (TEST_USE_TODOIST) {

      var todoistConfig = {    
          spreadsheetId:     responseSSID,
          rowNumber:         responseSheetMaxRows,
          token:             Config.get("TODOIST_AUTH_TOKEN"),
          taskTemplateId:    Config.get("TODOIST_TASKS_TEMPLATE_ID"),
          commentTemplateId: Config.get("TODOIST_COMMENT_TEMPLATE_ID"),
          staffSheetId:      Config.get("STAFF_DATA_GSHEET_ID"),
          properties:        PropertiesService.getDocumentProperties(), 
          lock:              LockService.getDocumentLock()
      }
  
      Todoist.onFormSubmit(todoistConfig);
    }
    
    if (TEST_CHECK_PROMOTION_CALENDAR_) {
      checkPromotionCalendar_(e);
    }
    
} // onFormSubmit()

/**
 * Email confrimation of reciept with event summary.
 *
 * @param {String} recipientEmail - Recipient's email
 * @param {Object} templateData   - Data to be passed to email template
 */
function sendConfirmationEmail(recipientEmail, templateData) {
    var template = HtmlService.createTemplateFromFile('confirmation_email_template');
    
    Object.assign(template, templateData);
    
    MailApp.sendEmail(
        {
            "to":recipientEmail,
            "subject": templateData.event.tier + ' request received for [ ' + templateData.event.startTimeShort + ' ] ' + templateData.event.title,
            "htmlBody":template.evaluate().getContent()
        }
    );
    
}

/**
 * Create calendar event based on a event data object with
 * the following properties:
 *
 *          - calendarName
 *          - title
 *          - startTime
 *          - localStartTime
 *          - endTime
 *          - localEndTime
 *          - options
 *              - description
 *              - location
 *              - guests
 *              - sendInvites
 *
 * @param {Object} eventData - Event information  
 * @param {Object} templateData - Template data
 */
function createCalendarEvent(eventData, templateData) {

    var template = HtmlService.createTemplateFromFile('calendar_event_info_template'),
        options = Object.assign({}, eventData.options),
        calendar = (eventData.calendarName && CalendarApp.getCalendarsByName(eventData.calendarName)[0]) || CalendarApp.getDefaultCalendar();

    Object.assign(template, templateData);
    options.description = template.evaluate().getContent();
    
    // calendar = CalendarApp.getCalendarsByName("")[0];
    calendar.createEvent(
        eventData.title, 
        eventData.startTime, 
        eventData.endTime,
        options
    );
    
}

/**
 * Notifies resource team leader.
 *
 * @param {Object} templateData   - Data to be passed to email template
 */
function emailResourceTeamLeader(templateData) {

    var template = HtmlService.createTemplateFromFile('email_resource_team_leader_template'),
//        propertyCache = new PropertyCache(),
        ss = SpreadsheetApp.openById(Config.get("STAFF_DATA_GSHEET_ID"));
    
    Object.assign(template, templateData);
    
    ss.getDataRange().getValues().some(function(row) {
        var state = false;
        
        if (row[11] === "Resource Team" && row[12] === "Yes") {
        
            MailApp.sendEmail(
                {
                    "to":row[8],
                    "subject":'Attention: ' + templateData.event.title,
                    "htmlBody":template.evaluate().getContent()
                }
            );
            
            state = true;
            
        }
        
        return state;
    }); 
}

/**
 * Validates a form response for promotion request.
 *
 * @param data {Object} Data to validate
 */
function validateFormResponse(data) {
    var errors = [],
        startTime = data.startTime.getTime();
    
    (data.endTime.getTime() - startTime <= 0) && errors.push("End date/time of event cannot precede start date/time");
    (data.registration.type !== "NONE") && !!data.registration.deadlineDate && (startTime - data.registration.deadlineDate.getTime() < 0) && errors.push("Registration due date cannot exceed event date");
    
    return {"errors":errors};
}

/**
 * Sends an error report email to respondent with instructions
 * to edit and resubmit.
 *
 * @param recipientEmail {String} Sponsor/recipient email
 * @param data {Object} Data to inject into the email template
 */
function sendErrorReportEmailToRespondent(recipientEmail, data) {
    var template = HtmlService.createTemplateFromFile('error_report_template');
    
    Object.assign(template, data);
    
    MailApp.sendEmail(
        {
            "to":recipientEmail,
            "subject": "Action required!",
            "htmlBody":template.evaluate().getContent()
        }
    );   
}

/**
 * Converts a given date to Excel serial number format.
 *
 * @param {Date} date - A date object
 */
function toExcelSerialNumberFormat(date) {

    return 25569.0 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
  
}