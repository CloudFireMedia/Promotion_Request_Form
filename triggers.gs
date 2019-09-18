/** Trigger implementations and related callbacks */

/** Force enable drive scope */
/** Drive.Files.list(); */  

function onInstall() {
  
  var propertyCache = new PropertyCache;
  
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
  
  Log_('Installed PRF');
}

/** Shows all of the cache settings */
function showCache() {
    var propertyCache = new PropertyCache(),

        prompt =
            "Staff Spreadsheet ID: "             + Config.get("STAFF_DATA_GSHEET_ID") + "\n" +
            "Responses Spreadsheet ID: "         + Config.get("PROMOTION_FORM_RESPONSES_GSHEET_ID") + "\n" +
            "Responses Sheet name: "             + DEFAULT_RESPONSE_SHEET_NAME + "\n" +
            "Pull range for users names: "       + DEFAULT_PULL_RANGE_FOR_USER_NAMES + "\n" +
            "Pull range for users emails: "      + DEFAULT_PULL_RANGE_FOR_USER_EMAILS + "\n" +
            "Hootsuite Spreadsheet ID: "         + Config.get("HOOTSUITE_SPREADSHEET_ID") + "\n" +
            "Hootsuite Sheet Name: "             + DEFAULT_HOOTSUITE_SHEET_NAME + "\n" +
            "Calendar Name: "                    + Config.get("GOOGLE_CALENDAR_NAME") + "\n" +
            "Todoist Auth token: "               + Config.get("TODOIST_AUTH_TOKEN") + "\n" +
            "Todoist Tasks Template CSV ID: "    + Config.get("TODOIST_TASKS_TEMPLATE_ID") + "\n" +
            "Todoist Comment Template GDoc ID: " + Config.get("TODOIST_COMMENT_TEMPLATE_ID")

    FormApp.getUi().alert(prompt)
}

function initialised(e) {
    var alreadyInitialised = false;
    
    if (e !== undefined && e.hasOwnProperty('authMode') && e.authMode !== ScriptApp.AuthMode.NONE) {

        var propertyCache = new PropertyCache,
            updateFormTriggerId = propertyCache.get(UPDATE_FORM_CONTEXT_ID),
            formSubmitTriggerId = propertyCache.get(FORM_SUBMIT_CONTEXT_ID);
        
        if (updateFormTriggerId !== null && formSubmitTriggerId !== null) {
          return true;
        }       
    }

    return alreadyInitialised;
}

function clearConfig() {
  var propertyCache = new PropertyCache,
      updateFormTriggerId = propertyCache.remove(UPDATE_FORM_CONTEXT_ID, true),
      formSubmitTriggerId = propertyCache.remove(FORM_SUBMIT_CONTEXT_ID, true);
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
        ssID = Config.get("STAFF_DATA_GSHEET_ID"),
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
        updateNamePatternForSponsors_(ssID);
    }
    
}

/**
 * Updates the regex pattern for sponsor names in the 
 * corresponding form item.
 *
 */
function updateNamePatternForSponsors_(ssID) {
    var ss = SpreadsheetApp.openById(ssID),
        nameTextItem = FormApp.getActiveForm().getItemById(FORM_NAME_ITEM_ID).asTextItem(),
        pattern = "",
        validationBuilder,
        values;
    
    nameTextItem.clearValidation();
    
    values = ss.getRange(DEFAULT_PULL_RANGE_FOR_USER_NAMES).getValues();
    
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
    Log_(SCRIPT_VERSION);
    
    var hootsuiteSSID = Config.get("HOOTSUITE_SPREADSHEET_ID"),
        hootsuiteSheet = SpreadsheetApp.openById(hootsuiteSSID).getSheetByName(DEFAULT_HOOTSUITE_SHEET_NAME),
        hootsuiteSheetID = hootsuiteSheet.getSheetId(),  
        
        responseSSID = Config.get("PROMOTION_FORM_RESPONSES_GSHEET_ID"),
        responseSheet = SpreadsheetApp.openById(responseSSID).getSheetByName(DEFAULT_RESPONSE_SHEET_NAME),
        responseSheetID = responseSheet.getSheetId(),
        
        staffSSID = Config.get("STAFF_DATA_GSHEET_ID"),
                
        formResponse = e.response,
        form = e.source,
        
        rowData,
        responseSheetMaxRows,
        
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
                var staffValues = Sheets.Spreadsheets.Values.get(staffSSID, DEFAULT_PULL_RANGE_FOR_USER_NAMES).values;
                Log_(staffValues);
            
                var rowIndex = staffValues.map(function(row){
                    return row[0] + " " + row[1];
                }).indexOf(name);
                                
                Log_('rowIndex: ' + rowIndex);            
                                
                if (rowIndex === -1) {
                  throw new Error('Could not find staff member "' + name + "'");
                }
                                
                item.value = Sheets.Spreadsheets.Values.get(staffSSID, DEFAULT_PULL_RANGE_FOR_USER_EMAILS).values[rowIndex].join();
                
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
    
    updateResponseSheet();
    
    updateHootSuite();
           
    sendConfirmationEmail(sponsorEmail, templateData);
    
    eventData.calendarName = Config.get("GOOGLE_CALENDAR_NAME");
    createCalendarEvent(eventData, templateData);
    
    emailResourceTeamLeader_(templateData);

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
        
        Log_(todoistConfig);
    }
    
    checkPromotionCalendar_(e);    
    Log_('Finished processing form submission');
    
    return;
    
    // Private Functions
    // -----------------

    function updateResponseSheet() {
    
        rowData = schema.reduce(
            function(rowData, item) {
                var cellData = Sheets.newCellData();
                
                rowData.values = rowData.values || [];
                
                cellData.userEnteredValue = Sheets.newExtendedValue();
                
                if (typeof item.value === 'object') {
                  // Assume the object is an array
                  cellData.userEnteredValue["stringValue"] = item.value[0];                
                } else {
                  cellData.userEnteredValue[(typeof item.value) + "Value"] = item.value;
                }
                
                rowData.values.push(cellData);
                
                return rowData;
            }, 
            Sheets.newRowData()
        );
        
        responseSheetMaxRows = responseSheet.getMaxRows();
        
        PropertiesService.getDocumentProperties().setProperty("rowData", JSON.stringify(rowData));
        
        var updateValues = {
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
        };
        
        Log_(updateValues)

        Sheets.Spreadsheets.batchUpdate(updateValues, responseSSID);
        
        Log_('Written to response sheet "' + responseSSID + '"');
        
    } // onFormSubmit.updateResponseSheet()

    function updateHootSuite() {
    
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
            cellData.userEnteredValue.stringValue = registrationURL ? registrationURL : "";
            
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
        
        Log_('Written to hootsuite sheet "' + hootsuiteSSID + '"');
            
    } // onFormSubmit.updateHootSuite()

} // onFormSubmit()

/**
 * Email confrimation of reciept with event summary.
 *
 * @param {String} recipientEmail - Recipient's email
 * @param {Object} templateData   - Data to be passed to email template
 */
function sendConfirmationEmail(recipientEmail, templateData) {

    if (!TEST_SEND_EMAILS) {
        return;
    }

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
function emailResourceTeamLeader_(templateData) {

    if (!TEST_SEND_EMAILS) {
        return;
    }

    var template = HtmlService.createTemplateFromFile('email_resource_team_leader_template'),
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

    if (!TEST_SEND_EMAILS) {
        return;
    }

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