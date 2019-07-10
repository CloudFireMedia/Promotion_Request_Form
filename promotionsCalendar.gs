/**
 * On a form submission check every response against every entry in the 
 * the promotions calendar to see if there are any matches. If so record 
 * in the calendar GSheet that a promotion request has been made for this event.
 */

function checkPromotionCalendar_(event) {

  if (!TEST_CHECK_PROMOTION_CALENDAR) {
    return;
  }

  // Can't work out why this won't work!! Something is corrupting "e", although it is still an object. 
//  var isRunFromTrigger = event.hasOwnProperty('triggerUid'); 

  var isRunFromTrigger = false

  if (typeof event === 'object' && event.triggerUid !== undefined) {
    isRunFromTrigger = true
  }

  var responsesSpreadsheet = SpreadsheetApp.openById(Config.get('PROMOTION_FORM_RESPONSES_GSHEET_ID'));
  
  if (responsesSpreadsheet === null) {
    throw new Error('Can not find the responses spreadsheet');
  }
  
  var responseSheet = responsesSpreadsheet.getSheetByName(DATA_SHEET_NAME);
  
  if (responseSheet === null) {
    throw new Error('Unable to find sheet named "' + DATA_SHEET_NAME + '".');
  }
  
  var responseSheetValues = responseSheet.getDataRange().getValues();
  
  var promotionDeadlinesCalendarId = Config.get('PROMOTION_DEADLINES_CALENDAR_ID');
  var calendarSS = SpreadsheetApp.openById(promotionDeadlinesCalendarId); 
  
  if (!calendarSS) {
  
    throw new Error (
      'Unable to find CCN Events Promotion Calendar spreadsheet. ' + 
      'Expected to find GSheet "' + promotionDeadlinesCalendarId + '"');
  }
  
  var calendarSheet = calendarSS.getSheetByName(PROMOTION_CALENDAR_SHEET_NAME); 
  
  if (!calendarSheet) {
  
    throw new Error (
      'Unable to find sheet named "' + PROMOTION_CALENDAR_SHEET_NAME + '" ' + 
      'on ' + calendarSS.getName() + ' spreadsheet');      
  }
  
  var calendarValues = calendarSheet.getDataRange().getValues();
  var columnNumbers = Utils.getPRFColumns(responsesSpreadsheet); 

  var errors = [];
  var warnings = [];
  var recordsFound = [];
  
  var startDateColumnIndex = columnNumbers.EventStart - 1
  var titleColumnIndex = columnNumbers.EventTitle - 1
  
  var matchThreshold = getConfig('MATCH_THRESHOLD')
  var maxEventDayDiff = getConfig('MAX_EVENT_DATE_DIFF')

  for (var i = 1; i < responseSheetValues.length; i++){ //i==1 to skip header row
  
    // check for "errors" first
    var eventDate = responseSheetValues[i][startDateColumnIndex];
    
    if (!eventDate) {
      errors.push('No event date found, line ' + (i+1) + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
      continue;
    }
    
    if (!eventDate instanceof Date){
      errors.push('Date ' + eventDate + ' on row ' + (i+1) + ' of ' + DATA_SHEET_NAME + ' sheet is not a valid date.  Skipping this row.');
      continue;
    }
    
    var eventTitle = responseSheetValues[i][titleColumnIndex].trim();
    
    if (!eventTitle) {
      warnings.push('No event title found, line ' + (i+1) + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
      continue;
    }
    
    // find matching event on calendar sheet
    
    //j=3 skips header rows 0-2 - they aren't frozen and we can't count on them staying set since humans are involved so we just skip three
    for (var j = 3; j < calendarValues.length; j++){ 
    
      //check for "errors" first
      var calendarEventDate = calendarValues[j][3];//column 4 is SHORT START DATE
      
      if (!calendarEventDate) {
        errors.push('No event date found, line ' + (j+1) + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
        continue;
      }
      
      if (!(calendarEventDate instanceof Date)) {
        errors.push('Date ' + calendarEventDate + ' on row ' + (j+1) + ' of ' + DATA_SHEET_NAME + ' sheet is not a valid date.  Skipping this row.');
        continue;
      }
      
      var calendarEventTitle = calendarValues[j][4].trim();//column 5 is EVENT TITLE
      
      if (!calendarEventTitle){
        warnings.push('No event title found, line ' + (j+1) + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
        continue;
      }
      
      // compare eventTitle and calendarEventTitle if dates are within the allowed range
      var dayDiff = Utils.DateDiff.inDays(calendarEventDate, eventDate);
      
      if (dayDiff <= maxEventDayDiff) {
      
        var matchPercent = getMatchValue(calendarEventTitle, eventTitle);
      
        if (matchPercent !== null && matchPercent[0][0] > matchThreshold) {
        
          recordsFound.push([
            'Title on this spreadsheet: ' + eventTitle,
            'Title on CCN Events Promotion Calendar: ' + calendarEventTitle,
            'Percent Match: ' + (matchPercent[0][0]*100) + '%',
            'Row on this spreadsheet: ' + (i+1),
            'Row on CCN Events Promotion Calendar: ' + (j+1),
            'Date on this spreadsheet: ' + eventDate,
            'Date on CCN Events Promotion Calendar: ' + calendarEventDate,
            'Date difference (# days): ' + dayDiff
          ]);
                    
          if (TEST_WRITE_TO_CALENDAR) {
            calendarSheet.getRange(j+1,7).setValue('Yes');
            Log_('Promo initiated flag set for "' + calendarEventTitle + '"');
          }
          
        } // end matchPercent 
        
      } // end dayDiff 
      
    } // next calendar value
    
  } // next response sheet value

  Log_('Errors:');
  Log_(errors);
  Log_('Warnings:');  
  Log_(warnings);
  Log_('recordsFound:');  
  Log_(recordsFound);
    
  if (!isRunFromTrigger) { // build response html only if run manually
  
    // if 0, No Records; if 1, 1 Record; else n Records - just to be gramatically more preciserer
    var html = '<h1>' + (recordsFound.length === 0 ? 'No' : recordsFound.length) + ' Matching Record' + (recordsFound.length === 1 ? '' : 's') + '</h1>';
    
    for (var r in recordsFound) {
    
      for (var rr in recordsFound[i]) {
        html += recordsFound[r][rr] + '<br>';
      }
      
      html += '<br>';
    }
    
    if (warnings.length) {
    
      html += '<h2>Warnings</h2><br>';
      
      for (var w in warnings) {
        html += warnings[w] + '<br>';
      }
    }
    
    if (errors.length) {
    
      html += '<h2>Errors</h2><br>';
      
      for (var err in errors) {
        html += errors[err] + '<br>';
      }
    }

    // show response
    var modal = HtmlService.createHtmlOutput(html).setWidth(800).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(modal, 'Results');
    
  } // !isRunFromTrigger
  
  Log_('Checked promotion calendar'); 
  return
  
  // Private Functions
  // -----------------
  
  /**
   * Get config
   *
   * @param {String} key
   *
   * @return {Object} value
   */
   
  function getConfig(key) {   

    var value = null;
        
    calendarSS.getSheetByName('Config').getDataRange().getValues().some(function(row) {
      if (row[0] === key) {
        value = row[1]
        return true;
      }
    })
    
    if (value === null) {
      throw new Error('Could not find ' + key + ' in the config');
    }
    
    return value;

  } //  checkPromotionCalendar_.getConfig()
  
  /**
   * Get match percentage
   *
   * @param {String} firstString
   * @param {String} secondString
   *
   * @return {Number} match as value from 0 to 1
   */
   
   function getMatchValue(firstString, secondString) {   
      var fuzzy = FuzzySet([firstString]);
      var match = fuzzy.get(secondString);
      return match
    } // checkPromotionCalendar_().getMatchValue()
  
} // checkPromotionCalendar_()