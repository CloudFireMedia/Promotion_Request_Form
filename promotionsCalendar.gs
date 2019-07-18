/**
 * On a form submission check every response against every entry in the 
 * the promotions calendar to see if there are any matches. If so record 
 * in the calendar GSheet that a promotion request has been made for this event.
 */

function checkPromotionCalendar_() {

  if (!TEST_CHECK_PROMOTION_CALENDAR) {
    return;
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
  
  if (calendarSheet === null) {
  
    throw new Error (
      'Unable to find sheet named "' + PROMOTION_CALENDAR_SHEET_NAME + '" ' + 
      'on ' + calendarSS.getName() + ' spreadsheet');      
  }
  
  var calendarValues = calendarSheet.getDataRange().getValues();
  var columnNumbers = Utils.getPRFColumns(responsesSpreadsheet); 

  var errors = [];
  var warnings = [];
  var recordsFound = [];
  
  var startDateColumnIndex = columnNumbers.EventStart - 1;
  var titleColumnIndex = columnNumbers.EventTitle - 1;
  
  var matchThreshold = getConfig('MATCH_THRESHOLD');
  var maxEventDayDiff = getConfig('MAX_EVENT_DATE_DIFF');
  
  var recordsFound = {};

  for (var prfRowIndex = 1; prfRowIndex < responseSheetValues.length; prfRowIndex++){ // prfRowIndex==1 to skip header row
  
    var pRFRowNumber = prfRowIndex + 1;
  
    // check for "errors" first
    var eventDate = responseSheetValues[prfRowIndex][startDateColumnIndex];
    
    if (!eventDate) {
      errors.push('No event date found, line ' + pRFRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
      continue;
    }
    
    if (!eventDate instanceof Date){
      errors.push('Date ' + eventDate + ' on row ' + pRFRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet is not a valid date.  Skipping this row.');
      continue;
    }
    
    var eventTitle = responseSheetValues[prfRowIndex][titleColumnIndex].trim();
    
    if (!eventTitle) {
      warnings.push('No event title found, line ' + pRFRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
      continue;
    }
    
    // find matching event on calendar sheet
    
    // pdcRowIndex=3 skips header rows 0-2 - they aren't frozen and we can't count on them staying set since humans are involved so we just skip three
    for (var pdcRowIndex = 3; pdcRowIndex < calendarValues.length; pdcRowIndex++){ 
    
      // check for "errors" first
      var calendarEventDate = calendarValues[pdcRowIndex][3]; //column 4 is SHORT START DATE
      var pDCRowNumber = pdcRowIndex + 1
      
      if (!calendarEventDate) {
        errors.push('No event date found, line ' + pDCRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
        continue;
      }
      
      if (!(calendarEventDate instanceof Date)) {
        errors.push('Date ' + calendarEventDate + ' on row ' + pDCRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet is not a valid date.  Skipping this row.');
        continue;
      }
      
      var calendarEventTitle = calendarValues[pdcRowIndex][4].trim(); //column 5 is EVENT TITLE
      
      if (calendarEventTitle === ''){
        warnings.push('No event title found, line ' + pDCRowNumber + ' of ' + DATA_SHEET_NAME + ' sheet.  Skipping this row.');
        continue;
      }
      
      var dayDiff = Math.abs(Utils.DateDiff.inDays(calendarEventDate, eventDate));
      
      if (dayDiff > maxEventDayDiff) {
        continue;
      }
      
      var match = getMatchValue(calendarEventTitle, eventTitle);

      if (match === null || match < matchThreshold) {
        continue;
      }

      // If there is a match store all of them, for each each PRF response so we can later decide 
      // which one to use: recordsFound is an object where the PRF row number is the key and the 
      // value is an array of objects detailing each match to this PRF event. This will be 
      // processed later to determine which is the best match
           
      if (!recordsFound.hasOwnProperty(pRFRowNumber)) {
        recordsFound[pRFRowNumber] = []
      }     
      
      // Convert the match and data diff into a score between 0 and 10
      var matchScore = ((match - matchThreshold)/(1 - matchThreshold))*10
      var dateDiffScore = maxEventDayDiff - dayDiff
      
      recordsFound[pRFRowNumber].push({
        pRFTitle      : eventTitle,
        pDCTitle      : calendarEventTitle,
        match         : match,
        pRFRowNumber  : pRFRowNumber,
        pDCRowNumber  : pDCRowNumber,
        pRFDate       : eventDate,
        pDCDate       : calendarEventDate,
        dateDiff      : dayDiff,
        score         : matchScore + dateDiffScore,
      }); 
         
    } // next calendar value
    
  } // next response sheet value

  Log_('Errors:');
  Log_(errors);
  Log_('Warnings:');  
  Log_(warnings);
  Log_('recordsFound:');  
  Log_(recordsFound);

  setPromoRequested(calendarSheet)

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
      var fuzzy = Utils.FuzzySet([firstString]);
      var match = fuzzy.get(secondString);
      if (match !== null) {
        return match[0][0]
      } else {
        return null
      }
    } // checkPromotionCalendar_().getMatchValue()

    /**
     * Set the "Promo Requested" flag in the PDC
     */
     
    function setPromoRequested(calendarSheet) {
    
      if (!TEST_WRITE_TO_CALENDAR) {      
        return;
      }

      for (var pRFRowNumber in recordsFound) {
      
        if (!recordsFound.hasOwnProperty(pRFRowNumber)) {
          continue;
        }
        
        var pDCRowNumber
        var nextRecord = recordsFound[pRFRowNumber]

        // Chcek if there is only one match to this PRF event
        if (nextRecord.length === 1) {
          pDCRowNumber = nextRecord[0].pDCRowNumber
        } else {
          pDCRowNumber = getRowWithBestScore()          
        }

        if (pDCRowNumber === null || pDCRowNumber === undefined) {
          throw new Error('Found more than one matching entry for PRF row number ' + pRFRowNumber)
        }

        calendarSheet.getRange(pDCRowNumber, 7).setValue('Yes'); // column 7 is PROMO REQ?        
        Log_('Promo initiated flag set for in PDC rown number ' + pDCRowNumber);
        
      } // for each record
      
      return; 
      
      // Private Functions
      // -----------------
         
      /** 
       * Select the PDC event with the best match/date diff score
       */
       
      function getRowWithBestScore() {
      
        var pDCRowNumber
        var bestScoreCount = {}
        var bestScore = null
        
        nextRecord.forEach(function(record) {
        
          if (bestScore === null || record.score >= bestScore) {
          
            bestScore = record.score
            pDCRowNumber = record.pDCRowNumber
            Log_('New best score: ' + bestScore + ', PDC row number ' + pDCRowNumber)

            if (!bestScoreCount.hasOwnProperty(bestScore)) {
              bestScoreCount[bestScore] = 0
            }
            
            bestScoreCount[bestScore]++
          }
        })
        
        if (bestScoreCount[bestScore] === 1) {
          return pDCRowNumber
        } else {
          return null // More than one record with this high score
        }
              
      } // checkPromotionCalendar_.setPromoRequested.getRowWithBestScore()
         
    } // checkPromotionCalendar_.setPromoRequested()

} // checkPromotionCalendar_()