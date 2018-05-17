var SCRIPT_NAME = "Promotion Request Form"
var SCRIPT_VERSION = "v1.3"

var CONFIG = {
  TEST_WRITE_TO_CALENDAR : false,
  DATA_SHEET_NAME : 'Incoming_Data',
  FILES : {
    RESPONSES_SPREADSHEET_ID           : '1JEqPQJSiBliliqw1y-wrrdP6ikU11DPuIF72l-rN84g', // Live
    PROMOTION_CALENDAR_SPREADSHEET_URL : 'https://docs.google.com/spreadsheets/d/1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc/edit', // Live
    PROMOTION_CALENDAR_SHEET_NAME      : 'Communications Director Master',
  },
  MATCH_THRESHOLD_PERCENT : 0.75, // fuzzy logic matching
  MAX_EVENT_DATE_DIFF : 10,  
};