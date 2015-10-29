function isNullOrWS(value) { try  //for logging
{
  return (value === 'undefined' || value == null || value == '' || typeof value === 'string' && value.trim() == '')
} catch(e) { logError(e); throw(e); } } //for logging


function isRangeCrypted(range) { try  //for logging
{
  return range.getFontLine() == 'line-through';
} catch(e) { logError(e); throw(e); } } //for logging


function isLocked()  { try  //for logging
{
  var pek = getProperty('PEK');
  return isNullOrWS(pek);
} catch(e) { logError(e); throw(e); } } //for logging


function tryRemoveAllTriggers() { try  //for logging
{
  // Remove all triggers
  try { // only try because if called during onLoad event, then access to triggers is forbidden
    var triggers = ScriptApp.getProjectTriggers();
    for (i=0; i<triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  } catch (ex) {  }
} catch(e) { logError(e); throw(e); } } //for logging

function embrace(value)  { try  //for logging
{
  return '[' + value + ']';
} catch(e) { logError(e); throw(e); } } //for loggin


//############################
//##        LOGGING
//############################
function logError(ex)
{
  log("Error", ex.fileName + ', line ' + ex.lineNumber + ': ' +  ex.message, ex.stack);
}
function log(type, message, details)
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  var time  = new Date(); 
  if (isNullOrWS(details)) details = '';
  sheet.appendRow([time, type, message, details]);
}

function emptyLogs() {  try  //for logging
{
  var logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs').getRange('A2:D10000').clearContent();;
} catch(e) { logError(e); throw(e); } } //for loggin

//############################
//##    GENERATE PASSWORD
//############################


// Usefull link: http://keepass.info/help/base/pwgenerator.html

function genPass(length, mode) {  try  //for logging
{
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  for (var i = range.getColumnIndex(); i <= range.getLastColumn(); ++i)
  {
    for (var j = range.getRowIndex(); j <= range.getLastRow(); ++j) 
    {
      sheet.getRange(j,i).setNumberFormat('@STRING@').setValue(genNewPassword(length, mode));
    }
  }
} catch(e) { logError(e); throw(e); } } //for loggin

function genNewPassword(length, mode) {  try  //for logging
{
  // mode 1: Numerical
  // mode 2: Alpha
  // mode 3: AlphaNumerical
  // mode 4: With symbols
  // mode 5: With punctuation
  var charset = '';
  if (mode == 1) charset = getSetting('GenPassNum');
  else if (mode == 2) charset = getSetting('GenPassAlpha');
  else charset = getSetting('GenPassNum') + getSetting('GenPassAlpha');
  if (mode > 3) charset += getSetting('GenPassSymb');
  if (mode > 4) charset += getSetting('GenPassPunct');
  var retVal = '';
  for (var i = 0, n = charset.length; i < length; ++i) retVal += charset.charAt(Math.floor(Math.random() * n));
  return retVal;
} catch(e) { logError(e); throw(e); } } //for loggin

function genPassNum256(value) { genPass(256, 1); }
function genPassNum128(value) { genPass(128, 1); }
function genPassNum64(value) { genPass(64, 1); }
function genPassNum20(value) { genPass(20, 1); }
function genPassNum16(value) { genPass(16, 1); }
function genPassNum12(value) { genPass(12, 1); }
function genPassNum10(value) { genPass(10, 1); }
function genPassNum8(value) { genPass(8, 1); }
function genPassNum6(value) { genPass(6, 1); }
function genPassNum4(value) { genPass(4, 1); }

function genPassAlpha256(value) { genPass(256, 2); }
function genPassAlpha128(value) { genPass(128, 2); }
function genPassAlpha64(value) { genPass(64, 2); }
function genPassAlpha20(value) { genPass(20, 2); }
function genPassAlpha16(value) { genPass(16, 2); }
function genPassAlpha12(value) { genPass(12, 2); }
function genPassAlpha10(value) { genPass(10, 2); }
function genPassAlpha8(value) { genPass(8, 2); }
function genPassAlpha6(value) { genPass(6, 2); }
function genPassAlpha4(value) { genPass(4, 2); }

function genPassAlphaNum256(value) { genPass(256, 3); }
function genPassAlphaNum128(value) { genPass(128, 3); }
function genPassAlphaNum64(value) { genPass(64, 3); }
function genPassAlphaNum20(value) { genPass(20, 3); }
function genPassAlphaNum16(value) { genPass(16, 3); }
function genPassAlphaNum12(value) { genPass(12, 3); }
function genPassAlphaNum10(value) { genPass(10, 3); }
function genPassAlphaNum8(value) { genPass(8, 3); }
function genPassAlphaNum6(value) { genPass(6, 3); }
function genPassAlphaNum4(value) { genPass(4, 3); }

function genPassSymb256(value) { genPass(256, 4); }
function genPassSymb128(value) { genPass(128, 4); }
function genPassSymb64(value) { genPass(64, 4); }
function genPassSymb20(value) { genPass(20, 4); }
function genPassSymb16(value) { genPass(16, 4); }
function genPassSymb12(value) { genPass(12, 4); }
function genPassSymb10(value) { genPass(10, 4); }
function genPassSymb8(value) { genPass(8, 4); }
function genPassSymb6(value) { genPass(6, 4); }
function genPassSymb4(value) { genPass(4, 4); }

function genPassPunct256(value) { genPass(256, 5); }
function genPassPunct128(value) { genPass(128, 5); }
function genPassPunct64(value) { genPass(64, 5); }
function genPassPunct20(value) { genPass(20, 5); }
function genPassPunct16(value) { genPass(16, 5); }
function genPassPunct12(value) { genPass(12, 5); }
function genPassPunct10(value) { genPass(10, 5); }
function genPassPunct8(value) { genPass(8, 5); }
function genPassPunct6(value) { genPass(6, 5); }
function genPassPunct4(value) { genPass(4, 5); }

