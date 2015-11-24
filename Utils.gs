/* ----------------------------------------------------------
Utils.gs -- 
Copyright (C) 2015 Mifourno

This software may be modified and distributed under the terms
of the MIT license.  See the LICENSE file for details.

GitHub: https://github.com/mifourno/keystore/
Contact: mifourno@gmail.com

DEPENDENCIES:
---------------------------------------------------------- */

/**
 * @OnlyCurrentDoc
 */


function isNullOrWS(value) { try  //for logging
{
  return (typeof value == 'undefined' || value == null || value == '' || typeof value === 'string' && value.trim() == '');
} catch(e) { handleError(e); } } //for logging

function getUniqueId(range) { try  //for logging
{
  return range.getSheet().getSheetId() + ',' + range.getA1Notation();
} catch(e) { handleError(e); } } //for logging

function serverSideAlert(title, msg) {  try  //for logging
{
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
} catch(e) { handleError(e); } } //for logging


function tryRemoveAllTriggers() { try  //for logging
{
  // Remove all triggers
  try { // only try because if called during onLoad event, then access to triggers is forbidden
    var triggers = ScriptApp.getProjectTriggers();
    for (i=0; i<triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  } catch (ex) {  }
} catch(e) { handleError(e); } } //for logging


function removeAllProtections()  { try  //for logging
{
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var protectionParams = getProtectionParams();
  
  for (var j = 0; j < allSheets.length; j++) {
    var protections = allSheets[j].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.getDescription() == protectionParams.protectionMessage) {
        var isLineThrough = p_isCellLineThrough(protection.getRange());
        p_setStateInit(protection.getRange(), protectionParams, !isLineThrough, protection);
      }
    }
  }
} catch(e) { handleError(e); } } //for logging


function embrace(value)  { try  //for logging
{
  return '[' + value + ']';
} catch(e) { handleError(e); } } //for logging

function splitBraces(value)  { try  //for logging
{
  return value.substring(1, value.length-2).split('][');
} catch(e) { handleError(e); } } //for logging


function isOwner()  { try  //for logging
{
  return (SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail() == getP_CurrentUser());
} catch(e) { handleError(e); } } //for logging

function isEditor()  { try  //for logging
{
  var userEmail = getP_CurrentUser();
  if (isNullOrWS(userEmail)) return null;
  var editors = SpreadsheetApp.getActiveSpreadsheet().getEditors();
  for (var i = 0; i < editors.length; i++) {
    if (userEmail == editors[i].getEmail()) return true;
  }
  return false;
} catch(e) { handleError(e); } } //for logging


function customDialog(title, msg, height)  { try  //for logging
{
  if (isNullOrWS(height)) height = 200;
  var html = HtmlService.createTemplateFromFile('CustomDialog');
  html.data = { 'message' : msg };
  dialog = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(dialog, title);
} catch(e) { handleError(e); } } //for logging


//############################
//##        LOGGING
//############################
function handleError(ex)
{
  if (isNullOrWS(ex)) {
    log("Error", 'An error occured');
    SpreadsheetApp.getActiveSpreadsheet().toast('An error occured');
  } else if (isNullOrWS(ex.message)) {
    log("Error", ex);
    SpreadsheetApp.getActiveSpreadsheet().toast('An error occured: ' + ex);
  } else {
    log("Error", ex.fileName + ', line ' + ex.lineNumber + ': ' +  ex.message, ex.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast('An error occured: ' + ex.message);
  }
}
function log(type, message, details)
{
  if (isNullOrWS(details)) Logger.log('Type: %s, Message: %s', type, message);
  else Logger.log('Type: %s, Message: %s\nDetails: %s', type, message, details);
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  if (sheet != null) {
    var time  = new Date(); 
    if (isNullOrWS(details)) details = '';
    sheet.appendRow([time, type, message, details]);
  }
}

function emptyLogs() {  try  //for logging
{
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Clear logs', 'Are you sure you want to empty this log table ?', ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    var logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    if (logsSheet != null) logsSheet.getRange('A2:D10000').clearContent();
  }
} catch(e) { handleError(e); } } //for logging





//############################
//##    GENERATE PASSWORD
//############################


// Usefull link: http://keepass.info/help/base/pwgenerator.html

function genPass(length, mode, autoEncryptPassword) { try  //for logging
{
  if (isNullOrWS(autoEncryptPassword)) autoEncryptPassword = getP_AutoEncryptNewPassword();
  var params = length + ',' + mode + ',' + autoEncryptPassword;
  if (autoEncryptPassword && isLocked()) promptMasterPassword('assert', 'genPass_Admin', null, params);
  else genPass_Admin(params);
} catch(e) { handleError(e); } } //for logging


function genPass_Admin(params) {  try  //for logging
{
  var splitParams = params.split(',');
  var length = parseInt(splitParams[0]);
  var mode = splitParams[1];
  var autoEncryptPassword = (splitParams[2] === "true");
  var newPasswordList = [];
  
  var pek = '';
  if (autoEncryptPassword) getP_PEK(true);
  else getP_PEK(false);
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  
  var protectionParams = getProtectionParams();
  var protectionDictionary = getProtectionDictionary(sheet, protectionParams);
  
  var foundEncryptedCells = false;
  for (var i = range.getColumnIndex(); i <= range.getLastColumn(); ++i)
  {
    for (var j = range.getRowIndex(); j <= range.getLastRow(); ++j) 
    {
      checkConsistancySingleCell(sheet.getRange(j,i), pek, protectionParams, protectionDictionary);
      if (isCellEncrypted(sheet.getRange(j,i))) foundEncryptedCells = true;
    }
  }
  
  if (foundEncryptedCells) {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Warning !', 'Some cells in selection are encrypted !\nGenerating password will overwrite your encrypted data.\nAre you sure you want to continue ?', ui.ButtonSet.YES_NO);
    if (result == ui.Button.NO) return;
  }
  
  
  
  // mode 1: Numerical
  // mode 2: Alpha
  // mode 3: AlphaNumerical
  // mode 4: With symbols
  // mode 5: With punctuation
  var charset = '';
  if (mode == 1) charset = getP_GenPassNum();
  else if (mode == 2) charset = getP_GenPassAlpha();
  else charset = getP_GenPassNum() + getP_GenPassAlpha();
  if (mode > 3) charset += getP_GenPassSymbols();
  if (mode > 4) charset += getP_GenPassPunctuations();
  
  for (var j = range.getRowIndex(); j <= range.getLastRow(); ++j) 
  {
    for (var i = range.getColumnIndex(); i <= range.getLastColumn(); ++i)
    {
      var newPassword = genNewPassword(length, charset);
      var currentRange = sheet.getRange(j,i);
      if (autoEncryptPassword) {
        newPasswordList.push(newPassword);
        currentRange.setValue(encryptValueForCell(newPassword, pek));
        p_setStateEncrypted(currentRange, protectionParams);
      } else {
        currentRange.setValue(newPassword);
        p_setStateInit(currentRange, protectionParams);
      }
    }
  }
  if (autoEncryptPassword) showRevealPopup(range, newPasswordList);
  
} catch(e) { handleError(e); } } //for logging

function genNewPassword(length, charset) {  try  //for logging
{
  var retVal = '';
  for (var i = 0, n = charset.length; i < length; ++i) retVal += charset.charAt(Math.floor(Math.random() * n));
  return retVal;
} catch(e) { handleError(e); } } //for logging

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

