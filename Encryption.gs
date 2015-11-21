/* ----------------------------------------------------------
Encryption.gs -- 
Copyright (C) 2015 Mifourno

This software may be modified and distributed under the terms
of the MIT license.  See the LICENSE file for details.

GitHub: https://github.com/mifourno/keystore/
Contact: mifourno@gmail.com
---------------------------------------------------------- */

/**
 * @OnlyCurrentDoc
 */


// ##############################################
// ##             ENCRYPTION UTILS
// ##############################################

function generateEncryptionKey() { try  //for logging
{
  var chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghijklmnopqrstuvwxyz";
  var string_length = 20;
  var randomstring = '';
  for (var i=0; i<string_length; i++) {
    var rnum = Math.floor(Math.random() * chars.length);
    randomstring += chars.substring(rnum,rnum+1);
  }
  return randomstring;
} catch(e) { handleError(e); } } //for logging



function isCellEncrypted(range) { try  //for logging
{
  return p_isCellLineThrough(range);
  //if (checkFormatOnly) return p_isCellLineThrough(range);
  //else return p_isCellLineThrough(range) || p_isCellProtected(range);
} catch(e) { handleError(e); } } //for logging


function p_isCellValueCrypted(range, pek) { try  //for logging
{
  // Return values:
  //  0 => cell is empty
  //  1 => cell does not start with encrytion header
  //  2 => value cannot be decrypted
  //  8 => cannot test further than 0 and 1 because the spreadsheet is locked
  //  9 => cell is actually crypted 
  if (isNullOrWS(range.getValue())) return 0;
  if (typeof range.getValue() != 'string' || range.getValue().indexOf("################# - Data:") != 0) return 1;
  if (isNullOrWS(pek)) return 8;
  if (isNullOrWS(decryptValueForCell(range.getValue(), pek))) return 2;
  return 9;
} catch(e) { handleError(e); } } //for logging

function p_isCellLineThrough(range) { try  //for logging
{
  return range.getFontLine() == 'line-through';
} catch(e) { handleError(e); } } //for logging

function p_getCellProtection(range, protectionParams, protectionDictionary) { try  //for logging
{
  if (protectionDictionary != null) return protectionDictionary[getUniqueId(range)];
  
  var sheet = range.getSheet();
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == protectionParams.protectionMessage && protections[i].getRange().getA1Notation() == range.getA1Notation()) return protections[i];
  }
  return null;
} catch(e) { handleError(e); } } //for logging


function p_removeProtection(range, protectionParams, protection, protectionDictionary) { try  //for logging
{
  if (!isNullOrWS(protection)) {
    protection.remove();
    if (protectionDictionary != null) {
      var rangeId = getUniqueId(range);
      protectionDictionary[rangeId] = null;
    }
  } else {
    var protections = range.getSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var j = 0; j < protections.length; j++) {
      if (protections[j].getRange().getA1Notation() == range.getA1Notation() && protections[j].getDescription() == protectionParams.protectionMessage) {
        protections[j].remove();
        break;
      }
    }
  }
} catch(e) { handleError(e); } } //for logging


function getProtectionParams() { try //for logging
{
  var protectionParams = {};
  protectionParams.protectionMessage = getP_ProtectionMessage();
  protectionParams.setFormatAtEncryption = getP_SetFormatAtEncryption();
  protectionParams.initFormat_Background = getP_InitFormat_Background();
  protectionParams.initFormat_Color = getP_InitFormat_Color();
  protectionParams.encryptedFormat_Background = getP_EncryptedFormat_Background();
  protectionParams.encryptedFormat_Color = getP_EncryptedFormat_Color();
  return protectionParams;
} catch(e) { handleError(e); } } //for logging


function encryptValueForCell(value, pek) { try //for logging
{
  return "################# - Data:" + encrypt(value, pek);
} catch(e) { handleError(e); } } //for logging


function decryptValueForCell(encryptedValue, pek) { try //for logging
{
  return decrypt(encryptedValue.substring(25), pek);
} catch(e) { handleError(e); } } //for logging


// ##############################################
// ##              SET STATES
// ##############################################

function p_setStateInit(range, protectionParams, skipFormat, protection, protectionDictionary) { try  //for logging
{
  //Protection
  p_removeProtection(range, protectionParams, protection, protectionDictionary);
  
  //Format
  if (skipFormat != true) {
    range.setNumberFormat('@STRING@');
    range.setFontLine('none');
    if (protectionParams.setFormatAtEncryption) {
      range.setBackground(protectionParams.initFormat_Background);
      range.setFontColor(protectionParams.initFormat_Color);
    }
    if (range.offset(0,1).getValue() == ' ') range.offset(0,1).clear();
  }
} catch(e) { handleError(e); } } //for logging


function p_setStateEncrypted(range, protectionParams, skipFormat, protection, protectionDictionary) { try  //for logging
{
  if (arguments.length < 4) protection = p_getCellProtection(range, protectionParams, protectionDictionary);
  
  //Protection
  if (isNullOrWS(protection)) protection = range.protect().setDescription(protectionParams.protectionMessage).setWarningOnly(true);
  
  //Protection Dictionary
  if (protectionDictionary != null) protectionDictionary[getUniqueId(range)] = protection;
  
  //Format
  if (skipFormat != true) {
    range.setFontLine('line-through');
    if (protectionParams.setFormatAtEncryption) {
      range.setBackground(protectionParams.encryptedFormat_Background);
      range.setFontColor(protectionParams.encryptedFormat_Color);
    }
  }
  if (range.offset(0,1).isBlank()) range.offset(0,1).setValue(' ');
} catch(e) { handleError(e); } } //for logging


// ##############################################
// ##    CONSISTANCY (protection and format)
// ##############################################

function getProtectionDictionary(sheet, protectionParams)  { try  //for logging
{
  var keystoreProtections = new Array();
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == protectionParams.protectionMessage) {
      keystoreProtections[getUniqueId(protections[i].getRange())] = protections[i];
    }
  }
  return keystoreProtections;
} catch(e) { handleError(e); } } //for logging

function checkConsistancyAllSheets()  { try  //for logging
{
  var dirtyCellsBySheet = [];
  var protectionParams = getProtectionParams();
  var pek = getP_PEK(false);
  
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    var dirtyCells = checkConsistancyFullSheet(allSheets[j], pek, protectionParams);
    if (dirtyCells.length > 0) dirtyCellsBySheet.push( { sheetName: allSheets[j].getName(), dirtyCells: dirtyCells } );
  }
  if (dirtyCellsBySheet.length > 0) {
    var dirtyCellsList = '';
    for (var i = 0; i < dirtyCellsBySheet.length; i++) {
      dirtyCellsList += '\nSheet "' + dirtyCellsBySheet[i].sheetName + '": ';
      for (var j = 0; j < dirtyCellsBySheet[i].dirtyCells.length; j++) dirtyCellsList += dirtyCellsBySheet[i].dirtyCells[j].getA1Notation() + ', ';
      dirtyCellsList = dirtyCellsList.substring(0, dirtyCellsList.length - 2);
    }
    serverSideAlert('Some cells cannot be read', "The following cells' data that cannot be decrypted.\nPlease review their content and clear them to return to normal.\n" + dirtyCellsList);
  }
  return dirtyCells;
} catch(e) { handleError(e); } } //for logging


function checkConsistancyFullSheet(sheet, pek, protectionParams) { try  //for logging
{
  if (isNullOrWS(sheet)) sheet = SpreadsheetApp.getActiveSheet();
  if (isNullOrWS(protectionParams)) protectionParams = getProtectionParams();
  var protectionDictionary = getProtectionDictionary(sheet, protectionParams);
  
  var dirtyCells = [];
  var alreadyChecked = '';
  
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == protectionParams.protectionMessage) {
      var range = protections[i].getRange();
      if (!checkConsistancySingleCell(range, pek, protectionParams, protectionDictionary)) dirtyCells.push(range);
      alreadyChecked += range.getA1Notation() + ',';
    }
  }
  
  // This represents ALL the data
  var fullRange = sheet.getDataRange();
  var values = fullRange.getValues();

  // This logs the spreadsheet in CSV format with a trailing comma
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var range = fullRange.getCell(i+1, j+1);
      if (alreadyChecked.indexOf(range.getA1Notation() + ',') < 0) {
        var isLineThrough = p_isCellLineThrough(range);
        if (p_isCellLineThrough(range) || (typeof values[i][j] == 'string' && !isNullOrWS(values[i][j]) && values[i][j].indexOf("################# - Data:") == 0)) {
          if (!checkConsistancySingleCell(range, pek, protectionParams, protectionDictionary)) dirtyCells.push(range);
        }
      }
    }
  }
 
  return dirtyCells;
} catch(e) { handleError(e); } } //for logging


function checkConsistancySingleCell(range, pek, protectionParams, protectionDictionary) { try  //for logging
{
  if (isNullOrWS(range)) range = SpreadsheetApp.getActiveRange();
  if (isNullOrWS(protectionParams)) protectionParams = getProtectionParams();
  if (protectionDictionary == null) protectionDictionary = getProtectionDictionary(range.getSheet(), protectionParams);
  var protection = protectionDictionary[getUniqueId(range)];
  
  var isLineThrough = p_isCellLineThrough(range);
  
  if (isNullOrWS(range.getValue()) && (isLineThrough || protection != null)) {
    p_setStateInit(range, protectionParams, !isLineThrough, protection, protectionDictionary);
    
    return true;
  }
  
  var isCrypted = p_isCellValueCrypted(range, pek);
  
  // Cell is supposed to be encrypted but cannot be decrypted even though PEK is available
  if (isCrypted == 2) {
    p_setStateEncrypted(range, protectionParams, isLineThrough, protection, protectionDictionary); 
    return false;
  }
  
  
  // Cell is empty or does not start with encryption header
  if (isCrypted < 2) {
    if (isLineThrough || protection != null)
    {
      p_setStateInit(range, protectionParams, !isLineThrough, protection, protectionDictionary);
    }
  }
  
  // Cell is supposed to be encrypted. We either don't know if value can be decrypted (spreadsheet is locked) or it's value can be decrypted with success
  if (isCrypted >= 8) {
    if (!isLineThrough || protection == null)
    {
      p_setStateEncrypted(range, protectionParams, isLineThrough, protection, protectionDictionary); 
    }
  }
  return true;
  
  
  
//  if (isLineThrough && protection != null) {
//    // Cell is empty or does not start with encryption header
//    if (isCrypted < 2) p_setStateInit(range, !isLineThrough, protection, protectionDictionary);
//    // Cell is supposed to be encrypted but cannot be decrypted even though PEK is available
//    else if (isCrypted == 2) return false;
//  }
//  else if (isLineThrough || protection != null) {
//    // Cell is empty or does not start with encryption header
//    if (isCrypted < 2) p_setStateInit(range, !isLineThrough, protection, protectionDictionary);
//    // Cell is supposed to be encrypted but cannot be decrypted even though PEK is available
//    else if (isCrypted == 2) {
//      p_setStateEncrypted(range, isLineThrough, protection, protectionDictionary); 
//      return false;
//    }
//    // Cell's value can be decrypted with success
//    else if (isCrypted >= 8) p_setStateEncrypted(range, isLineThrough, protection, protectionDictionary); 
//  } else {
//    // Cell is supposed to be encrypted but cannot be decrypted even though PEK is available
//    if (isCrypted == 2) {
//      p_setStateEncrypted(range, isLineThrough, protection, protectionDictionary); 
//      return false;
//    }
//    if (isCrypted >= 8) p_setStateEncrypted(range, isLineThrough, protection, protectionDictionary); 
//  }
  
  
  
} catch(e) { handleError(e); } } //for logging


// ##############################################
// ##                ENCRYPTING
// ##############################################


function markAsSensitive() { try  //for logging
{
  if (isLocked()) promptMasterPassword('assert', 'markAsSensitiveAdmin');
  else markAsSensitiveAdmin();
} catch(e) { handleError(e); } } //for logging


function markAsSensitiveAdmin() { try  //for logging
{
  var pek = getP_PEK(true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  
  var protectionParams = getProtectionParams();
  var protectionDictionary = getProtectionDictionary(sheet, protectionParams);
    
  // Check whether some cells are already encrypted
  for (var i = range.getColumnIndex(); i <= range.getLastColumn(); ++i)
  {
    for (var j = range.getRowIndex(); j <= range.getLastRow(); ++j) 
    {
      checkConsistancySingleCell(sheet.getRange(j,i), pek, protectionParams, protectionDictionary);
      if (isCellEncrypted(sheet.getRange(j,i)))
      {
        serverSideAlert('Already encrypted !', 'Some cells in selection are already encrypted ! Encryption Aborted !');
        return;
      }
    }
  }
  
  var rangeCounter = 0;
  for (var col = range.getColumnIndex(); col <= range.getLastColumn(); ++col)
  {
    for (var row = range.getRowIndex(); row <= range.getLastRow(); ++row) 
    {
      if (markAsSensitive_SingleCell(sheet, protectionParams, row, col, pek)) rangeCounter++;
    }
  }
  
  if (rangeCounter == 0) serverSideAlert('Empty selection', 'No data to encrypt within this range selection');
  
} catch(e) { handleError(e); } } //for logging


function markAsSensitive_SingleCell(sheet, protectionParams, row, col, pek) { try //for logging
{
  var rangeToEncrypt = sheet.getRange(row,col);
  if (isNullOrWS(rangeToEncrypt.getValue())) return false;  
  p_setStateEncrypted(rangeToEncrypt, protectionParams);
  rangeToEncrypt.setValue(encryptValueForCell(rangeToEncrypt.getValue(), pek));
  return true;
} catch(e) { handleError(e); } } //for logging


// ##############################################
// ##                REVEALING
// ##############################################

function removeEncryption() { try  //for logging
{
  reveal('permanent');
} catch(e) { handleError(e); } } //for logging

function revealPopup() { try  //for logging
{
  reveal('popup');
} catch(e) { handleError(e); } } //for logging

// mode: permanent|popup
function reveal(mode) { try  //for logging
{
  if (isLocked()) promptMasterPassword('assert', 'revealAdmin', null, mode);
  else revealAdmin(mode);
} catch(e) { handleError(e); } } //for logging


// mode: permanent|popup
function revealAdmin(mode) { try  //for logging
{
  var pek = getP_PEK(true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  
  var encryptedRangeCounter = 0;
  var corruptedRangeCounter = 0;
  var revealedRangeCounter = 0;
  var corruptedRanges = '';
  var revealedValuesList = [];
  
  var protectionParams = getProtectionParams();
  var protectionDictionary = getProtectionDictionary(sheet, protectionParams);
  
  
  for (var row = range.getRowIndex(); row <= range.getLastRow(); ++row) 
  {
    for (var col = range.getColumnIndex(); col <= range.getLastColumn(); ++col)
    {
      checkConsistancySingleCell(sheet.getRange(row,col), pek, protectionParams, protectionDictionary);
      var currentRange = sheet.getRange(row,col);
      if (!isCellEncrypted(currentRange)) {
        revealedValuesList.push(currentRange.getValue());
      } else {
        encryptedRangeCounter++;
        var revealedValue = decryptValueForCell(currentRange.getValue(),pek);
        if (isNullOrWS(revealedValue)) {
          revealedValuesList.push('N/A');
          corruptedRanges += currentRange.getA1Notation() + ", ";
          corruptedRangeCounter++;
        } else {
          revealedValuesList.push(revealedValue);
          revealedRangeCounter++;
          if (mode != 'popup') {
            currentRange.setValue(revealedValue);
            p_setStateInit(currentRange, protectionParams);
          }
        }
      }
    }
  }
  if (encryptedRangeCounter == 0) {
    customDialog('Nothing to decrypt', 'The cells you selected are not encrypted.', 100);
  } 
  if (revealedRangeCounter > 0) {
    if (mode == 'popup') {
      showRevealPopup(range, revealedValuesList);
    }
  }
  if ((mode != 'popup' || revealedRangeCounter == 0) && corruptedRangeCounter > 0) customDialog('Warning !', 'The following cells could not be decrypted:\n' + corruptedRanges.substring(0, corruptedRanges.length - 2), 130);
  
} catch(e) { handleError(e); } } //for logging


function showRevealPopup(range, valuesList) { try  //for logging
{
  var sheet = SpreadsheetApp.getActiveSheet();
  height = Math.min(75 + 23 * range.getNumRows(), 400);
  var html = HtmlService.createTemplateFromFile('RevealPopup');
  var columns = [];
  for (var i = 0; i<range.getNumColumns(); i++) {
    var a1Notation = sheet.getRange(1, range.getColumn() + i).getA1Notation();
    columns.push(a1Notation.substring(0, a1Notation.length - 1));
  }
  html.data = { 'values' : valuesList, 'columns' : columns, 'firstRow' : range.getRow(), 'height' : height };
  dialog = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(dialog, 'Revealed values');
} catch(e) { handleError(e); } } //for logging

