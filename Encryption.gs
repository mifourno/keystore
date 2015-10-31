/* ----------------------------------------------------------
Encryption.gs -- 
Copyright (C) 2015 Mifourno

This software may be modified and distributed under the terms
of the MIT license.  See the LICENSE file for details.

GitHub: https://github.com/mifourno/keystore/
Contact: mifourno@gmail.com

DEPENDENCIES:
- Encryption.gs
- Menu.gs
- Properties.gs
- Utils.gs
- Locking.gs
- CryptoJSWrapper.gs
- CryptoJS Files:
    => CryptoJS_aes.gs
---------------------------------------------------------- */

/**
 * @OnlyCurrentDoc
 */

function markAsSensitive() { try  //for logging
{
  if (!assertUnlocked()) return false;
  var pek = getP_PEK();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  var protectionMessage = getP_ProtectionMessage();
  var setFormatAtEncryption = getP_SetFormatAtEncryption();
  var background = getP_EncryptedFormat_Background();
  var color = getP_EncryptedFormat_Color();
  for (var i = range.getColumnIndex(); i <= range.getLastColumn(); ++i)
  {
    for (var j = range.getRowIndex(); j <= range.getLastRow(); ++j) 
    {
      if (isRangeCrypted(sheet.getRange(j,i)))
      {
        var ui = SpreadsheetApp.getUi(); // Same variations.
        ui.alert('Already encrypted !', 'Some cells in selection are already encrypted ! Encryption Aborted !', ui.ButtonSet.OK);
        return;
      }
    }
  }
  var rangeCounter = 0;
  for (var col = range.getColumnIndex(); col <= range.getLastColumn(); ++col)
  {
    for (var row = range.getRowIndex(); row <= range.getLastRow(); ++row) 
    {
      if (markAsSensitive_SingleCell(sheet, row,col, pek, protectionMessage, setFormatAtEncryption, background, color)) rangeCounter++;
    }
  }
  
  if (rangeCounter == 0) SpreadsheetApp.getActiveSpreadsheet().toast('No data to encrypt within this range selection');
  
} catch(e) { handleError(e); } } //for logging


function markAsSensitive_SingleCell(sheet, row, col, pek, protectionMessage, setFormatAtEncryption, background, color) { try //for logging
{
  var rangeToEncrypt = sheet.getRange(row,col);
  if (isNullOrWS(rangeToEncrypt.getValue())) return false;  
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (rangeToEncrypt.offset(0,1).isBlank()) rangeToEncrypt.offset(0,1).setValue(' ');
  rangeToEncrypt.setFontLine('line-through');
  rangeToEncrypt.setValue(encrypt(rangeToEncrypt.getValue(),pek));
  var rangeFound = false;
  for (var j = 0; j < protections.length; j++) {
    if (protections[j].getDescription() == protectionMessage && protections[j].getRange().getA1Notation() == rangeToEncrypt.getA1Notation()) { rangeFound = true; break; }
  }
  if (!rangeFound) rangeToEncrypt.protect().setDescription(protectionMessage).setWarningOnly(true);
  if (setFormatAtEncryption) {
    rangeToEncrypt.setBackground(background);
    rangeToEncrypt.setFontColor(color);
  }
  
  return true;
} catch(e) { handleError(e); } } //for logging


function revealFewSeconds() { try  //for logging
{
  reveal('fewSeconds');
} catch(e) { handleError(e); } } //for logging

function revealUntilLock() { try  //for logging
{
  reveal('untilLock');
} catch(e) { handleError(e); } } //for logging

function removeEncryption() { try  //for logging
{
  reveal('permanent');
} catch(e) { handleError(e); } } //for logging

function revealPopup() { try  //for logging
{
  reveal('popup');
} catch(e) { handleError(e); } } //for logging


// mode: permanent|fewSeconds|untilLock|popup
function reveal(mode) { try  //for logging
{
  if (!assertUnlocked()) return false;
  var pek = getP_PEK();
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  
  var rangeCounter = 0;
  if (mode == 'popup') {
    if (isRangeCrypted(range)) {
      rangeCounter++;
      var ui = SpreadsheetApp.getUi();
      ui.alert('Select the value below and hit Ctrl+C', decrypt(range.getValue(),pek), ui.ButtonSet.OK);
    }
  } else {
    if (mode != 'permanent') {
      var revealedRangeSheets = getP_RevealedRangeSheets();
      if (revealedRangeSheets.indexOf(embrace(sheet.getSheetId())) < 0) setP_RevealedRangeSheets(revealedRangeSheets + embrace(sheet.getSheetId()));
    }
    for (var col = range.getColumnIndex(); col <= range.getLastColumn(); ++col)
    {
      for (var row = range.getRowIndex(); row <= range.getLastRow(); ++row) 
      {
        var currentRange = sheet.getRange(row,col);
        if (isRangeCrypted(currentRange)) {
          rangeCounter++;
          currentRange.setNumberFormat('@STRING@');
          currentRange.setValue(decrypt(currentRange.getValue(),pek));
          currentRange.setFontLine('none');
          if (mode == 'permanent') {
            removeProtection(currentRange);
            if (getP_SetFormatAtEncryption()) {
              currentRange.setBackground(getP_DecryptedFormat_Background());
              currentRange.setFontColor(getP_DecryptedFormat_Color());
            }
          } else {
            if (getP_SetFormatAtEncryption()) {
              currentRange.setBackground(getP_RevealedFormat_Background());
              currentRange.setFontColor(getP_RevealedFormat_Color());
            }
          }
          if (currentRange.offset(0,1).getValue() == ' ') currentRange.offset(0,1).clear();
        }
      }
    }
    if (mode == 'fewSeconds') ScriptApp.newTrigger("autoReencryptRevealedRange").timeBased().after(60000).create();
  }
  if (rangeCounter == 0) SpreadsheetApp.getUi().alert('No encrypted cell found within this range selection');
} catch(e) { handleError(e); } } //for logging


/*function justAlert() { try  //for logging
{
  //var spreadsheet = SpreadsheetApp.openById('14WbRsYVJsZbrf9_gNS3eOUjPLI7gdFxiPwVXbrkO7bY');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Data');
  sheet.getRange("A1").setValue(new Date());
} catch(e) { handleError(e); } } //for logging
*/


function autoReencryptRevealedRange() { try  //for logging
{
  log('Diagnostic', 'Auto reencrypt revealed range');
  reencryptRevealedRange();
} catch(e) { handleError(e); } } //for logging


function reencryptRevealedRange() { try  //for logging
{
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var revealedRangeSheets = getP_RevealedRangeSheets();
  if (!isNullOrWS(revealedRangeSheets)) {
    for (var j = 0; j < allSheets.length; j++) {
      if (revealedRangeSheets.indexOf(embrace(allSheets[j].getSheetId())) >= 0)
      {
        syncLineThroughAndProtectionOnSheet(allSheets[j]);
        revealedRangeSheets = revealedRangeSheets.replace(embrace(allSheets[j].getSheetId()), '');
        setP_RevealedRangeSheets(revealedRangeSheets);
      }
    }
  }
} catch(e) { handleError(e); } } //for logging


function removeProtection(range) { try  //for logging
{
  var protections = range.getSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var j = 0; j < protections.length; j++) {
    if (protections[j].getRange().getA1Notation() == range.getA1Notation()) protections[j].remove();
  }
} catch(e) { handleError(e); } } //for logging


function syncLineThroughAndProtection()  { try  //for logging
{
  var rangeList = '';
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    syncLineThroughAndProtectionOnSheet(allSheets[j]);
  }
} catch(e) { handleError(e); } } //for logging


function syncLineThroughAndProtectionOnSheet(sheet) { try  //for logging
{
  var pek = getP_PEK();
  var protectionMessage = getP_ProtectionMessage();
  var setFormatAtEncryption = getP_SetFormatAtEncryption();
  var background = getP_EncryptedFormat_Background();
  var color = getP_EncryptedFormat_Color();
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var dirtyRange = '';
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() == protectionMessage) {
      if (protections[i].getRange().isBlank()) protections[i].remove();
      else if (!isRangeCrypted(protections[i].getRange())) { 
        if (!isNullOrWS(pek)) markAsSensitive_SingleCell(sheet, protections[i].getRange().getRow(), protections[i].getRange().getColumn(), pek, protectionMessage, setFormatAtEncryption, background, color);
        else dirtyRange += protections[i].getRange().getA1Notation() + ', ';
      }
    }
  }
  if (!isNullOrWS(dirtyRange)) 
  {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Oops: dirty range !', 'The following cells should be encrypted, but font is not striked out. Please review and either encrypt them or simply set the font strike out to fix the problem\n\n' + 'Sheet "' + sheet.getName() + '": ' + dirtyRange, ui.ButtonSet.OK);
  }
} catch(e) { handleError(e); } } //for logging
