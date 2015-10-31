/* ----------------------------------------------------------
Locking.gs -- 
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


function assertUnlocked(message, requirePassword)  { try  //for logging
{
  setP_LastUpdate(new Date());
  var pek = getP_PEK();
  if (!requirePassword && !isNullOrWS(pek)) return true;
  
  if (isNullOrWS(message)) message = '';
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.prompt('Master Password', message + 'Please enter your master password:', ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  var button = result.getSelectedButton();
  var input = result.getResponseText();
  if (button == ui.Button.OK) {
    if (!unlockSpreasheet(input)) {
      return assertUnlocked('Invalid password !\n');
    } else {
      return true;
    }
  }
  return false;
} catch(e) { logError(e); throw(e); } } //for logging

function unlockSpreasheet(masterPassword) { try  //for logging
{
  var pek = decrypt(getP_EEK(), masterPassword);
  if (isNullOrWS(pek)) return false;
  
  log('Diagnostic', 'Unlock spreadshit');
  setP_LastUpdate(new Date());
  setP_PEK(pek);
  startAutoLockTrigger();
  return true;
  
} catch(e) { logError(e); throw(e); } } //for logging


function manualLockSpreasheet(source) { try  //for logging
{
  lockSpreasheet('manual');
} catch(e) { logError(e); throw(e); } } //for logging


function lockSpreasheet(source) { try  //for logging
{
  log('Diagnostic', 'Lock spreadshit (' + source + ')');
  reencryptRevealedRange();
  setP_PEK('');
  setP_LockedAt(new Date());
  tryRemoveAllTriggers();
} catch(e) { logError(e); throw(e); } } //for logging


//##################################
//##     RESET MASTER PASSWORD
//##################################

function changeMasterPassword() { try  //for logging
{
  log('Diagnostic', 'Change master-password');
  if (!assertUnlocked(null, true)) return false;
  
  var pek = getP_PEK();
  var newPassword = askForNewMasterPassword();
  if (newPassword != null) {
    setP_EEK(encrypt(pek, newPassword));
    var ui = SpreadsheetApp.getUi();
    ui.alert('Success !', 'Your master-password has been changed successfully', ui.ButtonSet.OK);
  }
} catch(e) { logError(e); throw(e); } } //for logging


function resetSpreasheet() { try  //for logging
{
  var programName = getP_ProgramName();
  
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Reset ' + programName + ' ?',
     'Reseting ' + programName + ' will result in the loss of all crypted data in the current spreadsheet. If you continue, you will be prompt for a new Master Password. Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    var newPassword = askForNewMasterPassword();
    if (newPassword != null) {
      initializeProperties(false);
      var newEncryptionKey = generateEncryptionKey();
      setP_EEK(encrypt(newEncryptionKey, newPassword));
      unlockSpreasheet(newPassword);
    };
  }   
} catch(e) { logError(e); throw(e); } } //for logging


function askForNewMasterPassword() { try  //for logging
{
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt('New master-password', 'Please enter your new master-password:\nWARNING: if you loose this password, there will be no way to recover encrypted data', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var input = result.getResponseText();
  if (button == ui.Button.OK) {
    return confirmMasterPassword(input, '');
  } else if (button == ui.Button.CANCEL) {
    return null;
  } else if (button == ui.Button.CLOSE) {
    return null;
  }
  return null;
} catch(e) { logError(e); throw(e); } } //for logging


function confirmMasterPassword(password, message) { try  //for logging
{
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.prompt('Confirm your master-password', message + 'Please confirm your master-password:', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var input = result.getResponseText();
  if (button == ui.Button.OK) {
    if (input == password) return input;
    else confirmMasterPassword(password, 'Your input is different from the password you entered. \n');
  } else if (button == ui.Button.CANCEL) {
    return null;
  } else if (button == ui.Button.CLOSE) {
    return null;
  }
  return null;
} catch(e) { logError(e); throw(e); } } //for logging


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
} catch(e) { logError(e); throw(e); } } //for logging


//##################################
//##        AUTOLOCK
//##################################

function checkAutolock()  { try  //for logging
{
  if (isLocked()) { tryRemoveAllTriggers(); return; }
  var diffMs = (new Date()) - new Date(getP_LastUpdate());
  var diffSecs = Math.round(((diffMs % 86400000) % 3600000) / 1000); // minutes
  if (diffSecs > getP_AutolockDelay()*60) {
    lockSpreasheet('auto');
    SpreadsheetApp.getActiveSpreadsheet().toast('Keystore is locked', 'Autolock !');
  }
} catch(e) { logError(e); throw(e); } } //for logging

function startAutoLockTrigger()  { try  //for logging
{
  var autolockDelay = getP_AutolockDelay();
  if (autolockDelay > 0) {
    var delay = Math.ceil(Math.max(autolockDelay / 10, 1));
    ScriptApp.newTrigger("checkAutolock").timeBased().everyMinutes(delay).create();
  }
} catch(e) { logError(e); throw(e); } } //for logging
