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

function isLocked()  { try  //for logging
{
  var pek = getP_PEK();
  return isNullOrWS(pek);
} catch(e) { handleError(e); } } //for logging


function promptMasterPassword(mode, okHandlerName, cancelHandlerName, argsOkString, argsCancelString) { try  //for logging
{
  // Modes: assert, init, change
  if (isNullOrWS(mode)) mode = 'assert';
  if (isNullOrWS(okHandlerName)) okHandlerName = null;
  if (isNullOrWS(cancelHandlerName)) cancelHandlerName = null;
  if (isNullOrWS(argsOkString)) argsOkString = null;
  if (isNullOrWS(argsCancelString)) argsCancelString = null;
 
  if (mode == 'assert' && !isLocked()) {
    if (okHandlerName != null) eval(okHandlerName + "('" + argsOkString + "')");
    return;
  }
  
  var title = 'Master Password';
  if (mode == 'change') title = 'Change master password';
  var height = 400;
  switch (mode) {
    case 'change': height = 300; break;
    case 'init': height = 250; break;
    case 'assert': height = 140; break;
  }
  
  var html = HtmlService.createTemplateFromFile('MasterPassword');
  html.data = { 'mode': mode, 'okHandlerName' : okHandlerName, 'cancelHandlerName' : cancelHandlerName, 'argsOkString' : argsOkString, 'argsCancelString' : argsCancelString };
  dialog = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(dialog, title);
} catch(e) { handleError(e); } } //for logging
  

function unlockSpreasheet(masterPassword) { try  //for logging
{
  var pek = decrypt(getP_EEK(), masterPassword);
  if (isNullOrWS(pek)) return false;
  
  log('Diagnostic', 'Unlock spreadshit');
  setP_LastUpdate(new Date());
  setP_PEK(pek);
  startAutoLockTrigger();
  return true;
  
} catch(e) { handleError(e); } } //for logging


function manualLockSpreasheet(source) { try  //for logging
{
  lockSpreasheet('manual');
} catch(e) { handleError(e); } } //for logging


function lockSpreasheet(source) { try  //for logging
{
  log('Diagnostic', 'Lock spreadshit (' + source + ')');
  reencryptRevealedRange();
  setP_PEK('');
  setP_LockedAt(new Date());
  tryRemoveAllTriggers();
} catch(e) { handleError(e); } } //for logging


//##################################
//##     RESET MASTER PASSWORD
//##################################


function changeMasterPassword() { try  //for logging
{
  promptMasterPassword('change', 'changeMasterPasswordAdmin');
} catch(e) { handleError(e); } } //for logging


function changeMasterPasswordAdmin(newMaster) { try  //for logging
{
  log('Diagnostic', 'Change master password');
  var pek = getP_PEK(true);
  if (newMaster != null) {
    setP_EEK(encrypt(pek, newMaster));
    var ui = SpreadsheetApp.getUi();
    ui.alert('Success !', 'Your master password has been changed successfully', ui.ButtonSet.OK);
  }
} catch(e) { handleError(e); } } //for logging


function resetSpreadheetAdminOk(newMaster) { try  //for logging
{
  if (newMaster != null) {
      initializeProperties(false);
      var newEncryptionKey = generateEncryptionKey();
      setP_EEK(encrypt(newEncryptionKey, newMaster));
      setP_IsKeystoreReady(true);
      onOpen();
      unlockSpreasheet(newMaster);
  }
} catch(e) { handleError(e); } } //for logging

function resetSpreadheetAdminCancel() { try  //for logging
{
  removeAllProperties();
} catch(e) { handleError(e); } } //for logging


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
} catch(e) { handleError(e); } } //for logging

function startAutoLockTrigger()  { try  //for logging
{
  var autolockDelay = getP_AutolockDelay();
  if (autolockDelay > 0) {
    var delay = Math.ceil(Math.max(autolockDelay / 10, 1));
    ScriptApp.newTrigger("checkAutolock").timeBased().everyMinutes(delay).create();
  }
} catch(e) { handleError(e); } } //for logging
