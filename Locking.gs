/* ----------------------------------------------------------
Locking.gs -- 
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
  var userEmail = getP_CurrentUser();
  
  var pek = decrypt(getP_EEK(userEmail), masterPassword);
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
  setP_PEK('');
  setP_LockedAt(new Date());
  tryRemoveAllTriggers();
} catch(e) { handleError(e); } } //for logging


//##################################
//##    CHANGE MASTER PASSWORD
//##################################


function changeMasterPassword() { try  //for logging
{
  readCurrentUser();
  promptMasterPassword('change', 'changeMasterPasswordAdmin');
} catch(e) { handleError(e); } } //for logging


function changeMasterPasswordAdmin(newMaster) { try  //for logging
{
  log('Diagnostic', 'Change master password');
  var userEmail = getP_CurrentUser();
  var pek = getP_PEK(true);
  if (newMaster != null) {
    setP_EEK(userEmail, encrypt(pek, newMaster));
    serverSideAlert('Success !', 'Your master password has been changed successfully');
  }
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
    var delay = 1;
    switch (autolockDelay) {
      case 3:
      case 5:
      case 8:
        delay = 1;
        break;
      case 15:
      case 30:
      case 60:
      default:
        delay = 5;
        break;
    }
    //var delay = Math.ceil(Math.max(autolockDelay / 10, 1));
    ScriptApp.newTrigger("checkAutolock").timeBased().everyMinutes(delay).create();
  }
} catch(e) { handleError(e); } } //for logging



//##################################
//##         SHARING
//##################################

function getUsers() { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can share this spreadsheet");
  
  var editors = SpreadsheetApp.getActiveSpreadsheet().getEditors();
  var viewers = SpreadsheetApp.getActiveSpreadsheet().getViewers();
  var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  
  var allKeys = new Array();
  var allUsers = [];
  for (var i = 0; i < editors.length; i++) {
    allUsers.push(getUserPoco(editors[i], true, owner));
    allKeys[editors[i].getEmail()] = true;
  }
  for (var i = 0; i < viewers.length; i++) {
    if (allKeys[viewers[i].getEmail()] != true) allUsers.push(getUserPoco(viewers[i], false));
  }
  return allUsers;
} catch(e) { handleError(e); } } //for logging

function addUser(email, canEdit) { try  //for logging
{
  if (isLocked()) promptMasterPassword('assert', 'addUserAdmin', null, email + ',' + canEdit);
  else addUserAdmin(email + ',' + canEdit);
} catch(e) { handleError(e); } } //for logging


function addUserAdmin(params) { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can share this spreadsheet");
  var pek = getP_PEK(true);
  
  var paramSplited = params.split(',');
  var email = paramSplited[0];
  var canEdit = paramSplited[1] == 'true';
  
  try
  {
    if (canEdit) SpreadsheetApp.getActiveSpreadsheet().addEditor(email);
    else SpreadsheetApp.getActiveSpreadsheet().addViewer(email);
  } catch(ex) {
    showSharing(email, canEdit, true);
    return;
  }
  var charset = getP_GenPassNum() + getP_GenPassAlpha();
  var newUserMaster = genNewPassword(16, charset);
  setP_EEK(email, encrypt(pek, newUserMaster));
  
  serverSideAlert('Success !', 'Your have just shared this file with ' + email + '\nTo be able to decrypt data he or she will need this temporary password:\n\n'+ newUserMaster + '\n\nThe person will be asked to change this password on the first connection.');
  showSharing();
} catch(e) { handleError(e); } } //for logging

function updateRights(email, canEdit) { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can share this spreadsheet");
  
  if (canEdit) SpreadsheetApp.getActiveSpreadsheet().addEditor(email);
  else { 
    SpreadsheetApp.getActiveSpreadsheet().removeEditor(email);
    SpreadsheetApp.getActiveSpreadsheet().addViewer(email);
  }
} catch(e) { handleError(e); } } //for logging

function removeUser(email, canEdit) { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can share this spreadsheet");
  
  deleteP_EEK(email);
  SpreadsheetApp.getActiveSpreadsheet().removeViewer(email);
} catch(e) { handleError(e); } } //for logging


function getUserPoco(user, canEdit, owner) { try  //for logging
{
  var email = user.getEmail();
  return { email : email, canEdit : canEdit, isOwner : email == owner };
} catch(e) { handleError(e); } } //for logging
