/* ----------------------------------------------------------
Menu.gs -- 
Copyright (C) 2015 Mifourno

This software may be modified and distributed under the terms
of the MIT license.  See the LICENSE file for details.

GitHub: https://github.com/mifourno/keystore/
Contact: mifourno@gmail.com
---------------------------------------------------------- */

/**
 * @OnlyCurrentDoc
 */


//##################################
//##    ENABLING AND AUTHORIZING
//##################################

function enableKeystore() { try  //for logging
{
  readCurrentUser();
  var isCurrentUserEditor = isEditor();
  if (isCurrentUserEditor == false) {
    serverSideAlert('Enabling keystore', 'You must have edit permission on the spreadsheet to enable Keystore.');
    return;
  } else if (isCurrentUserEditor == null) {
    updateInitMasterPasswordMenuEntries();
    serverSideAlert('Keystore is enabled !', 'You now need to initialize your master password. Please open the menu Add-on / Keystore / Set master password');
    return;
  } else {
    promptMasterPassword('init', 'initSpreadheetAdminOk', 'initSpreadheetAdminCancel');
  }
} catch(e) { handleError(e); } } //for logging


function checkAuthorization() { try  //for logging
{
  var userEmail = readCurrentUser();
  if (isNullOrWS(userEmail)) {
    updateChangeMasterPasswordMenuEntries();
  } else {
    startKeystore(userEmail);
  }
  return userEmail;
} catch(e) { handleError(e); } } //for logging


function authorizeKeystore() { try  //for logging
{
  var userEmail = readCurrentUser();
  if (isNullOrWS(userEmail)) {
    updateChangeMasterPasswordMenuEntries();
    showCheckAuthorizationSideBar();
    customDialog('Master password', "This encrypted spreadsheet has been shared with you and it's the first time that you open it. Before you use it, you must change your master password.\n(intial password must have been sent to you)\n\nPlease open the menu\nAdd-ons / Keystore / Change master password", 130);
  } else {
    startKeystore(userEmail);
  }
} catch(e) { handleError(e); } } //for logging


function startKeystore(userEmail) { try  //for logging
{
  initializeProperties(true);
  updateMenuEntries();
  lockSpreasheet('onOpen');
  showSideMenu(true);
  checkConsistancyAllSheets();
} catch(e) { handleError(e); } } //for logging


function disableKeystore()  { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can disable keystore on this spreadsheet.");
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'WARNING !',
     'If you continue, this spreadsheet will remain as it is but encrypted data will be impossible to recover. Are you sure you want to continue ?',
      ui.ButtonSet.YES_NO);
  if (result == ui.Button.YES) {
    removeAllProtections();
    removeAllProperties();
    onOpen();
  }
} catch(e) { handleError(e); } } //for logging


function initSpreadheetAdminOk(newMaster) { try  //for logging
{
  if (newMaster != null) {
    var userEmail = getP_CurrentUser();
    initializeProperties(false);
    var newEncryptionKey = generateEncryptionKey();
    setP_EEK(userEmail, encrypt(newEncryptionKey, newMaster));
    setP_IsKeystoreReady(true);
    onOpen();
    unlockSpreasheet(newMaster);
    serverSideAlert('READY !', "You're all set ! You can start encrypting cells. Use the side-bar buttons for that.");
  }
} catch(e) { handleError(e); } } //for logging

function initSpreadheetAdminCancel() { try  //for logging
{
  removeAllProperties();
} catch(e) { handleError(e); } } //for logging


function readCurrentUser() { try  //for logging
{
  var userEmail = Session.getActiveUser().getEmail();
  if (!isNullOrWS(userEmail)) setP_CurrentUser(userEmail);
  return getP_CurrentUser();
} catch(e) { handleError(e); } } //for logging


//##################################
//##        SHEET EVENTS
//##################################

function onOpen(event) { try  //for logging
{

  // CASE 1:
  // NONE or !IsKeystoreReady => User previously installed Add-on, but did not enable on the current document
  //   We just want to show menu "Enable Keystore"

  // CASE 2:
  // LIMITED and !IsKeystoreReady and unknown user => user clicked once on menu "Enable Keystore", he was prompt for authorization and just clicked Allow, but for some reason (ex bound script) the user is still unknown
  // if (ScriptError: not authorized yet)  it can happen if user revoke all rights after keystore was enabled
  //     We want to invite the user to authorize keystore by clicking the menu "Authorise keystore" (to prompt the authorization dialog)
  // else
  //     We want to invite the user to initialize his master password by clicking the menu "Initialize master password" (so that we can know the user email)

  // CASE 3:
  // LIMITED and !IsKeystoreReady and user is known => user clicked once on menu "Enable Keystore", he was prompt for authorization and just clicked Allow, and luckilly the user is known (ex add-on)
  // We can directly open the init password dialog

  // CASE 4:
  // LIMITED and IsKeystoreReady and unknown user
  // if (ScriptError: not authorized yet)  example: spreadsheet was shared with this user, thus the script is enabled, but the user didn't give his authorization yet
  //     We want to invite the user to authorize keystore by clicking the menu "Authorise keystore" (to prompt the authorization dialog)
  // else
  //     We want to invite the user to initialize his master password by clicking the menu "Initialize master password" (so that we can know the user email)
  
  // CASE 5:
  // LIMITED and authorized and user is known
  // We show SideMenu


  var currentUser = readCurrentUser();
  var isKeystoreReady = getP_IsKeystoreReady();
  
  if ((!isNullOrWS(event) && event.authMode.toString() == 'NONE') || !isKeystoreReady) {
    // CASE 1
    updateInitMenuEntries();
  } else {
    updateAuthorizationMenuEntries();
    showCheckAuthorizationSideBar();
  }
  
} catch(e) { handleError(e); } } //for logging


function onInstall(event) {
  onOpen(event);
  // Perform additional setup as needed.
}


function onEdit(event) { try  //for logging
{
  if (getP_IsKeystoreReady()) {
    setP_LastUpdate(new Date());
    if (event.range != null) {
      var sheet = event.range.getSheet();
      var protectionParams = getProtectionParams();
      var protectionDictionary = getProtectionDictionary(sheet, protectionParams);
      var pek = getP_PEK(false);
      for (var row = event.range.getRowIndex(); row <= event.range.getLastRow(); ++row) 
      {
        for (var col = event.range.getColumnIndex(); col <= event.range.getLastColumn(); ++col)
        {
          checkConsistancySingleCell(sheet.getRange(row,col), pek, protectionParams, protectionDictionary);
        }
      }
    }
  }
} catch(e) { handleError(e); } } //for logging



//##################################
//##     SHOW/HIDE SIDE MENUS
//##################################

function showSettings() { try  //for logging
{
  var ui = HtmlService.createHtmlOutputFromFile('Settings')
      .setTitle('Keystore settings')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
} catch(e) { handleError(e); } } //for logging

function showSharing(email, canEdit, isError) { try  //for logging
{
  if (!isEditor()) throw new Error("Only people with editor's permissions can share this spreadsheet");
  
  var html = HtmlService.createTemplateFromFile('Sharing');
  html.data = { email : email, canEdit : canEdit, isError : isError };
  sideMenu = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('Sharing');
  SpreadsheetApp.getUi().showSidebar(sideMenu);
  
} catch(e) { handleError(e); } } //for logging

function showSideMenu(visible) { try  //for logging
{
  if (typeof visible == 'undefined') visible = true;
  var htmlFile = 'SideMenu';
  if (!visible) htmlFile = 'CloseSideMenu';
  
  var html = HtmlService.createTemplateFromFile(htmlFile);
  html.data = { isEditor : isEditor() };
  sideMenu = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('Keystore');
  SpreadsheetApp.getUi().showSidebar(sideMenu);
  
} catch(e) { handleError(e); } } //for logging

function showCheckAuthorizationSideBar() { try  //for logging
{
  var ui = HtmlService.createHtmlOutputFromFile('CheckAuthorization').setTitle('Keystore').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
} catch(e) { handleError(e); } } //for logging



//##################################
//##        MENU / SUBMENU
//##################################

function updateInitMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu('Keystore');
  myMenu.addItem('Enable Keystore', 'enableKeystore');
  myMenu.addToUi();
} catch(e) { handleError(e); } } //for logging

function updateAuthorizationMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu('Keystore');
  myMenu.addItem('Authorize Keystore...', 'authorizeKeystore');
  myMenu.addToUi();
} catch(e) { handleError(e); } } //for logging

function updateInitMasterPasswordMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu('Keystore');
  myMenu.addItem('Set master password', 'enableKeystore');
  myMenu.addToUi();
} catch(e) { handleError(e); } } //for logging

function updateChangeMasterPasswordMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu('Keystore');
  myMenu.addItem('Change master password', 'changeMasterPassword');
  myMenu.addToUi();
} catch(e) { handleError(e); } } //for logging

function updateMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu('Keystore');
  
  myMenu.addItem('♜ Lock now', 'manualLockSpreasheet');
  myMenu.addItem('♺ Change master password', 'changeMasterPassword');
  myMenu.addSeparator();
  
  myMenu.addItem('⌦ Encrypt selection', 'markAsSensitive');
  myMenu.addSeparator();
  
  myMenu.addItem('⌫ Reveal: show in popup', 'revealPopup');
  myMenu.addItem('⌫ Reveal: decrypt the cell(s)', 'removeEncryption');
  myMenu.addSeparator();
  
  myMenu.addSubMenu(SpreadsheetApp.getUi().createMenu('Generate password')
                    .addSubMenu(genPassSubMenu(20))
                    .addSubMenu(genPassSubMenu(16))
                    .addSubMenu(genPassSubMenu(12))
                    .addSubMenu(genPassSubMenu(10))
                    .addSubMenu(genPassSubMenu(8))
                    .addSubMenu(genPassSubMenu(6))
                    .addSubMenu(genPassSubMenu(4))
                    .addSeparator()
                    .addSubMenu(genPassSubMenu(256))
                    .addSubMenu(genPassSubMenu(128))
                    .addSubMenu(genPassSubMenu(64)));
  myMenu.addSeparator();

  myMenu.addItem('≣ Show side menu', 'showSideMenu');
  if (isEditor() == true) myMenu.addItem('✉ Share', 'showSharing');
  myMenu.addItem('⚙ Settings', 'showSettings');
  if (isEditor() == true) {
    myMenu.addSeparator();
    myMenu.addItem('✖ Disable Keystore', 'disableKeystore');
  }
  
  myMenu.addToUi();
  
} catch(e) { handleError(e); } } //for logging

function genPassSubMenu(length)  { try  //for logging
{
  return SpreadsheetApp.getUi().createMenu('Length ' + length)
                                .addItem('With punctuation - Abc123#@%{}[];:,', 'genPassPunct' + length)
                                .addItem('With symbols - Abc123#@%', 'genPassSymb' + length)
                                .addItem('AlphaNum - Abc123', 'genPassAlphaNum' + length)
                                .addItem('Alpha - Abc', 'genPassAlpha' + length)
                                .addItem('Numerical - 123', 'genPassNum' + length)
} catch(e) { handleError(e); } } //for logging





