/* ----------------------------------------------------------
Menu.gs -- 
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



//##################################
//##        SHEET EVENTS
//##################################

function onOpen()  { try  //for logging
{
  initializeProperties(true);
  updateMenuEntries();
  lockSpreasheet('onOpen');
} catch(e) { handleError(e); } } //for logging

function onEdit(event) { try  //for logging
{
  setP_LastUpdate(new Date());
} catch(e) { handleError(e); } } //for logging


//##################################
//##        SHOW/HIDE SIDE BARS
//##################################

function showSettings() { try  //for logging
{
  var ui = HtmlService.createHtmlOutputFromFile('Settings')
      .setTitle(getP_ProgramName() + ' settings')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
} catch(e) { handleError(e); } } //for logging

function showSideMenu() { try  //for logging
{
  var ui = HtmlService.createHtmlOutputFromFile('SideMenu')
      .setTitle(getP_ProgramName())
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
} catch(e) { handleError(e); } } //for logging



//##################################
//##        INITIALIZATION
//##################################

function genPassSubMenu(length)  { try  //for logging
{
  return SpreadsheetApp.getUi().createMenu('Length ' + length)
                                .addItem('With punctuation - Abc123#@%{}[];:,', 'genPassPunct' + length)
                                .addItem('With symbols - Abc123#@%', 'genPassSymb' + length)
                                .addItem('AlphaNum - Abc123', 'genPassAlphaNum' + length)
                                .addItem('Alpha - Abc', 'genPassAlpha' + length)
                                .addItem('Numerical - 123', 'genPassNum' + length)
} catch(e) { handleError(e); } } //for logging

function updateMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu(getP_ProgramName());
  
  myMenu.addItem('♜ Lock now', 'manualLockSpreasheet');
  myMenu.addItem('♺ Change master-password', 'changeMasterPassword');
  myMenu.addSeparator();
  
  myMenu.addItem('⌦ Encrypt selection', 'markAsSensitive');
  myMenu.addSeparator();
  
  myMenu.addItem('⌫ Reveal in popup', 'revealPopup');
  myMenu.addItem('⌫ Reveal just for a minute', 'revealFewSeconds');
  myMenu.addItem('⌫ Reveal until next lock', 'revealUntilLock');
  myMenu.addItem('⌫ Reveal permanently', 'removeEncryption');
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
  myMenu.addItem('⚙ Settings', 'showSettings');
  myMenu.addSeparator();
    
  myMenu.addItem('☠ Reset this spreadsheet', 'resetSpreasheet');
  myMenu.addToUi();
  
} catch(e) { handleError(e); } } //for logging

