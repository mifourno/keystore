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

//##################################
//##        SHEET EVENTS
//##################################

function onOpen()  { try  //for logging
{
  initializeProperties(true);
  updateMenuEntries();
  lockSpreasheet('onOpen');
} catch(e) { logError(e); throw(e); } } //for logging

function onEdit(event) { try  //for logging
{
  setP_LastUpdate(new Date());
} catch(e) { logError(e); throw(e); } } //for logging


//##################################
//##        SHOW/HIDE SHEETS
//##################################

//function hideSheet(name) { try  //for logging
//{
//  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).hideSheet();
//  setProperty('Sheet' + name + 'Visible', false);
//  updateMenuEntries();
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function showSheet(name) { try  //for logging
//{
//  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).activate();
//  setProperty('Sheet' + name + 'Visible', true);
//  updateMenuEntries();
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function hideHelp() { try  //for logging
//{
//  hideSheet('Help');
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function showHelp() { try  //for logging
//{
//  showSheet('Help');
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function hideSettings() { try  //for logging
//{
//  hideSheet('Settings');
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function showSettings() { try  //for logging
//{
//  showSheet('Settings');
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function hideLogs() { try  //for logging
//{
//  hideSheet('Logs');
//} catch(e) { logError(e); throw(e); } } //for logging
//
//function showLogs() { try  //for logging
//{
//  showSheet('Logs');
//} catch(e) { logError(e); throw(e); } } //for logging



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
} catch(e) { logError(e); throw(e); } } //for logging

function updateMenuEntries()  { try  //for logging
{
  var myMenu = SpreadsheetApp.getUi().createMenu(getP_ProgramName());
  
  myMenu.addItem('♜ Lock immediately', 'manualLockSpreasheet');
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
  
//  if (getProperty('Sheet' + 'Help' + 'Visible') == 'true') myMenu.addItem('✖ Hide Help', 'hideHelp');
//  else myMenu.addItem('� Help', 'showHelp');
//  if (getProperty('Sheet' + 'Settings' + 'Visible') == 'true') myMenu.addItem('✖ Hide Settings', 'hideSettings');
//  else myMenu.addItem('⚙ Settings', 'showSettings');
//  if (getProperty('Sheet' + 'Logs' + 'Visible') == 'true') myMenu.addItem('✖ Hide Logs', 'hideLogs');
//  else myMenu.addItem('✎ Logs (advanced)', 'showLogs');
//  myMenu.addSeparator();
  
  myMenu.addItem('☠ Reset this spreadsheet', 'resetSpreasheet');
  myMenu.addToUi();
  
} catch(e) { logError(e); throw(e); } } //for logging

