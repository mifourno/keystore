/* ----------------------------------------------------------
Properties.gs -- 
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


// ################ User Properties #################
function getP_PEK() { return PropertiesService.getUserProperties().getProperty('PEK'); }
function setP_PEK(value) { PropertiesService.getUserProperties().setProperty('PEK', value); }

function getP_EEK() { return PropertiesService.getUserProperties().getProperty('EEK'); }
function setP_EEK(value) { PropertiesService.getUserProperties().setProperty('EEK', value); }

function getP_LastUpdate() { return PropertiesService.getUserProperties().getProperty('LastUpdate'); }
function setP_LastUpdate(value) { PropertiesService.getUserProperties().setProperty('LastUpdate', value); }

function getP_LockedAt() { return PropertiesService.getUserProperties().getProperty('LockedAt'); }
function setP_LockedAt(value) { PropertiesService.getUserProperties().setProperty('LockedAt', value); }

function getP_PasswordLength() { return PropertiesService.getUserProperties().getProperty('PasswordLength'); }
function setP_PasswordLength(value) { PropertiesService.getUserProperties().setProperty('PasswordLength', value); }

function getP_PasswordChars() { return PropertiesService.getUserProperties().getProperty('PasswordChars'); }
function setP_PasswordChars(value) { PropertiesService.getUserProperties().setProperty('PasswordChars', value); }

function getP_RevealMode() { return PropertiesService.getUserProperties().getProperty('RevealMode'); }
function setP_RevealMode(value) { PropertiesService.getUserProperties().setProperty('RevealMode', value); }

// ################ Script Properties #################
function getP_ProgramName() { return PropertiesService.getScriptProperties().getProperty('ProgramName'); }
function setP_ProgramName(value) { PropertiesService.getScriptProperties().setProperty('ProgramName', value); }

function getP_ProtectionMessage() { return PropertiesService.getScriptProperties().getProperty('ProtectionMessage'); }
function setP_ProtectionMessage(value) { PropertiesService.getScriptProperties().setProperty('ProtectionMessage', value); }

// ################ Document Properties #################
function getP_RevealedRangeSheets() { return PropertiesService.getDocumentProperties().getProperty('RevealedRangeSheets'); }
function setP_RevealedRangeSheets(value) { PropertiesService.getDocumentProperties().setProperty('RevealedRangeSheets', value); }

function getP_AutolockDelay() { return Number(PropertiesService.getDocumentProperties().getProperty('AutolockDelay')); }
function setP_AutolockDelay(value) { PropertiesService.getDocumentProperties().setProperty('AutolockDelay', value); }

function getP_SetFormatAtEncryption() { return (PropertiesService.getDocumentProperties().getProperty('SetFormatAtEncryption') === 'true'); }
function setP_SetFormatAtEncryption(value) { PropertiesService.getDocumentProperties().setProperty('SetFormatAtEncryption', value); }

function getP_GenPassPunctuations() { return PropertiesService.getDocumentProperties().getProperty('GenPassPunctuations'); }
function setP_GenPassPunctuations(value) { PropertiesService.getDocumentProperties().setProperty('GenPassPunctuations', value); }

function getP_GenPassSymbols() { return PropertiesService.getDocumentProperties().getProperty('GenPassSymbols'); }
function setP_GenPassSymbols(value) { PropertiesService.getDocumentProperties().setProperty('GenPassSymbols', value); }

function getP_GenPassAlpha() { return PropertiesService.getDocumentProperties().getProperty('GenPassAlpha'); }
function setP_GenPassAlpha(value) { PropertiesService.getDocumentProperties().setProperty('GenPassAlpha', value); }

function getP_GenPassNum() { return PropertiesService.getDocumentProperties().getProperty('GenPassNum'); }
function setP_GenPassNum(value) { PropertiesService.getDocumentProperties().setProperty('GenPassNum', value); }

function getP_EncryptedFormat_Background() { return PropertiesService.getDocumentProperties().getProperty('EF_BG'); }
function setP_EncryptedFormat_Background(value) { PropertiesService.getDocumentProperties().setProperty('EF_BG', value); }

function getP_EncryptedFormat_Color() { return PropertiesService.getDocumentProperties().getProperty('EF_COL'); }
function setP_EncryptedFormat_Color(value) { PropertiesService.getDocumentProperties().setProperty('EF_COL', value); }

function getP_RevealedFormat_Background() { return PropertiesService.getDocumentProperties().getProperty('RF_BG'); }
function setP_RevealedFormat_Background(value) { PropertiesService.getDocumentProperties().setProperty('RF_BG', value); }

function getP_RevealedFormat_Color() { return PropertiesService.getDocumentProperties().getProperty('RF_COL'); }
function setP_RevealedFormat_Color(value) { PropertiesService.getDocumentProperties().setProperty('RF_COL', value); }

function getP_DecryptedFormat_Background() { return PropertiesService.getDocumentProperties().getProperty('DF_BG'); }
function setP_DecryptedFormat_Background(value) { PropertiesService.getDocumentProperties().setProperty('DF_BG', value); }

function getP_DecryptedFormat_Color() { return PropertiesService.getDocumentProperties().getProperty('DF_COL'); }
function setP_DecryptedFormat_Color(value) { PropertiesService.getDocumentProperties().setProperty('DF_COL', value); }


function initializeProperties(onlyIfNotExist)  { try  //for logging
{
  var userProperties = PropertiesService.getUserProperties();
  var scriptProperties = PropertiesService.getScriptProperties();
  var documentProperties = PropertiesService.getDocumentProperties();
  var defautSettings = getDefaultSettings();
  // User Properties
  initializePropertyIfNotExist(userProperties, 'PEK', '', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'EEK', '', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'LastUpdate', new Date(), onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'LockedAt', new Date(), onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'PasswordLength', '20', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'PasswordChars', '4', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'RevealMode', 'symbols', onlyIfNotExist);
  
  // Script Properties
  initializePropertyIfNotExist(scriptProperties, 'ProgramName', 'Keystore', onlyIfNotExist);
  initializePropertyIfNotExist(scriptProperties, 'ProtectionMessage', 'Marked as sensitive', onlyIfNotExist);
  
  // Document Properties
  initializePropertyIfNotExist(documentProperties, 'RevealedRangeSheets', '', onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'AutolockDelay', defautSettings.autolockDelay, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'SetFormatAtEncryption', defautSettings.setFormatAtEncryption, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassPunctuations', defautSettings.genPassPunctuations, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassSymbols', defautSettings.genPassSymbols, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassAlpha', defautSettings.genPassAlpha, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassNum', defautSettings.genPassNum);
  
  initializePropertyIfNotExist(documentProperties, 'EF_BG', defautSettings.EF_BG, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'EF_COL', defautSettings.EF_COL, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'RF_BG', defautSettings.RF_BG, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'RF_COL', defautSettings.RF_COL, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'DF_BG', defautSettings.DF_BG, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'DF_COL', defautSettings.DF_COL, onlyIfNotExist);
  
  syncLineThroughAndProtection();
  
} catch(e) { logError(e); throw(e); } } //for logging


function initializePropertyIfNotExist(propertySet, name, value, onlyIfNotExist)  { try  //for logging
{
  var propertyValue = propertySet.getProperty(name);
  if (!onlyIfNotExist || propertyValue === 'undefined' || propertyValue == null || propertyValue == '') propertySet.setProperty(name, value);
} catch(e) { logError(e); throw(e); } } //for logging

function getSettings()  { try  //for logging
{
  var settings = {
    autolockDelay: getP_AutolockDelay(),
    EF_BG: getP_EncryptedFormat_Background(),
    EF_COL: getP_EncryptedFormat_Color(),
    RF_BG: getP_RevealedFormat_Background(),
    RF_COL: getP_RevealedFormat_Color(),
    DF_BG: getP_DecryptedFormat_Background(),
    DF_COL: getP_DecryptedFormat_Color(),
    setFormatAtEncryption : getP_SetFormatAtEncryption(),
    genPassPunctuations: getP_GenPassPunctuations(),
    genPassSymbols: getP_GenPassSymbols(),
    genPassAlpha: getP_GenPassAlpha(),
    genPassNum: getP_GenPassNum()
  };
  return settings;
} catch(e) { logError(e); throw(e); } } //for logging

function saveSettings(settings)  { try  //for logging
{
  setP_AutolockDelay(settings.autolockDelay);
  setP_EncryptedFormat_Background(settings.EF_BG);
  setP_EncryptedFormat_Color(settings.EF_COL);
  setP_RevealedFormat_Background(settings.RF_BG);
  setP_RevealedFormat_Color(settings.RF_COL);
  setP_DecryptedFormat_Background(settings.DF_BG);
  setP_DecryptedFormat_Color(settings.DF_COL);
  setP_SetFormatAtEncryption(settings.setFormatAtEncryption);
  setP_GenPassPunctuations(settings.genPassPunctuations);
  setP_GenPassSymbols(settings.genPassSymbols);
  setP_GenPassAlpha(settings.genPassAlpha);
  setP_GenPassNum(settings.genPassNum);
} catch(e) { logError(e); throw(e); } } //for logging


function getDefaultSettings()  { try  //for logging
{
  var settings = {
    autolockDelay: 5,
    EF_BG: '#f4cccc',
    EF_COL: '#990000',
    RF_BG: '#fce5cd',
    RF_COL: '#b45f06',
    DF_BG: '#d9ead3',
    DF_COL: '#38761d',
    setFormatAtEncryption : true,
    genPassPunctuations: ',.;:()[]{}/\\',
    genPassSymbols: '!#$%&*+-<=>?@^_|~',
    genPassAlpha: 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ',
    genPassNum: '0123456789'
  };
  return settings;
} catch(e) { logError(e); throw(e); } } //for logging


function getPreferences()  { try  //for logging
{
  var settings = {
    passwordLength: getP_PasswordLength(),
    passwordChars: getP_PasswordChars(),
    revealMode: getP_RevealMode()
  };
  return settings;
} catch(e) { logError(e); throw(e); } } //for logging

function savePreferences(preferences)  { try  //for logging
{
  setP_PasswordLength(preferences.passwordLength);
  setP_PasswordChars(preferences.passwordChars);
  setP_RevealMode(preferences.revealMode);
} catch(e) { logError(e); throw(e); } } //for logging

