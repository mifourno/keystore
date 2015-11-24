/* ----------------------------------------------------------
Properties.gs -- 
Copyright (C) 2015 Mifourno

This software may be modified and distributed under the terms
of the MIT license.  See the LICENSE file for details.

GitHub: https://github.com/mifourno/keystore/
Contact: mifourno@gmail.com
---------------------------------------------------------- */

/**
 * @OnlyCurrentDoc
 */


// ################ User Properties #################
function getP_PEK(required) { 
  var pek = PropertiesService.getUserProperties().getProperty('PEK'); 
  if (required == true && isNullOrWS(pek)) throw new Error('Encryption key is empty. Impossible to complete action.');
  return pek;
}
function setP_PEK(value) { PropertiesService.getUserProperties().setProperty('PEK', value); }

function getP_CurrentUser() { return PropertiesService.getUserProperties().getProperty('CurrentUser'); }
function setP_CurrentUser(value) { PropertiesService.getUserProperties().setProperty('CurrentUser', value); }

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
         
function getP_AutoEncryptNewPassword() { return (PropertiesService.getUserProperties().getProperty('AutoEncryptNewPassword') === 'true'); }
function setP_AutoEncryptNewPassword(value) { PropertiesService.getUserProperties().setProperty('AutoEncryptNewPassword', value); }


// ################ Script Properties #################
function getP_ProgramName() { return PropertiesService.getScriptProperties().getProperty('ProgramName'); }
function setP_ProgramName(value) { PropertiesService.getScriptProperties().setProperty('ProgramName', value); }

function getP_ProtectionMessage() { return PropertiesService.getScriptProperties().getProperty('ProtectionMessage'); }
function setP_ProtectionMessage(value) { PropertiesService.getScriptProperties().setProperty('ProtectionMessage', value); }

// ################ Document Properties #################

function getP_EEK(userEmail) { 
  var eekArray = JSON.parse(PropertiesService.getDocumentProperties().getProperty('EEK'));
  if (eekArray != null) {
    for (var i = 0; i < eekArray.length; i++) {
      if (eekArray[i].userEmailLowered == userEmail.toLowerCase()) return eekArray[i].EEK;
    }
  }
  return null;
}
function setP_EEK(userEmail, value) {
  var eekArray = deleteP_EEK(userEmail);
  eekArray.push({ 'userEmailLowered' : userEmail.toLowerCase(), 'EEK' : value });
  PropertiesService.getDocumentProperties().setProperty('EEK', JSON.stringify(eekArray)); 
}
function deleteP_EEK(userEmail) {
  var allEEK = PropertiesService.getDocumentProperties().getProperty('EEK');
  var eekArray = null;
  if (isNullOrWS(allEEK)) eekArray = [];
  else eekArray = JSON.parse(allEEK);
  if (eekArray != null) {
    for (var i = 0; i < eekArray.length; i++) {
      if (eekArray[i].userEmailLowered == userEmail.toLowerCase()) eekArray.splice(i--,1);
    }
  }
  PropertiesService.getDocumentProperties().setProperty('EEK', JSON.stringify(eekArray));
  return eekArray;
}

function getP_Sharing() { return PropertiesService.getDocumentProperties().getProperty('Sharing'); }
function setP_Sharing(value) { PropertiesService.getDocumentProperties().setProperty('Sharing', value); }

function getP_IsKeystoreReady() { return (PropertiesService.getDocumentProperties().getProperty('IsKeystoreReady') === 'true'); }
function setP_IsKeystoreReady(value) { PropertiesService.getDocumentProperties().setProperty('IsKeystoreReady', value); }

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

function getP_InitFormat_Background() { return PropertiesService.getDocumentProperties().getProperty('IF_BG'); }
function setP_InitFormat_Background(value) { PropertiesService.getDocumentProperties().setProperty('IF_BG', value); }

function getP_InitFormat_Color() { return PropertiesService.getDocumentProperties().getProperty('IF_COL'); }
function setP_InitFormat_Color(value) { PropertiesService.getDocumentProperties().setProperty('IF_COL', value); }


function initializeScriptProperties(onlyIfNotExist)  { try  //for logging
{
  var scriptProperties = PropertiesService.getScriptProperties();
  // Script Properties
  initializePropertyIfNotExist(scriptProperties, 'ProgramName', 'Keystore', onlyIfNotExist);
  initializePropertyIfNotExist(scriptProperties, 'ProtectionMessage', 'Encrypted', onlyIfNotExist);
} catch(e) { handleError(e); } } //for logging

function initializeProperties(onlyIfNotExist)  { try  //for logging
{
  var userProperties = PropertiesService.getUserProperties();
  var documentProperties = PropertiesService.getDocumentProperties();
  var defautSettings = getDefaultSettings();
  // User Properties
  initializePropertyIfNotExist(userProperties, 'PEK', '', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'CurrentUser', '', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'LastUpdate', new Date(), onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'LockedAt', new Date(), onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'PasswordLength', '20', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'PasswordChars', '4', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'RevealMode', 'symbols', onlyIfNotExist);
  initializePropertyIfNotExist(userProperties, 'AutoEncryptNewPassword', 'true', onlyIfNotExist);
  
  // Document Properties
  initializePropertyIfNotExist(documentProperties, 'EEK', '', onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'Sharing', '', onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'RevealedRangeSheets', '', onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'AutolockDelay', defautSettings.autolockDelay, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'SetFormatAtEncryption', defautSettings.setFormatAtEncryption, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassPunctuations', defautSettings.genPassPunctuations, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassSymbols', defautSettings.genPassSymbols, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassAlpha', defautSettings.genPassAlpha, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'GenPassNum', defautSettings.genPassNum);
  
  initializePropertyIfNotExist(documentProperties, 'EF_BG', defautSettings.EF_BG, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'EF_COL', defautSettings.EF_COL, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'IF_BG', defautSettings.IF_BG, onlyIfNotExist);
  initializePropertyIfNotExist(documentProperties, 'IF_COL', defautSettings.IF_COL, onlyIfNotExist);
  
} catch(e) { handleError(e); } } //for logging


function initializePropertyIfNotExist(propertySet, name, value, onlyIfNotExist)  { try  //for logging
{
  var propertyValue = propertySet.getProperty(name);
  if (!onlyIfNotExist || typeof propertyValue == 'undefined' || propertyValue == null || propertyValue == '') propertySet.setProperty(name, value);
} catch(e) { handleError(e); } } //for logging

function removeAllProperties()  { try  //for logging
{
  PropertiesService.getUserProperties().deleteAllProperties();
  //PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getDocumentProperties().deleteAllProperties();
} catch(e) { handleError(e); } } //for logging

function getSettings()  { try  //for logging
{
  var settings = {
    autolockDelay: getP_AutolockDelay(),
    EF_BG: getP_EncryptedFormat_Background(),
    EF_COL: getP_EncryptedFormat_Color(),
    IF_BG: getP_InitFormat_Background(),
    IF_COL: getP_InitFormat_Color(),
    setFormatAtEncryption : getP_SetFormatAtEncryption(),
    genPassPunctuations: getP_GenPassPunctuations(),
    genPassSymbols: getP_GenPassSymbols(),
    genPassAlpha: getP_GenPassAlpha(),
    genPassNum: getP_GenPassNum()
  };
  return settings;
} catch(e) { handleError(e); } } //for logging

function saveSettings(settings)  { try  //for logging
{
  setP_AutolockDelay(settings.autolockDelay);
  setP_EncryptedFormat_Background(settings.EF_BG);
  setP_EncryptedFormat_Color(settings.EF_COL);
  setP_InitFormat_Background(settings.IF_BG);
  setP_InitFormat_Color(settings.IF_COL);
  setP_SetFormatAtEncryption(settings.setFormatAtEncryption);
  setP_GenPassPunctuations(settings.genPassPunctuations);
  setP_GenPassSymbols(settings.genPassSymbols);
  setP_GenPassAlpha(settings.genPassAlpha);
  setP_GenPassNum(settings.genPassNum);
} catch(e) { handleError(e); } } //for logging


function getDefaultSettings()  { try  //for logging
{
  var settings = {
    autolockDelay: 5,
    EF_BG: '#f4cccc',
    EF_COL: '#990000',
    //RF_BG: '#fce5cd',
    //RF_COL: '#b45f06',
    //DF_BG: '#d9ead3',
    //DF_COL: '#38761d',
    IF_BG: '#ffffff',
    IF_COL: '#000000',
    setFormatAtEncryption : true,
    genPassPunctuations: ',.;:()[]{}/\\',
    genPassSymbols: '!#$%&*+-<=>?@^_|~',
    genPassAlpha: 'abcdefgijkmnopqrstwxyzABCDEFGHJKLMNPQRSTWXYZ', // 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ',
    genPassNum: '123456789' // '0123456789'
  };
  return settings;
} catch(e) { handleError(e); } } //for logging


function getPreferences()  { try  //for logging
{
  var settings = {
    passwordLength: getP_PasswordLength(),
    passwordChars: getP_PasswordChars(),
    revealMode: getP_RevealMode(),
    autoEncryptNewPassword: getP_AutoEncryptNewPassword()
  };
  return settings;
} catch(e) { handleError(e); } } //for logging

function savePreferences(preferences)  { try  //for logging
{
  setP_PasswordLength(preferences.passwordLength);
  setP_PasswordChars(preferences.passwordChars);
  setP_RevealMode(preferences.revealMode);
  setP_AutoEncryptNewPassword(preferences.autoEncryptNewPassword);
} catch(e) { handleError(e); } } //for logging

