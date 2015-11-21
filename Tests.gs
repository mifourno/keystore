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

function test_CheckConsistancy() {
//  var d1 = new Date();
  checkConsistancyAllSheets();
//  var d2 = new Date();
}


function test_RemoveAllEncryption(sheet) {
  if (isNullOrWS(sheet)) sheet = SpreadsheetApp.getActiveSheet();
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var j = 0; j < protections.length; j++) {
    if (protections[j].getDescription() == getP_ProtectionMessage()) {
      protections[j].remove();
    }
  }
  
}

function temp() {
  setP_EncryptionList('');
}
