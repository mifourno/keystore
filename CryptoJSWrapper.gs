/* ----------------------------------------------------------
CryptoJSWrapper.gs -- 
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


function encrypt(text,password)
{
  return encryptAES(text,password);
}
function decrypt(text,password)
{
  try {
    return decryptAES(text,password); 
  } catch (e) {
    return '';
  }
}

//#########################################################################

function encryptAES(text,password)
{
  return CryptoJS.AES.encrypt(text, password).toString();
}
function decryptAES(text,password)
{
  return CryptoJS.AES.decrypt(text, password).toString(CryptoJS.enc.Utf8);
}
