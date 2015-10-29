function encrypt(text,password)
{
  return encryptAES(text,password);
}
function decrypt(text,password)
{
  return decryptAES(text,password); 
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
