const CryptoJS = require('crypto-js');
const clipboardy = require('clipboardy');

let decryptAES = function (message, key) {
    return CryptoJS.AES.decrypt(message, CryptoJS.enc.Utf8.parse(key), {
        mode: CryptoJS.mode.ECB,
        padding: CryptoJS.pad.Pkcs7
    }).toString(CryptoJS.enc.Utf8);
}

const string = clipboardy.readSync();
if (!string) {
    console.error(`请输入解密字符串`);
    return;
}
console.log(`str: ${string}`);
const key = "12345678876543211234567887654abc";
console.log(`key: ${key}`);
let result = decryptAES(string, key);
if (!result) {
    console.log(`decrypt failed: ${result}`);
}
else {
    console.log(`decrypt success: ${result}`);
}
