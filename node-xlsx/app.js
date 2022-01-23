var xlsx = require('node-xlsx');
const fs = require('fs');

const keys = [{ sheetName: "firstSesstion", keyName: '"isFirstSession":true,' },
{ sheetName: "clickedTenjinLink", keyName: '"clickedTenjinLink":true,' },];

const parseProperty = function (sheet, keyName) {
    const arraySession = [];
    const title = ['name']//这是第一行 俗称列名 
    arraySession.push(title);

    // 读取每行内容
    for (var rowId in sheet['data']) {
        var row = sheet['data'][rowId];
        const content = row[1];
        const ret = content.search(keyName);
        if (ret > -1) {
            const data = [];
            data.push(content);
            arraySession.push(data);
        }
    }
    return arraySession;
}

// console.log(process.execPath)
// console.log(__dirname)
// console.log(process.cwd())

// 解析得到文档中的所有 sheet
var sheets = xlsx.parse(__dirname + '/gm_ten_jin.xlsx');
// 遍历 sheet
sheets.forEach(function (sheet) {
    console.log(sheet['name']);

    const buildInfo = [];
    keys.forEach(element => {
        const datas = parseProperty(sheet, element.keyName);
        console.log(`name: ${element.sheetName}, len: ${datas.length}`);
        buildInfo.push({
            name: element.sheetName,
            data: datas
        });
    });
    let buffer = xlsx.build(buildInfo);
    fs.writeFileSync(__dirname + '/gm_ten_jin_out.xlsx', buffer, { 'flag': 'w' });
});

console.log('finished');