
const XLSX = require('xlsx-extract').XLSX;
const xlsx = require('node-xlsx');

const fs = require('fs')

const path = require("path");
const writeLineStream = require('lei-stream').writeLine;
const readline = require('readline');

const beginTime = new Date().getTime();


/**
 * 读取配置
 */
//----------------------------------------------------------------------------------
const string = fs.readFileSync('./data/cfg.TXT', 'utf8');
let cfgMonth = 0;
const match = string.match(/=(.{1,2});/);
if (!match) {
    console.error("====================配置月份错误====================");
}
else {
    cfgMonth = match[1] % 12;
    if (cfgMonth < 1 || cfgMonth > 12) {
        console.error("====================配置月份数值错误====================");
    }
    else {
        console.log(`找到分析月份为：${cfgMonth}`);
    }
}
//----------------------------------------------------------------------------------
/**
 * 读取模板头
 */
//----------------------------------------------------------------------------------
const pathTemplate = `./data/template.xlsx`;
var template = xlsx.parse(pathTemplate);
console.log(`读取模板文件成功：${pathTemplate}`);
const rawData = template[0].data;
//----------------------------------------------------------------------------------
/**
 * 输出流
 */
//----------------------------------------------------------------------------------
const filePath = './output/a.xlsx';
const outputStream = writeLineStream(fs.createWriteStream(filePath), {
    // 换行符，默认\n
    newline: '\n',
    // 编码器，可以为函数或字符串（内置编码器：json，base64），默认null
    // encoding: function (data) {
    //   return JSON.stringify(data);
    // },
    // 缓存的行数，默认为0（表示不缓存），此选项主要用于优化写文件性能，写入的内容会先存储到缓存中，当内容超过指定数量时再一次性写入到流中，可以提高写速度
    cacheLines: 1000
});

let g_idx = 0;

const writeLine = (array) => {
    let string = '';
    const l = array.length;
    array.forEach((element, index) => {
        if (index != (l - 1)) {
            string += element + `\t`;
        }
        else {
            string += element;
        }
    });
    outputStream.write(string, () => {
        // 回调函数可选
    });
}
// write head
const row0 = rawData[0].toString().replace(/,/g, '\t');
outputStream.write(row0);
g_idx++;
const row1 = rawData[1].toString().replace(/,/g, '\t');
outputStream.write(row1);
g_idx++;
outputStream.flush();

//----------------------------------------------------------------------------------
/**
 * 读取3月配置表
 */
//----------------------------------------------------------------------------------
let month3RowCount = 0;
cfgMonth = 1;
const pathMonth3 = `./data/${cfgMonth}.xlsx`;
let beginMonth3 = 0;
let f3 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth3, { sheet_id: 1 }) // or sheet_name or sheet_nr
            .on('sheet', function (sheet) {
                // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
                console.log(`找到 {${cfgMonth}} 月数据成功`);
                beginMonth3 = new Date().getTime();
            })
            .on('row', function (row) {
                //搜索词
                const words = row[1];
                if (!words) {
                    return;
                }
                month3RowCount++;
                //搜索词单词数
                const wordCount = words.split(" ").length;
                //单月出现次数（取搜索词重复）
                //3个月平均排名
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                // console.log(`找到第 {${month3RowCount}} 条数据`);
                ++g_idx;
                let array = [
                    row[0],
                    words, //Search Term
                    wordCount, //Search Term count
                    1, //searchCount3
                    0, //(row.rank1 + row.rank2 + row.rank3) /3
                    0, //rank1
                    0, //rank2
                    rank, //rank3
                    "",
                    0,//row.rank1-row.rank3,
                    0, //row.searchCount1 + row.searchCount2 + row.searchCount3,
                    0,//G_K_O_1
                    0,//G_K_O_2
                    0,//G_K_O_3
                    0,//X1ClickedAsin_1,        
                    0,//X2ClickedAsin_1    
                    row[3],//X3ClickedAsin_1    
                    0,//X1Product_1, 	        
                    0,//X2Product_1        
                    row[4],//X3Product_1        
                    0,//X1ClickShare_1, 	    
                    0,//X2ClickShare_1     
                    row[5],//X3ClickShare_1     
                    0,//X1ConversionShare_1, 	
                    0,//X2ConversionShare_1
                    row[6],//X3ConversionShare_1
                    0,//X1ClickedAsin_2,        
                    0,//X2ClickedAsin_2    
                    row[7],//X3ClickedAsin_2    
                    0,//X1Product_2,            
                    0,//X2Product_2        
                    row[8],//X3Product_2        
                    0,//X1ClickShare_2,   	    
                    0,//X2ClickShare_2     
                    row[9],//X3ClickShare_2     
                    0,//X1ConversionShare_2,	
                    0,//X2ConversionShare_2
                    row[10],//X3ConversionShare_2
                    0,//X1ClickedAsin_3, 	    
                    0,//X2ClickedAsin_3    
                    row[11],//X3ClickedAsin_3    
                    0,//X1Product_3,            
                    0,//X2Product_3        
                    row[12],//X3Product_3        
                    0,//X1ClickShare_3,   	    
                    0,//X2ClickShare_3     
                    row[13],//X3ClickShare_3     
                    0,//X1ConversionShare_3, 	
                    0,//X2ConversionShare_3
                    row[14],//X3ConversionShare_3
                ];
                writeLine(array);

            })
            .on('error', function (err) {
                console.error('error', err);
            })
            .on('end', function (err) {
                // console.log('eof 3');
                // outputStream.flush();
                console.log(`读取${cfgMonth}月表时间为：${(new Date().getTime() - beginMonth3) / 1000}`)
                outputStream.end();

                resolve(1);
            });
    });
}
/**
 * 读取3月配置表
 */
//----------------------------------------------------------------------------------
const f4 = function () {
    new Promise(() => {
        console.log(`month1RowCount: ${month3RowCount}`);

        for (var item of result.entries()) {
            idx++;
            const row = item[1];
            let array = [
                idx,
                item[0],
                row.wordCount,
                row.searchCount3,
                (row.rank1 + row.rank2 + row.rank3) / 3,
                row.rank1,
                row.rank2,
                row.rank3,
                "",
                row.rank1 - row.rank3,
                row.searchCount1 + row.searchCount2 + row.searchCount3,
                row.X1ConversionShare,
                row.X2ConversionShare,
                row.X3ConversionShare,
                row.X1ClickedAsin_1, row.X2ClickedAsin_1, row.X3ClickedAsin_1,
                row.X1Product_1, row.X2Product_1, row.X3Product_1,
                row.X1ClickShare_1, row.X2ClickShare_1, row.X3ClickShare_1,
                row.X1ConversionShare_1, row.X2ConversionShare_1, row.X3ConversionShare_1,
                row.X1ClickedAsin_2, row.X2ClickedAsin_2, row.X3ClickedAsin_2,
                row.X1Product_2, row.X2Product_2, row.X3Product_2,
                row.X1ClickShare_2, row.X2ClickShare_2, row.X3ClickShare_2,
                row.X1ConversionShare_2, row.X2ConversionShare_2, row.X3ConversionShare_2,
                row.X1ClickedAsin_3, row.X2ClickedAsin_3, row.X3ClickedAsin_3,
                row.X1Product_3, row.X2Product_3, row.X3Product_3,
                row.X1ClickShare_3, row.X2ClickShare_3, row.X3ClickShare_3,
                row.X1ConversionShare_3, row.X2ConversionShare_3, row.X3ConversionShare_3,
            ];

        }
        // 结束
        outputStream.end(() => {
            // 回调函数可选
            console.log(`统计完成，文件路径为: ${filePath}`);
            const deltaAppTime = (new Date().getTime() - beginTime) / 1000;
            console.log(`delta app time: ${deltaAppTime}s`);
        });
        // 写时出错
        outputStream.on('error', (err) => {
            console.error(err);
        });
    });
}
const finish = () => {

}
const plist = [f3, finish];
// 递归调用
// 使用 await & async
async function promise_queue(list) {
    let index = 0
    while (index >= 0 && index < list.length) {
        await list[index]();
        index++
    }
}
promise_queue(plist);

setInterval(() => {
    console.log(`month3RowCount: ${month3RowCount}`);
}, 1000);
