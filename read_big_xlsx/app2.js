
String.prototype.toBytes = function(encoding){
    var bytes = [];
    var buff = Buffer.alloc(100);
    for(var i= 0; i< buff.length; i++){
      var byteint = buff[i];
      bytes.push(byteint);
    }
    return bytes;
}
   
const XLSX = require('xlsx-extract').XLSX;
const xlsx = require('node-xlsx');

const fs = require('fs')

const path = require("path");
const writeLineStream = require('lei-stream').writeLine;

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
const beginTime = new Date().getTime();

var result = new Map();


// 读取1
let month1RowCount = 0;
let month1 = (cfgMonth + 12 - 2) % 12;
if (month1 == 0) {
    month1 = 12;
}
const pathMonth1 = `./data/${month1}.xlsx`;
const f1 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth1, { sheet_id: 1 }) // or sheet_name or sheet_nr
        .on('sheet', function (sheet) {
            // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
            console.log(`找到 {${month1}} 月数据成功`);
        })
        .on('row', function (row) {
            // console.log('row', row);  //row is a array of values or []
            month1RowCount++;
            const words = row[1];
            if (!words) {
                return;
            }
            //搜索词单词数
            const wordCount = words.split(" ").length;
            //单月出现次数（取搜索词重复）
            //3个月平均排名
            const rank = row[2];
            const r = Number(rank);
            if (isNaN(r)) {
                return;
            }
            const rs = result.get(words);
            if (rs) {
                rs.searchCount1++;
                rs.X1ConversionShare = (row[6] + row[10] + row[14]) * 100;
                rs.rank1                = rank ;
                rs.X1ClickedAsin_1      = row[3];   
                rs.X1Product_1          = row[4];     
                rs.X1ClickShare_1       = row[5];  
                rs.X1ConversionShare_1  = row[6];  
                rs.X1ClickedAsin_2      = row[7]; 
                rs.X1Product_2          = row[8];     
                rs.X1ClickShare_2       = row[9];  
                rs.X1ConversionShare_2  = row[10];  
                rs.X1ClickedAsin_3      = row[11]; 
                rs.X1Product_3          = row[12];     
                rs.X1ClickShare_3       = row[13];  
                rs.X1ConversionShare_3  = row[14]; 
            }
            else {
                //1. 2. 3月份词频排名
                result.set(words, { wordCount: wordCount,
                    X3ConversionShare:              0,
                    X2ConversionShare:              0,
                    X1ConversionShare:              (row[6] + row[10] + row[14]) * 100,
                    rank1:  rank,                       rank2: 0,                       rank3: 0, 
                    searchCount1: 1,                    searchCount2: 0,                searchCount3: 0,
                    X1ClickedAsin_1:row[3],             X2ClickedAsin_1:"",             X3ClickedAsin_1:"",
                    X1Product_1:row[4], 	            X2Product_1:"", 	            X3Product_1:"", 
                    X1ClickShare_1:row[5], 	            X2ClickShare_1:0, 		        X3ClickShare_1:0, 	
                    X1ConversionShare_1:row[6],         X2ConversionShare_1:0, 	        X3ConversionShare_1:0, 	
                    X1ClickedAsin_2:row[7],             X2ClickedAsin_2:"",   	        X3ClickedAsin_2:"",     
                    X1Product_2:row[8],     	        X2Product_2:"",     	        X3Product_2:"",     
                    X1ClickShare_2:row[9],   	        X2ClickShare_2:0,   	        X3ClickShare_2:0,
                    X1ConversionShare_2:row[10],        X2ConversionShare_2:0,	        X3ConversionShare_2:0,
                    X1ClickedAsin_3:row[11], 	        X2ClickedAsin_3:"", 	        X3ClickedAsin_3:"",
                    X1Product_3:row[12],     	        X2Product_3:"",     	        X3Product_3:"", 
                    X1ClickShare_3:row[13],   	        X2ClickShare_3:0,   	        X3ClickShare_3:0, 
                    X1ConversionShare_3:row[14],        X2ConversionShare_3:0, 	        X3ConversionShare_3:0
                 });        
            }
    
        })
        .on('error', function (err) {
            console.error('error', err);
        })
        .on('end', function (err) {
            // console.log('eof 1');
            resolve(1);
        });
    });
}

// 读取2月
let month2 = (cfgMonth + 12 - 1) % 12;
if (month2 == 0) {
    month2 = 12;
}
const pathMonth2 = `./data/${month2}.xlsx`;
const f2 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth2, { sheet_id: 1 }) // or sheet_name or sheet_nr
            .on('sheet', function (sheet) {
                // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
                console.log(`找到 {${month2}} 月数据成功`);
            })
            .on('row', function (row) {
                // console.log('row', row);  //row is a array of values or []
                //搜索词
                const words = row[1];
                if (!words) {
                    return;
                }
                //搜索词单词数
                const wordCount = words.split(" ").length;
                //单月出现次数（取搜索词重复）
                //3个月平均排名
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                const rs = result.get(words);
                if (rs) {
                    rs.searchCount2++;
                    rs.X2ConversionShare = (row[6] + row[10] + row[14]) * 100;
                    rs.rank2                = rank ;
                    rs.X2ClickedAsin_1      = row[3];   
                    rs.X2Product_1          = row[4];     
                    rs.X2ClickShare_1       = row[5];  
                    rs.X2ConversionShare_1  = row[6];  
                    rs.X2ClickedAsin_2      = row[7]; 
                    rs.X2Product_2          = row[8];     
                    rs.X2ClickShare_2       = row[9];  
                    rs.X2ConversionShare_2  = row[10];  
                    rs.X2ClickedAsin_3      = row[11]; 
                    rs.X2Product_3          = row[12];     
                    rs.X2ClickShare_3       = row[13];  
                    rs.X2ConversionShare_3  = row[14];  
                }
                else {
                    //1. 2. 3月份词频排名
                    result.set(words, { 
                        wordCount: wordCount,
                        X3ConversionShare:              0,
                        X2ConversionShare:              (row[6] + row[10] + row[14]) * 100,
                        X1ConversionShare:              0,
                        rank1:  0,                              rank2: rank,                rank3: 0, 
                        searchCount1: 0,                        searchCount2: 1,            searchCount3: 0,
                        X1ClickedAsin_1:"",                     X2ClickedAsin_1:row[3],     X3ClickedAsin_1:"",
                        X1Product_1:            "", 	        X2Product_1:row[4], 	    X3Product_1:"", 
                        X1ClickShare_1:         "", 	        X2ClickShare_1:row[5], 		X3ClickShare_1:0, 	
                        X1ConversionShare_1:    "", 	        X2ConversionShare_1:row[6], 	    X3ConversionShare_1:0, 	
                        X1ClickedAsin_2:        "",             X2ClickedAsin_2:row[7],   	    X3ClickedAsin_2:"",     
                        X1Product_2:            "",             X2Product_2:row[8],         X3Product_2:"",     
                        X1ClickShare_2:         "",   	        X2ClickShare_2:row[9],   	    X3ClickShare_2:0,
                        X1ConversionShare_2:    "",	            X2ConversionShare_2:row[10],	    X3ConversionShare_2:0,
                        X1ClickedAsin_3:        "", 	        X2ClickedAsin_3:row[11], 	    X3ClickedAsin_3:"",
                        X1Product_3:            "",             X2Product_3:row[12],         X3Product_3:"", 
                        X1ClickShare_3:         "",   	        X2ClickShare_3:row[13],   	    X3ClickShare_3:0, 
                        X1ConversionShare_3:    "", 	        X2ConversionShare_3:row[14], 	    X3ConversionShare_3:0
                    });
                }
                })
                .on('error', function (err) {
                    console.error('error', err);
                })
                .on('end', function (err) {
                    // console.log('eof 2');
                    resolve(1);
                });
    });
};


// 读取3月
const pathMonth3 = `./data/${cfgMonth}.xlsx`;

let f3 = function() {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth3, { sheet_id: 1 }) // or sheet_name or sheet_nr
        .on('sheet', function (sheet) {
            // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
            console.log(`找到 {${cfgMonth}} 月数据成功`);
        })
        .on('row', function (row) {
            //搜索词
            const words = row[1];
            if (!words) {
                return;
            }
            //搜索词单词数
            const wordCount = words.split(" ").length;
            //单月出现次数（取搜索词重复）
            //3个月平均排名
            const rank = row[2];
            const r = Number(rank);
            if (isNaN(r)) {
                return;
            }
            const rs = result.get(words);
            if (rs) {
                rs.searchCount3++;
            }
            else {
                //1. 2. 3月份词频排名
                result.set(words, { 
                    wordCount:              wordCount,
                    X3ConversionShare:              (row[6] + row[10] + row[14]) * 100,
                    X2ConversionShare:              0,
                    X1ConversionShare:              0,
                    rank1:                  0,              rank2:                      0,                  rank3:                      rank, 
                    searchCount1:           0,              searchCount2:               0,                  searchCount3:               1,
                    X1ClickedAsin_1:        "",             X2ClickedAsin_1:            "",             X3ClickedAsin_1:            row[3],
                    X1Product_1:            "", 	        X2Product_1:                "", 	        X3Product_1:                row[4],
                    X1ClickShare_1:         "", 	        X2ClickShare_1:             "", 		    X3ClickShare_1:             row[5],
                    X1ConversionShare_1:    "", 	        X2ConversionShare_1:        "", 	        X3ConversionShare_1:        row[6],
                    X1ClickedAsin_2:        "",             X2ClickedAsin_2:            "",   	        X3ClickedAsin_2:            row[7],   
                    X1Product_2:            "",             X2Product_2:                "",             X3Product_2:                row[8],    
                    X1ClickShare_2:         "",   	        X2ClickShare_2:             "",   	        X3ClickShare_2:             row[9],
                    X1ConversionShare_2:    "",	            X2ConversionShare_2:        "",	            X3ConversionShare_2:        row[10],
                    X1ClickedAsin_3:        "", 	        X2ClickedAsin_3:            "", 	        X3ClickedAsin_3:            row[11],
                    X1Product_3:            "",             X2Product_3:                "",             X3Product_3:                row[12],
                    X1ClickShare_3:         "",   	        X2ClickShare_3:             "",   	        X3ClickShare_3:             row[13],
                    X1ConversionShare_3:    "", 	        X2ConversionShare_3:        "", 	        X3ConversionShare_3:        row[14],
                 });
            }
            
        })
        .on('error', function (err) {
            console.error('error', err);
        })
        .on('end', function (err) {
            // console.log('eof 3');
            resolve(1);
        });
    });
}
const f4 = function () {
    new Promise(() => {
        console.log(`month1RowCount: ${month1RowCount}`);

        const pathTemplate = `./data/template.xlsx`;
        var template = xlsx.parse(pathTemplate);
        console.log(`读取模板文件成功：${pathTemplate}`);
        const rawData = template[0].data;
        let idx = 0;
        const datas = [rawData[0],
        rawData[1]];
        const filePath = './output/a.xlsx';
        const s = writeLineStream(fs.createWriteStream(filePath), {
            // 换行符，默认\n
            newline: '\n',
            // 编码器，可以为函数或字符串（内置编码器：json，base64），默认null
            // encoding: function (data) {
            //   return JSON.stringify(data);
            // },
            // 缓存的行数，默认为0（表示不缓存），此选项主要用于优化写文件性能，写入的内容会先存储到缓存中，当内容超过指定数量时再一次性写入到流中，可以提高写速度
            cacheLines: 0
          });
        const row0 = rawData[0].toString().replace(/,/g, '\t');
        s.write(row0, () => {
            // 回调函数可选
        });
        const row1 = rawData[1].toString().replace(/,/g, '\t');
        s.write(row1, () => {
            // 回调函数可选
        });
        
        
        for (var item of result.entries()) {
            idx++;
            const row = item[1];
            let array = [
                idx,
                item[0],
                row.wordCount,
                row.searchCount3,
                (row.rank1 + row.rank2 + row.rank3) /3,
                row.rank1,
                row.rank2,
                row.rank3,
                "",
                row.rank1-row.rank3,
                row.searchCount1 + row.searchCount2 + row.searchCount3,
                row.X1ConversionShare,
                row.X2ConversionShare,
                row.X3ConversionShare,
                row.X1ClickedAsin_1,                row.X2ClickedAsin_1         ,           row.X3ClickedAsin_1     ,
                row.X1Product_1, 	                row.X2Product_1             , 	        row.X3Product_1         ,
                row.X1ClickShare_1, 	            row.X2ClickShare_1          , 		    row.X3ClickShare_1      ,
                row.X1ConversionShare_1, 	        row.X2ConversionShare_1     , 	        row.X3ConversionShare_1 ,
                row.X1ClickedAsin_2,                row.X2ClickedAsin_2         ,           row.X3ClickedAsin_2     ,
                row.X1Product_2,                    row.X2Product_2             ,           row.X3Product_2         ,
                row.X1ClickShare_2,   	            row.X2ClickShare_2          ,	        row.X3ClickShare_2      ,
                row.X1ConversionShare_2,	        row.X2ConversionShare_2     ,           row.X3ConversionShare_2 ,
                row.X1ClickedAsin_3, 	            row.X2ClickedAsin_3         ,           row.X3ClickedAsin_3     ,
                row.X1Product_3,                    row.X2Product_3             ,           row.X3Product_3         ,
                row.X1ClickShare_3,   	            row.X2ClickShare_3          ,	        row.X3ClickShare_3      ,
                row.X1ConversionShare_3, 	        row.X2ConversionShare_3     ,           row.X3ConversionShare_3 ,
            ];
            let string = '';
            const l = array.length;
            array.forEach((element ,index) => {
                if (index != (l-1)) {
                    string += element + `\t`;
                }
                else {
                    string += element;
                }
            });
            s.write(string, () => {
                // 回调函数可选
            });
        }
        // 结束
        s.end(() => {
            // 回调函数可选
            console.log(`统计完成，文件路径为: ${filePath}`);
            const deltaAppTime = (new Date().getTime() - beginTime) / 1000;
            console.log(`delta app time: ${deltaAppTime}s`);
        });
        // 写时出错
        s.on('error', (err) => {
            console.error(err);
        });
    });
}

const plist = [f3, f2, f1, f4];
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
    console.log(`month1RowCount: ${month1RowCount}`);
}, 1000);




