var xlsx = require('node-xlsx');

const fs = require('fs')

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
const debug = false;

const beginTime = new Date().getTime();

// 读取1
let month1 = (cfgMonth + 12 - 2) % 12;
if (month1 == 0) {
    month1 = 12;
}
const pathMonth1 = `./data/${month1}.xlsx`;
const sheets1 = xlsx.parse(pathMonth1);
console.log(`找到 {${month1}} 月数据成功`);

// 读取2月
let month2 = (cfgMonth + 12 - 1) % 12;
if (month2 == 0) {
    month2 = 12;
}
const pathMonth2 = `./data/${month2}.xlsx`;
const sheets2 = xlsx.parse(pathMonth2);
console.log(`找到 {${month2}} 月数据成功`);

// 读取3月
const pathMonth3 = `./data/${cfgMonth}.xlsx`;
var sheets3 = xlsx.parse(pathMonth3);
console.log(`找到 {${cfgMonth}} 月数据成功`);

const deltaTime = (new Date().getTime() - beginTime) / 1000;
console.log(`delta time: ${deltaTime}s`);

// 解析得到文档中的所有 sheet
var sheets = xlsx.parse('./data/test.xlsx');

var result = new Map();

// 遍历 sheet
sheets1.forEach(function (sheet) {
    //表头
    // console.log(sheet['name']);
    // 读取每行内容
    const rawData = sheet['data'];
    for (let rowId = 0; rowId < rawData.length; rowId++) {
        const row = rawData[rowId];
        //去掉表头
        if (rowId < 2) {
            continue;
        }
        console.log(row);
        //搜索词
        const words = row[1];
        //搜索词单词数
        const wordCount = words.split(" ").length;
        //单月出现次数（取搜索词重复）
        //3个月平均排名
        const rank = row[2];
        const rs = result.get(words);
        if (rs) {
            rs.searchCount1++;
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
        //匹配数（3个月搜索词重复）
        //GKO1

    }

});

sheets2.forEach(function (sheet) {
    //表头
    // console.log(sheet['name']);
    // 读取每行内容
    const rawData = sheet['data'];
    for (let rowId = 0; rowId < rawData.length; rowId++) {
        const row = rawData[rowId];
        //去掉表头
        if (rowId < 2) {
            continue;
        }
        console.log(row);
        //搜索词
        const words = row[1];
        //搜索词单词数
        const wordCount = words.split(" ").length;
        //单月出现次数（取搜索词重复）
        //3个月平均排名
        const rank = row[2];
        const rs = result.get(words);
        if (rs) {
            rs.searchCount2++;
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
        //匹配数（3个月搜索词重复）
        //GKO1

    }

});

sheets3.forEach(function (sheet) {
    //表头
    // console.log(sheet['name']);
    // 读取每行内容
    const rawData = sheet['data'];
    for (let rowId = 0; rowId < rawData.length; rowId++) {
        const row = rawData[rowId];
        //去掉表头
        if (rowId < 2) {
            continue;
        }
        console.log(row);
        //搜索词
        const words = row[1];
        //搜索词单词数
        const wordCount = words.split(" ").length;
        //单月出现次数（取搜索词重复）
        //3个月平均排名
        const rank = row[2];
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
        //匹配数（3个月搜索词重复）
        //GKO1

    }

});



const pathTemplate = `./data/template.xlsx`;
var template = xlsx.parse(pathTemplate);
console.log(`读取模板文件成功：${pathTemplate}`);
const rawData = template[0].data;
let idx = 0;
const datas = [rawData[0],
rawData[1]];
for (var item of result.entries()) {
    idx++;
    const row = item[1];
    datas.push([
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
    ]);
  }
//排名最终变化
var data = [{
    name: 'sheet1',
    data: datas
}
]
var buffer = xlsx.build(data);

// 写入文件
const filePath = './output/a.xlsx';
fs.writeFileSync(filePath, buffer);
console.log(`统计完成，文件路径为: ${filePath}`);
const deltaAppTime = (new Date().getTime() - beginTime) / 1000;
console.log(`delta app time: ${deltaAppTime}s`);
