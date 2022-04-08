const XLSX = require('xlsx-extract').XLSX;
const fs = require('fs')

let tbNme = `tb_${new Date().getTime()}`;
// tbNme = `tb_1649405786104`;

var SqliteDB = require('./sqlite.js').SqliteDB;
var sqliteDB = new SqliteDB();

// 递归调用
// 使用 await & async
async function promise_queue(list) {
    let index = 0
    while (index >= 0 && index < list.length) {
        await list[index]();
        index++
    }
}

function promise_queue1(arry) {
    var sequence = Promise.resolve()
    arry.forEach(function (item) {
        sequence = sequence.then(item)
    })
    return sequence;
}

//----------------------------------------------------------------------------------
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
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                // console.log(`找到第 {${month3RowCount}} 条数据`);

                //todo先查询，再插入
                const sql = `select * from ${tbNme} where SearchTerm = ?`;
                sqliteDB.getDataPromise(sql, [words]).then(function (rowData) {
                    if (rowData) {
                        // 更新数量
                        const sql = `update ${tbNme} set SearchCount3 = ? where SearchTerm = ?`;
                        const array = [++rowData.SearchCount3, words];
                        sqliteDB.runSql(sql, array);
                    }
                    else {
                        const strSql = `insert into ${tbNme} values(?, ?, ?, ?, ?, 
                            ?, ?, ?, 
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?,
                            ?, ?, ?
                            )`;
                        const array = [words, wordCount,
                            0, 0, 1,
                            0, 0, rank,
                            0, 0, 0,
                            "", "", row[3],
                            "", "", row[4],
                            "", "", row[5],
                            "", "", row[6],
                            "", "", row[7],
                            "", "", row[8],
                            "", "", row[9],
                            "", "", row[10],
                            "", "", row[11],
                            "", "", row[12],
                            "", "", row[13],
                            "", "", row[14]
                        ];
                        sqliteDB.runSql(strSql, array);
                    }
                })

            })
            .on('error', function (err) {
                console.error('error', err);
                reject(err);
            })
            .on('end', function (err) {
                console.log(`读取${cfgMonth}月表时间为：${(new Date().getTime() - beginMonth3) / 1000 / 3600} 小时`)

                resolve();
            });
    });
}
const timer = setInterval(() => {
    console.log(`month3RowCount: ${month3RowCount}`);
}, 3000);

if (cfgMonth) {
    const file = `./tmp/data.db`;
    sqliteDB.createDataBase(file)
        .then(function () {
            const createTableSql = `create table if not exists ${tbNme}(
            SearchTerm varchar(100) NOT NULL primary key,
            SearchTermCount INT,
            SearchCount1 INT,
            SearchCount2 INT,
            SearchCount3 INT,
            Rank1 INT,
            Rank2 INT,
            Rank3 INT,
            G_K_O_1 INT,
            G_K_O_2 INT,
            G_K_O_3 INT,
            X1ClickedAsin_1 varchar(100),
            X2ClickedAsin_1 varchar(100),
            X3ClickedAsin_1 varchar(100),
            X1Product_1 varchar(100), 	     
            X2Product_1 varchar(100),       
            X3Product_1 varchar(100),        
            X1ClickShare_1 float,	    
            X2ClickShare_1 float,
            X3ClickShare_1 float,
            X1ConversionShare_1 float,
            X2ConversionShare_1 float,
            X3ConversionShare_1 float,
            X1ClickedAsin_2 varchar(100),
            X2ClickedAsin_2 varchar(100),    
            X3ClickedAsin_2 varchar(100), 
            X1Product_2 varchar(100),       
            X2Product_2 varchar(100),        
            X3Product_2 varchar(100), 
            X1ClickShare_2 float,
            X2ClickShare_2 float,
            X3ClickShare_2 float,  
            X1ConversionShare_2 float,
            X2ConversionShare_2 float,
            X3ConversionShare_2 float,
            X1ClickedAsin_3 varchar(100),
            X2ClickedAsin_3 varchar(100),
            X3ClickedAsin_3 varchar(100),
            X1Product_3 varchar(100),    
            X2Product_3 varchar(100),
            X3Product_3 varchar(100),
            X1ClickShare_3 float,
            X2ClickShare_3 float,
            X3ClickShare_3 float,
            X1ConversionShare_3 float,
            X2ConversionShare_3 float,
            X3ConversionShare_3 float
            )`;
            return sqliteDB.createTable(createTableSql);
        })
        .then(f3).then(function (data) {
            clearInterval(timer);
        }).catch(function (err) {
            console.error(err);
        })
}

