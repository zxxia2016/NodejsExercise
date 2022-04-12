const XLSX = require('xlsx-extract').XLSX;
const fs = require('fs')

let tbNme = `tb_${new Date().getTime()}`;
// tbNme = `tb_1649405786104`;

const OPEN_TIMER = false;
const TIMER_DELTA = 5000;

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
        console.log(`分析月份为：${cfgMonth}`);
    }
}
//----------------------------------------------------------------------------------
/**
 * 读取3月配置表
 */
//----------------------------------------------------------------------------------

let month3RowCount = 0;
const pathMonth3 = `./data/${cfgMonth}.xlsx`;
let beginMonth3 = 0;
let timer3 = null;
let f3 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth3, { sheet_id: 1 }) // or sheet_name or sheet_nr
            .on('sheet', function (sheet) {
                // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
                console.log(`找到 {${cfgMonth}} 月数据成功`);
                beginMonth3 = new Date().getTime();
                if (OPEN_TIMER) {
                    timer3 = setInterval(() => {
                        console.log(`month3RowCount: ${month3RowCount}`);
                    }, TIMER_DELTA);
                }
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
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                month3RowCount++;
                // console.log(`找到第 {${month3RowCount}} 条数据`);

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
                console.log(`读取${cfgMonth}月表时间为：${(new Date().getTime() - beginMonth3) / 1000 / 60} 分钟`);
                console.log(`读取总数为: ${month3RowCount}`);
                resolve();
            });
    });
}


let month2RowCount = 0;

// 读取2月
let month2 = (cfgMonth + 12 - 1) % 12;
if (month2 == 0) {
    month2 = 12;
}
const pathMonth2 = `./data/${month2}.xlsx`;
let beginMonth2 = 0;
let timer2 = null;
const f2 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth2, { sheet_id: 1 }) // or sheet_name or sheet_nr
            .on('sheet', function (sheet) {
                // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
                console.log(`找到 {${month2}} 月数据成功`);
                beginMonth2 = new Date().getTime();

                if (OPEN_TIMER) {
                    timer2 = setInterval(() => {
                        console.log(`month2RowCount: ${month2RowCount}`);
                    }, TIMER_DELTA);
                }
            })
            .on('row', function (row) {
                // console.log('row', row);  //row is a array of values or []
                //搜索词
                const words = row[1];
                if (!words) {
                    return;
                }
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                month2RowCount++;

                const sql = `select * from ${tbNme} where SearchTerm = ?`;
                sqliteDB.getDataPromise(sql, [words]).then(function (rowData) {
                    if (rowData) {
                        // 更新数量
                        const sql = `update ${tbNme} set 
                        SearchCount2 = ?, 
                        Rank2 = ?,
                        X2ClickedAsin_1 = ?,
                        X2Product_1 = ?,
                        X2ClickShare_1 = ?,
                        X2ConversionShare_1 = ?,
                        X2ClickedAsin_2 = ?, 
                        X2Product_2 = ?,
                        X2ClickShare_2 = ?,
                        X2ConversionShare_2 = ?,
                        X2ClickedAsin_3 = ?,
                        X2Product_3 = ?,
                        X2ClickShare_3 = ?,
                        X2ConversionShare_3 = ? 
                        where SearchTerm = ?`;
                        const array = [++rowData.SearchCount2, rank, row[3], row[4], row[5],
                        row[6], row[7], row[8], row[9], row[10], row[11], row[12],
                        row[13], row[14], words];
                        sqliteDB.runSql(sql, array);
                    }
                })

            })
            .on('error', function (err) {
                console.error('error', err);
            })
            .on('end', function (err) {
                console.log(`读取${month2}月表时间为：${(new Date().getTime() - beginMonth2) / 1000 / 60} 分钟`);
                console.log(`读取总数为: ${month2RowCount}`);
                resolve();
            });
    });
};

// 读取1
let month1RowCount = 0;
let month1 = (cfgMonth + 12 - 2) % 12;
if (month1 == 0) {
    month1 = 12;
}
const pathMonth1 = `./data/${month1}.xlsx`;
let beginMonth1 = 0;
let timer1 = null;
const f1 = function () {
    return new Promise((resolve, reject) => {
        new XLSX().extract(pathMonth1, { sheet_id: 1 }) // or sheet_name or sheet_nr
            .on('sheet', function (sheet) {
                // console.log('sheet', sheet);  //sheet is array [sheetname, sheetid, sheetnr]
                console.log(`找到 {${month1}} 月数据成功`);
                beginMonth1 = new Date().getTime();

                if (OPEN_TIMER) {
                    timer1 = setInterval(() => {
                        console.log(`month1RowCount: ${month1RowCount}`);
                    }, TIMER_DELTA);
                }
            })
            .on('row', function (row) {
                // console.log('row', row);  //row is a array of values or []
                month1RowCount++;
                const words = row[1];
                if (!words) {
                    return;
                }
                const rank = row[2];
                const r = Number(rank);
                if (isNaN(r)) {
                    return;
                }
                month1RowCount++;

                const sql = `select * from ${tbNme} where SearchTerm = ?`;
                sqliteDB.getDataPromise(sql, [words]).then(function (rowData) {
                    if (rowData) {
                        // 更新数量
                        const sql = `update ${tbNme} set 
                        SearchCount1 = ?, 
                        Rank1 = ?,
                        X1ClickedAsin_1 = ?, 
                        X1Product_1 = ?, 
                        X1ClickShare_1 = ?,
                        X1ConversionShare_1 = ?, 
                        X1ClickedAsin_2 = ?, 
                        X1Product_2 = ?,
                        X1ClickShare_2 = ?,
                        X1ConversionShare_2 = ?,
                        X1ClickedAsin_3 = ?,
                        X1Product_3 = ?,
                        X1ClickShare_3 = ?,
                        X1ConversionShare_3 = ? 
                        where SearchTerm = ?`;
                        const array = [++rowData.SearchCount1, rank, row[3], row[4], row[5],
                        row[6], row[7], row[8], row[9], row[10], row[11], row[12],
                        row[13], row[14], words];
                        sqliteDB.runSql(sql, array);
                    }
                })

            })
            .on('error', function (err) {
                console.error('error', err);
            })
            .on('end', function (err) {
                console.log(`读取${month1}月表时间为：${(new Date().getTime() - beginMonth1) / 1000 / 60} 分钟`);
                console.log(`读取总数为: ${month1RowCount}`);
                resolve();
            });
    });
}



if (cfgMonth) {
    const file = ":memory:";
    // const file = `./tmp/data.db`;
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
        .then(f3).then(function () {
            return new Promise(function (resolve, reject) {
                clearInterval(timer3);
                resolve();
            })
        }).then(f2).then(function () {
            return new Promise(function (resolve, reject) {
                clearInterval(timer2);
                resolve();
            })
        }).then(f1).then(function () {
            return new Promise(function (resolve, reject) {
                clearInterval(timer1);
                console.log(`总共花费时间为：${(new Date().getTime() - beginMonth3) / 1000 / 60} 分钟`);
                resolve();
            })
        }).catch(function (err) {
            console.error(err);
        })
}

