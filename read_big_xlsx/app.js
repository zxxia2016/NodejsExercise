const XLSX = require('xlsx-extract').XLSX;
const fs = require('fs')

var sqlite3 = require('sqlite3');
const tbNme = `tb_${new Date().getTime()}`;
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


let g_db = null;
let f1 = function () {
    return new Promise((resolve, reject) => {
        var db = new sqlite3.Database(`./tmp/data.db`, function (err) {
            if (err) {
                console.error(`create db error:`, err);
                reject(err);
                return;
            }
            console.log(`create data db success`);
            g_db = db;
            resolve(db);
        });
    })
};

const f2 = function (db) {
    return new Promise((resolve, reject) => {
        db.run(`create table if not exists ${tbNme}(
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
            )`, function (err) {
            if (err) {
                console.error(`create table error:`, err);
                reject()
                return;
            }
            console.log(`create table ${tbNme} success`);
            resolve()
        });
    })
};
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
cfgMonth = 2;
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
                //todo先查询，再插入
                const strSql = `insert into ${tbNme} values('${words}', ${wordCount}, 0, 0, 1, 
                    0, 0, ${rank}, 
                    0, 0, 0,
                    NULL, NULL, "${row[3]}",
                    NULL, NULL, "${row[4]}",
                    NULL, NULL, ${row[5]},
                    NULL, NULL, ${row[6]},
                    NULL, NULL, "${row[7]}",
                    NULL, NULL, "${row[8]}",
                    NULL, NULL, ${row[9]},
                    NULL, NULL, ${row[10]},
                    NULL, NULL, "${row[11]}",
                    NULL, NULL, "${row[12]}",
                    NULL, NULL, ${row[13]},
                    NULL, NULL, ${row[14]}
                    )`;
                g_db.run(strSql, function (result) {
                    if (result) {
                        console.error(`insert into ${tbNme} error:`, strSql, result);
                        return;
                    }

                })

            })
            .on('error', function (err) {
                console.error('error', err);
            })
            .on('end', function (err) {
                // console.log('eof 3');
                console.log(`读取${cfgMonth}月表时间为：${(new Date().getTime() - beginMonth3) / 1000}`)
                g_db.run(`select count(*) from ${tbNme}`, function (err, res) {
                    if (!err)
                        console.log(JSON.stringify(res));
                    else
                        console.log(err);
                });
                resolve(1);
            });
    });
}
const plist = [f1, f2, f3];

promise_queue1(plist);

// var db = new sqlite3.Database(`./tmp/data.db`, function (err) {
//     if (err) {
//         console.error(`create db error:`, err);
//         return;
//     }
//     console.log(`success`);
//     db.run(`create table ${time} (name varchar(15))`, function (result) {
//         if (result) {
//             console.error(`create table error:`, result);
//             return;
//         }

//         db.run("insert into test values('hello,world')", function (result) {
//             if (result) {
//                 console.error(`insert into error:`, result);
//                 return;
//             }
//             db.all("select * from test", function (err, res) {
//                 if (!err)
//                     console.log(JSON.stringify(res));
//                 else
//                     console.log(err);
//             });
//         })
//     });
// });