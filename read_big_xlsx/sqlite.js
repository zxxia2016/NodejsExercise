var sqlite3 = require('sqlite3').verbose();

var DB = DB || {};

DB.SqliteDB = function () {
};

DB.SqliteDB.prototype.createDataBase = function (file) {
    return new Promise((resolve, reject) => {
        var db = new sqlite3.Database(file, function (err) {
            if (err) {
                DB.printErrorInfo(err);
                reject(err);
                return;
            }
            DB.db = db;
            resolve();
        });
    })
};
DB.printErrorInfo = function (err) {
    console.error("Error Message:" + err.message);
};

DB.SqliteDB.prototype.createTable = function (sql) {
    return new Promise((resolve, reject) => {
        DB.db.serialize(function (err) {
            if (err) {
                DB.printErrorInfo(err);
                return;
            }
            DB.db.run(sql, function (err) {
                if (err) {
                    DB.printErrorInfo(err);
                    reject(err);
                    return;
                }
                resolve();
            });
        });
    })
};

DB.SqliteDB.prototype.getRunSqlPromise = function (sql, array) {
    return new Promise((resolve, reject) => {
        DB.db.run(sql, array, function (err) {
            if (err) {
                DB.printErrorInfo(err);
                reject(err);
                return;
            }
            resolve();
        });
    })
};

//执行语句没有返回数据，只有返回结果
DB.SqliteDB.prototype.runSql = function (sql, array, cb) {
    DB.db.run(sql, array, function (err) {
        if (err) {
            DB.printErrorInfo(err);
            cb && cb(err);
            return;
        }
        cb && cb();
    });
};

//获取数据回调
DB.SqliteDB.prototype.getData = function (sql, array, cb) {
    DB.db.get(sql, array, function (err, row) {
        if (err) {
            DB.printErrorInfo(err);
            return;
        }
        cb && cb(row);
    });
};

//获取数据Promise
DB.SqliteDB.prototype.getDataPromise = function (sql, array) {
    return new Promise((resolve, reject) => {
        DB.db.get(sql, array, function (err, row) {
            if (err) {
                DB.printErrorInfo(err);
                reject(err);
                return;
            }
            resolve(row);
        });
    })
};

DB.SqliteDB.prototype.queryData = function (sql, cb) {
    DB.db.all(sql, function (err, rows) {
        if (null != err) {
            DB.printErrorInfo(err);
            cb && cb(null);
            return;
        }
        cb && cb(rows);
    });
};

DB.SqliteDB.prototype.close = function () {
    DB.db.close();
};

/// export SqliteDB.
exports.SqliteDB = DB.SqliteDB;