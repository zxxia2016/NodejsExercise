var sqlite3 = require('sqlite3');
var db = new sqlite3.Database('./tmp/1.db',function(err) {
  if (err) {
      console.error(err);
      return;
  }
  console.log(`success`);
});