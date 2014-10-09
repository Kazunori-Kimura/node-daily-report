// app.js
// Excelファイルのコピーおよびスクリプトのキックを行う
var Q = require("q"),
  colors = require("colors"),
  moment = require("moment"),
  fs = require("fs"),
  path = require("path"),
  util = require("util"),
  exec = require('child_process').exec;

// load config
var config = require("./config.json");

// define color theme
colors.setTheme({
    debug: "grey",
    info: "green",
    warn: "yellow",
    error: "red"
});

// テンプレートファイルをコピーする
// @return promise
function copyFile(){
  var deffered = Q.defer();

  var filePath = path.resolve(__dirname, config.TEMPLATE);

  fs.exists(filePath, function(exists){
    if(exists){
      var m = moment();
      var destPath = util.format(config.FILE_NAME,
        m.format("YYYYMMDD"));
      destPath = path.resolve(__dirname, destPath);

      // file copy
      fs.linkSync(filePath, destPath);

      console.log("copy: %s".info, destPath);

      deffered.resolve(destPath);
    }else{
      deffered.reject("file not found.");
    }
  });

  return deffered.promise;
}

// Excel更新
// @param {string} filePath
// @return promise
function updateExcel(filePath){
  var deffered = Q.defer();

  var cmd = 'cscript lib/writeExcel.js "%s"';
  cmd = util.format(cmd, filePath);
  console.log(cmd.debug);

  var child = exec(cmd, function(err, stdout, stderr){
    if(err !== null){
      deffered.reject(stderr);
    }
    deffered.resolve("Done.");
  });

  return deffered.promise;
}

copyFile()
  .then(updateExcel)
  .catch(function(err){
    console.log(err.error);
  })
  .done(function(msg){
    console.log(msg.info);
  });
