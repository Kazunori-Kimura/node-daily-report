// writeExcel.js
// Excelファイル更新
// (windows only)
var fso = new ActiveXObject("Scripting.FileSystemObject");

// 引数からファイル名取得
// @return {string} ファイルフルパス
function getFileName(){

  var params = WScript.arguments;

  if(params.length > 0){
    var file = params(0);
    return fso.GetAbsolutePathName(file);
  }

  return "";
}

// ファイル名から日付を取得する
// @param {string} filePath
// @return {Array} 日付 ["yyyymmdd", "yyyy", "mm", "dd"]
function getWorkDay(filePath){

  var fileName = fso.GetBaseName(filePath);

  if(/^\d{8}/.test(fileName)){
    // ファイル名から年月日を取得
    var arr = /^(\d{4})(\d{2})(\d{2})/.exec(fileName);

    if(arr.length != 4){
      // 年月日の取得失敗
      WScript.Quit(-1);
    }

    return arr;

  }else{
    // ファイル名が既定と違う -> Error
    WScript.Quit(-1);
  }

  return [];
}

// Excel更新処理
// @param {string} Excelファイルパス
// @param {Array} 日付
function updateExcel(filePath, workDay){
  var xls = null;
  var book = null;
  var sheet = null;

  try{
    // Excel起動
    xls = new ActiveXObject("Excel.Application");
    // ファイルOpen
    book = xls.Workbooks.Open(filePath);
    // シート取得
    sheet = book.Worksheets(1);

    // 作業日セット
    sheet.Range("Z10").Value = workDay[1] + "/" + workDay[2] + "/" + workDay[3];

    // シート名変更
    sheet.Name = workDay[3]; // 日をセット

    xls.DisplayAlerts = false;
    book.Save();
    book.Close();

  }catch(ex){
    try{
      book.Quit();
    }catch(e){}

    try{
      xls.Quit();
    }catch(e){}
  }
}


// 主処理
function main(){
  // 引数からファイル名取得
  var filePath = getFileName();

  // 該当ファイルがなければエラーとして終了
  if(!fso.FileExists(filePath)){
    // エラー終了
    WScript.Quit(-1);
  }

  // ファイル名から日付を取得
  var wd = getWorkDay(filePath);

  // 更新処理
  updateExcel(filePath, wd);

  // 終了
  WScript.Echo("更新終了.");
}

main();
