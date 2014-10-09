// writeExcel.js
// Excel�t�@�C���X�V
// (windows only)
var fso = new ActiveXObject("Scripting.FileSystemObject");

// ��������t�@�C�����擾
// @return {string} �t�@�C���t���p�X
function getFileName(){

  var params = WScript.arguments;

  if(params.length > 0){
    var file = params(0);
    return fso.GetAbsolutePathName(file);
  }

  return "";
}

// �t�@�C����������t���擾����
// @param {string} filePath
// @return {Array} ���t ["yyyymmdd", "yyyy", "mm", "dd"]
function getWorkDay(filePath){

  var fileName = fso.GetBaseName(filePath);

  if(/^\d{8}/.test(fileName)){
    // �t�@�C��������N�������擾
    var arr = /^(\d{4})(\d{2})(\d{2})/.exec(fileName);

    if(arr.length != 4){
      // �N�����̎擾���s
      WScript.Quit(-1);
    }

    return arr;

  }else{
    // �t�@�C����������ƈႤ -> Error
    WScript.Quit(-1);
  }

  return [];
}

// Excel�X�V����
// @param {string} Excel�t�@�C���p�X
// @param {Array} ���t
function updateExcel(filePath, workDay){
  var xls = null;
  var book = null;
  var sheet = null;

  try{
    // Excel�N��
    xls = new ActiveXObject("Excel.Application");
    // �t�@�C��Open
    book = xls.Workbooks.Open(filePath);
    // �V�[�g�擾
    sheet = book.Worksheets(1);

    // ��Ɠ��Z�b�g
    sheet.Range("Z10").Value = workDay[1] + "/" + workDay[2] + "/" + workDay[3];

    // �V�[�g���ύX
    sheet.Name = workDay[3]; // �����Z�b�g

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


// �又��
function main(){
  // ��������t�@�C�����擾
  var filePath = getFileName();

  // �Y���t�@�C�����Ȃ���΃G���[�Ƃ��ďI��
  if(!fso.FileExists(filePath)){
    // �G���[�I��
    WScript.Quit(-1);
  }

  // �t�@�C����������t���擾
  var wd = getWorkDay(filePath);

  // �X�V����
  updateExcel(filePath, wd);

  // �I��
  WScript.Echo("�X�V�I��.");
}

main();
