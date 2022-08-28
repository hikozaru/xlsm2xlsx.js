// https://docs.microsoft.com/ja-jp/office/vba/api/excel.xlfileformat
var xlWorkbookDefault = 51;

var ExcelApp = new ActiveXObject( "Excel.Application" );

// *.jsファイルにDrag&Dropされたファイルを順番に処理
var args = WScript.Arguments;
for (var i = 0; i < args.length; i++ ) {
  var orgFilename = args(i);

  // 特定の拡張子のみ処理を実施
  var orgExtension = orgFilename.split('.').pop();
  if( orgExtension == 'xlsm' ){
    var baseName = orgFilename.substring(0, orgFilename.indexOf('.'+orgExtension));

    // ファイルを開いて形式を変更して保存
    var book = ExcelApp.Workbooks.Open( orgFilename );
    ExcelApp.DisplayAlerts= false;
    book.SaveAs( baseName + '.xlsx' ,  xlWorkbookDefault );
    ExcelApp.DisplayAlerts= true;
    ExcelApp.Quit();
  }
}
WScript.Echo( "finish" );
ExcelApp = null;
