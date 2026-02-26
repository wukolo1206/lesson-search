// Code.gs

// 設定網頁標題
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('閱讀教學策略查詢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 修改後的讀取函式：一次讀取多個工作表
function getData() {
  // 使用明確的試算表 ID 確保連線穩定
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');

  // 1. 指定要讀取的工作表名稱清單 (請確認名稱跟下方標籤一模一樣)
  var targetSheets = ['三年級', '四年級', '五年級', '六年級'];

  var combinedData = [];

  // 2. 使用迴圈，一個一個工作表進去抓資料
  targetSheets.forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);

    // 確保工作表存在才讀取，避免報錯
    if (sheet) {
      // 讀取該工作表所有資料
      var data = sheet.getDataRange().getValues();

      // 確保資料不只一行（避免只有標題沒有內容）
      if (data.length > 1) {
        var headers = data[0]; // 第一列是標題
        var rows = data.slice(1); // 第二列之後是內容

        // 將每一列資料轉換成物件 (Key-Value)
        var sheetData = rows.map(function (row) {
          var obj = {};
          headers.forEach(function (header, index) {
            // 移除標題前後空白，確保對應準確
            var key = String(header).trim();
            obj[key] = row[index];
          });
          return obj;
        });

        // 將處理好的資料合併到大陣列中
        combinedData = combinedData.concat(sheetData);
      }
    }
  });

  // 3. 回傳整合後的所有年級資料
  return combinedData;
}
