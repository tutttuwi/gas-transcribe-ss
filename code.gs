function transcribe() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('設定');
  var settingsRange = settingsSheet.getRange('B4:L' + settingsSheet.getLastRow());
  var settingsData = settingsRange.getValues();

  for (var i = 0; i < settingsData.length; i++) {
    var sourceSpreadsheetId = settingsData[i][0];
    var sourceSheetName = settingsData[i][2];
    var sourceRange = settingsData[i][3];
    var targetSpreadsheetId = settingsData[i][4];
    var targetSheetName = settingsData[i][6];
    var targetCol = settingsData[i][7];
    var targetRow = settingsData[i][8];
    var initializeRange = settingsData[i][9];
    var append = settingsData[i][10];

    if (!sourceSpreadsheetId) continue; // SKIP

    try {
      var sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(sourceSheetName);
      const ssName = SpreadsheetApp.openById(sourceSpreadsheetId).getName();
      const sheetName = sourceSheet.getSheetName();
      var sourceData = sourceSheet.getRange(sourceRange).getValues();
      sourceData = sourceData.filter(ary => ary.filter(v => v != "").length)
      // Logger.log(sourceData);
      // Logger.log(sourceData.length);
      // Logger.log(sourceData[0].length);
      for (let i = 0; i < sourceData.length; i++) {
        if (i = 0) {
          // ヘッダ用
          sourceData[i].unshift("シート名");
          sourceData[i].unshift("スプシ名");
        } else {
          sourceData[i].unshift(sheetName);
          sourceData[i].unshift(ssName);
        }
      }
      // Logger.log(sourceData);

      var targetSheet = SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(targetSheetName);
      Logger.log("initializeRange:" + initializeRange);
      if (initializeRange) {
        targetSheet.getRange(initializeRange).clearContent();
        // targetSheet.clear(); // 初期化有無がTRUEの場合、転記先シートの情報を全削除
      }

      var startRow = append === true ? targetSheet.getLastRow() + 1 : targetRow; // 追記有無がTRUEの場合、最終行から転記

      targetSheet.getRange(startRow, targetCol, sourceData.length, sourceData[0].length).setValues(sourceData);

      // 転記成功の場合、設定シートに日時とOKを記入
      settingsSheet.getRange('M' + (i + 4)).setValue(new Date());
      settingsSheet.getRange('N' + (i + 4)).setValue('OK');
      Logger.log('転記成功: ' + (i + 1) + '行目');
    } catch (error) {
      Logger.log('転記エラー: ' + error);
      // 転記失敗の場合、設定シートに日時とNGを記入
      settingsSheet.getRange('M' + (i + 4)).setValue(new Date());
      settingsSheet.getRange('N' + (i + 4)).setValue('NG');
    }
  }
}