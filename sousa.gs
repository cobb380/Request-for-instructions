function main() {
  var startTime = new Date();
  Logger.log("Main function start: " + startTime);
  
  // 日付でシートをフィルタリング
  filterSheetByDate();
  
  // 終了・中止タブの条件に基づいて行を削除
  removeRowsBasedOnConditions();
  
  // 新しい依頼シートを作成
  copySheetWithFormattedName();

  // データをコピー
  copyDataToBaseSheet();
  
  var endTime = new Date();
  Logger.log("Main function end: " + endTime);
  Logger.log("Total execution time: " + (endTime - startTime) + " ms");
}

// 日付でシートをフィルタリング
function filterSheetByDate() {
  var startTime = new Date();
  Logger.log("filterSheetByDate start: " + startTime);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // タブ名にHNC09000P1が含まれているシートを探す
  var sheet = findSheetByNamePart(spreadsheet, "HNC09000P1");
  
  if (!sheet) {
    Logger.log("指定されたシートが見つかりません");
    return;
  }
  
  // 既存のフィルタを解除
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  
  // 指定された列を取得
  var filterRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var visibleDates = getVisibleDates();
  Logger.log("Visible Dates: " + visibleDates);
  
  var columnIdx = getColumnIndexByName(sheet, ["指示書期間(至)", "指示書期間(至）"]);
  if (columnIdx === -1) {
    Logger.log("指定された列が見つかりません");
    return;
  }
  
  var allDates = sheet.getRange(2, columnIdx, sheet.getLastRow() - 1, 1).getValues().flat();
  Logger.log("All Dates: " + allDates);

  var hiddenDates = allDates.filter(date => !visibleDates.includes(date));
  Logger.log("Hidden Dates: " + hiddenDates);

  var criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(hiddenDates)
      .build();

  // フィルタを設定
  filterRange.createFilter();
  filterRange.getFilter().setColumnFilterCriteria(columnIdx, criteria);
  
  var endTime = new Date();
  Logger.log("filterSheetByDate end: " + endTime);
  Logger.log("filterSheetByDate execution time: " + (endTime - startTime) + " ms");
}

// 特定の名前部分を含むシートを探す
function findSheetByNamePart(spreadsheet, namePart) {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(namePart) !== -1) {
      return sheets[i];
    }
  }
  return null;
}

// 表示すべき日付を取得
function getVisibleDates() {
  var dates = [];
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth();
  
  // スクリプト実施月の末日
  var startDate = new Date(year, month, 15); // 実施月の15日
  // スクリプト実施翌月の末日
  var endDate = new Date(year, month + 1, 14); // 翌月14日

  for (var d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    dates.push(formatDateReiwa(d));
  }
  return dates;
}

// 日付を令和形式でフォーマット
function formatDateReiwa(date) {
  var year = date.getFullYear() - 2018;
  var month = (date.getMonth() + 1).toString();
  var day = date.getDate().toString();
  return '令和' + year + '年' + month + '月' + day + '日';
}

// 指定された列名を含む列のインデックスを取得
function getColumnIndexByName(sheet, names) {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues();
  for (var i = 0; i < values[0].length; i++) {
    if (names.includes(values[0][i])) {
      return i + 1;
    }
  }
  return -1;
}

// 条件に基づいて行を削除
function removeRowsBasedOnConditions() {
  var startTime = new Date();
  Logger.log("removeRowsBasedOnConditions start: " + startTime);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var endOrCancelSheet = spreadsheet.getSheetByName("終了・中止");
  
  if (!endOrCancelSheet) {
    Logger.log("「終了・中止」シートが見つかりません");
    return;
  }

  // 終了・中止シートの値を取得
  var dataRange = endOrCancelSheet.getRange(4, 1, endOrCancelSheet.getLastRow() - 3, endOrCancelSheet.getLastColumn()).getValues();
  var valuesToCheck = dataRange.filter(row => row[2] === "サービス終了","サービス中止").map(row => row[0]);

 // valuesToCheckの内容をログに表示
  Logger.log("valuesToCheck: " + JSON.stringify(valuesToCheck));

  // HNC09000P1が含まれているシートを探す
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.getName().includes('HNC09000P1')) {
      // シートのデータを取得
      var sheetData = sheet.getRange(1, 1, sheet.getLastRow()).getValues().flat();
      for (var j = sheetData.length - 1; j >= 0; j--) {
        if (valuesToCheck.includes(sheetData[j])) {
          sheet.deleteRow(j + 1);
        }
      }
    }
  }
  Logger.log("削除完了");
  
  var endTime = new Date();
  Logger.log("removeRowsBasedOnConditions end: " + endTime);
  Logger.log("removeRowsBasedOnConditions execution time: " + (endTime - startTime) + " ms");
}

// 来月の依頼シートを作成
function copySheetWithFormattedName() {
  var startTime = new Date();
  Logger.log("copySheetWithFormattedName start: " + startTime);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = spreadsheet.getSheetByName("フォーマット原本");
  
  if (!originalSheet) {
    Logger.log("指定されたシートが見つかりません");
    return;
  }

  var today = new Date();
  var year = today.getFullYear() - 2018; // 令和表記
  var month = today.getMonth() + 1; // 月は0から始まるので1を加える
  var newSheetName = "R" + year + "." + month + "月依頼分";
  
  // 同じ名前のシートが存在するか確認
  var existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(newSheetName + "のシートはすでに存在します。上書きしますか？", ui.ButtonSet.OK_CANCEL);

    if (response == ui.Button.CANCEL) {
      Logger.log("シートのコピーをキャンセルしました");
      return;
    }

    // 既存のシートを削除
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // シートをコピーして名前を変更
  var newSheet = originalSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);
  
  // 7番目の位置に新しいシートを移動
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(7);
  
  Logger.log("シートがコピーされました: " + newSheetName);
  
  var endTime = new Date();
  Logger.log("copySheetWithFormattedName end: " + endTime);
  Logger.log("copySheetWithFormattedName execution time: " + (endTime - startTime) + " ms");
}

// フィルタ適用後の可視行のみを取得
function getFilteredData(sheet, columnIndexes) {
  var startTime = new Date();
  Logger.log("getFilteredData start: " + startTime);
  
  if (!sheet) {
    throw new Error("シートが無効です");
  }
  const range = sheet.getDataRange();
  const values = range.getValues();
  const filteredValues = [];

  for (let i = 1; i < values.length; i++) { // ヘッダー行を除外してデータ行をループ
    if (!sheet.isRowHiddenByFilter(i + 1)) {
      let row = [];
      for (let j = 0; j < columnIndexes.length; j++) {
        row.push(values[i][columnIndexes[j] - 1]); // columnIndexesは1から始まるため-1
      }
      filteredValues.push(row);
    }
  }
  
  var endTime = new Date();
  Logger.log("getFilteredData end: " + endTime);
  Logger.log("getFilteredData execution time: " + (endTime - startTime) + " ms");

  return filteredValues;
}

// 来月の依頼シートにコピー
function copyDataToBaseSheet() {
  var startTime = new Date();
  Logger.log("copyDataToBaseSheet start: " + startTime);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var year = today.getFullYear() - 2018; // 令和表記
  var month = today.getMonth() + 1; // 月は0から始まるので1を加える
  var baseSheetName = "R" + year + "." + month + "月依頼分";
  var baseSheet = spreadsheet.getSheetByName(baseSheetName);

  if (!baseSheet) {
    Logger.log(baseSheetName + "シートが見つかりません");
    return;
  }

  // フィルタ適用シートを取得
  var sourceSheet = findSheetByNamePart(spreadsheet, "HNC09000P1");
  Logger.log("Source Sheet: " + sourceSheet);

  if (!sourceSheet) {
    Logger.log("フィルタ適用シートが見つかりません");
    return;
  }

  // フィルタを適用
  filterSheetByDate();

  // コピーする列のインデックスを取得
  var columnsToCopy = ["利用者名", "利用者カナ", "生年月日", "指示区分", "指示書期間(至）"];
  var additionalColumns = ["医療機関名", "主治医名"];
  var columnIndexes = columnsToCopy.map(name => getColumnIndexByName(sourceSheet, [name]));
  var additionalIndexes = additionalColumns.map(name => getColumnIndexByName(sourceSheet, [name]));

  // インデックスの確認
  Logger.log("Column Indexes: " + columnIndexes);
  Logger.log("Additional Indexes: " + additionalIndexes);

  // フィルタ適用後のデータを取得
  var filteredData = getFilteredData(sourceSheet, columnIndexes);
  var additionalData = getFilteredData(sourceSheet, additionalIndexes);

  // データをコピー
  if (filteredData.length > 0) {
    baseSheet.getRange(3, 3, filteredData.length, filteredData[0].length).setValues(filteredData);
  }
  
  if (additionalData.length > 0) {
    baseSheet.getRange(3, 12, additionalData.length, additionalData[0].length).setValues(additionalData);
  }

  // G列とE列の日付を西暦表記に変換
  convertJapaneseErasInColumns(baseSheet);

  Logger.log("データのコピーが完了しました");
  
  var endTime = new Date();
  Logger.log("copyDataToBaseSheet end: " + endTime);
  Logger.log("copyDataToBaseSheet execution time: " + (endTime - startTime) + " ms");
}

// 和暦を西暦に変換する関数
function convertJapaneseEraToGregorian(eraDate) {
  var eraMap = {
    "明治": 1868,
    "大正": 1912,
    "昭和": 1926,
    "平成": 1989,
    "令和": 2019
  };

  var eraPattern = /(明治|大正|昭和|平成|令和)(\d+)年(\d+)月(\d+)日/;
  var match = eraDate.match(eraPattern);

  if (match) {
    var era = match[1];
    var year = parseInt(match[2], 10);
    var month = parseInt(match[3], 10) - 1; // 月は0から始まる
    var day = parseInt(match[4], 10);
    var gregorianYear = eraMap[era] + year - 1;

    return new Date(gregorianYear, month, day);
  }
  return null;
}

// G列とE列の和暦を西暦に変換
function convertJapaneseErasInColumns(sheet) {
  var gColumn = 7; // G列
  var eColumn = 5; // E列

  // G列の変換
  var gRange = sheet.getRange(3, gColumn, sheet.getLastRow() - 2);
  var gValues = gRange.getValues();
  for (var i = 0; i < gValues.length; i++) {
    var date = convertJapaneseEraToGregorian(gValues[i][0]);
    if (date) {
      gValues[i][0] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }
  }
  gRange.setValues(gValues);

  // E列の変換
  var eRange = sheet.getRange(3, eColumn, sheet.getLastRow() - 2);
  var eValues = eRange.getValues();
  for (var i = 0; i < eValues.length; i++) {
    var date = convertJapaneseEraToGregorian(eValues[i][0]);
    if (date) {
      eValues[i][0] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }
  }
  eRange.setValues(eValues);
}
