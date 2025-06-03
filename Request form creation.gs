function executeTask() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentDate = new Date();
  var year = currentDate.getFullYear() - 2018; // 令和表記
  var month = currentDate.getMonth() + 1;
  var sheetName = 'R' + year + '.' + month + '月依頼分';
  var sourceSheet = ss.getSheetByName(sheetName);

  if (!sourceSheet) {
    Logger.log(sheetName + ' シートが見つかりません。');
    return;
  }

  var data = sourceSheet.getDataRange().getValues();
  var header = data[1];
  var medicalInstitutionIndex = header.indexOf('医療機関名');
  var centerIndex = header.indexOf('2事業所');

  if (medicalInstitutionIndex === -1 || centerIndex === -1) {
    Logger.log('必要な列が見つかりません。');
    return;
  }

  var targetColumns = ['利用者名', '利用者カナ', '生年月日', '指示区分', '指示書期間(至）', '指示書依頼期間', 'リハビリ\n週/回数', 'リハビリ\n時間', 'その他/依頼内容'];
  var targetIndexes = targetColumns.map(function(col) {
    return header.indexOf(col);
  });

  if (targetIndexes.includes(-1)) {
    Logger.log('必要な列が見つかりません。');
    return;
  }

  var groups = {};

  // グループ化（センター南の有無で2つのグループを作成）
  data.slice(2).forEach(function(row) {
    var medicalInstitution = row[medicalInstitutionIndex];
    var centerValue = row[centerIndex];
    if (medicalInstitution) {  // 医療機関名に値がある場合のみ対象
      var groupKey = centerValue ? medicalInstitution + '（南）' : medicalInstitution;
      if (!groups[groupKey]) {
        groups[groupKey] = [];
      }
      groups[groupKey].push(row);
    }
  });

  for (var key in groups) {
    var group = groups[key];
    var isCenter = key.includes('（南）');
    var templateSheetName = isCenter ? '送り状 (センター南)' : '送り状';
    var templateSheet = ss.getSheetByName(templateSheetName);

    if (!templateSheet) {
      Logger.log(templateSheetName + ' シートが見つかりません。');
      continue;
    }

    var newSheet = templateSheet.copyTo(ss);
    newSheet.setName(key); // グループキーをタブ名として設定

    // 「医療機関名」列の値を新しいシートのB26に入力
    newSheet.getRange('B26').setValue(key.replace('（南）', '')); // タブ名の「（南）」は削除

    // 指定された列の値を新しいシートのB34を開始にして入力
    var valuesToCopy = group.map(function(row) {
      return targetIndexes.map(function(index) {
        return row[index];
      });
    });

    if (valuesToCopy.length <= 10) {
      // 10行以下の場合はB34から貼り付け
      newSheet.getRange(34, 2, valuesToCopy.length, valuesToCopy[0].length).setValues(valuesToCopy);
    } else {
      // A25:J47をコピーしてA49:J71に貼り付け（行の高さを維持）
      var sourceRange = newSheet.getRange('A25:J47');
      var destinationRange = newSheet.getRange('A49:J71');
      sourceRange.copyTo(destinationRange, { contentsOnly: false });

      // 最初の10行はB34から貼り付け
      newSheet.getRange(34, 2, 10, valuesToCopy[0].length).setValues(valuesToCopy.slice(0, 10));

      // 11行目以降はB58から貼り付け
      var remainingValues = valuesToCopy.slice(10);
      newSheet.getRange(58, 2, remainingValues.length, remainingValues[0].length).setValues(remainingValues);
    }

    // === 追加機能開始 ===
    var medicalInstitutionName = key.replace('（南）', '');

    var instructionsFormatSheet = ss.getSheetByName('指示書フォーマット同封');
    if (!instructionsFormatSheet) {
      Logger.log('指示書フォーマット同封 シートが見つかりません。');
      continue;
    }

    var instructionValues = instructionsFormatSheet.getRange('B3:B').getValues().flat().filter(String);

    if (instructionValues.includes(medicalInstitutionName)) {
      var instructionTemplateSheet = ss.getSheetByName('指示書');
      if (!instructionTemplateSheet) {
        Logger.log('指示書 シートが見つかりません。');
        continue;
      }

      var copiedInstructionSheet = instructionTemplateSheet.copyTo(ss);
      copiedInstructionSheet.setName(medicalInstitutionName + ' 指示書');

      // 新しいシートの直後に「指示書」シートを移動
      ss.setActiveSheet(copiedInstructionSheet);
      ss.moveActiveSheet(newSheet.getIndex() + 1);
    }
    // === 追加機能終了 ===
  }

  // 医療機関リストから値をコピー
  copyValuesFromMedicalInstitutionList();
}

function copyValuesFromMedicalInstitutionList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var targetSheet = findSheetByNamePart(ss, "ここから訪問看護リハビリケア_医療機関リスト");

  if (!targetSheet) {
    Logger.log("対象シートが見つかりません");
    return;
  }

  var aValues = targetSheet.getRange("A:A").getValues().flat();
  var dValues = targetSheet.getRange("D:D").getValues().flat();
  var eValues = targetSheet.getRange("E:E").getValues().flat();

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();

    // 「（南）」がタブ名に含まれている場合は除去してから検索
    var normalizedSheetName = sheetName.replace("（南）", "");

    var index = aValues.indexOf(normalizedSheetName);
    if (index !== -1) {
      sheet.getRange("B3").setValue("　　　" + aValues[index]);
      sheet.getRange("B1").setValue("　　　　" + dValues[index]);
      sheet.getRange("B2").setValue("　　　　" + eValues[index]);
      Logger.log("値をコピーしました: " + sheetName);
    }
  });
}

function findSheetByNamePart(spreadsheet, namePart) {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(namePart) !== -1) {
      return sheets[i];
    }
  }
  return null;
}
