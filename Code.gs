function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function openUploader() {
    var html = HtmlService.createHtmlOutputFromFile('Index.html')
        .setWidth(750)
        .setHeight(480);
    SpreadsheetApp.getUi().showModalDialog(html, 'ファイルアップローダ');
}

// Access Tokenを取得する
function getOAuthToken() {
    try {
        DriveApp.getRootFolder();
        return ScriptApp.getOAuthToken();
    } catch (e) {
        Logger.log('Error obtaining OAuth token: ' + e.toString());
        throw new Error('Failed to obtain OAuth token: ' + e.message);
    }
}

// スクリプトプロパティを取得する
function getScriptProperties() {
    try {
        var properties = PropertiesService.getScriptProperties().getProperties();
        Logger.log('Script properties: ' + JSON.stringify(properties));
        return properties;
    } catch (e) {
        Logger.log('Error retrieving script properties: ' + e.toString());
        throw new Error('Failed to retrieve script properties: ' + e.message);
    }
}

// ファイルIDを保存する
function saveFileId(fileId) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('LAST_UPLOADED_FILE_ID', fileId);
}

// ファイルをスプレッドシートに変換する
function convertFileToSheet(fileId, parentFolderId) {
    try {
        Logger.log('convertFileToSheet called with fileId: ' + fileId + ' and parentFolderId: ' + parentFolderId);
        // ファイルを取得
        var file = DriveApp.getFileById(fileId);
        var folder = DriveApp.getFolderById(parentFolderId);
        var filename = file.getName();
        // ファイルの変換
        var newFile = Drive.Files.insert({
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [{"id": parentFolderId}],
            "title": filename
        }, file.getBlob());    
        // 元のファイルを削除
        file.setTrashed(true);
        Logger.log('Original Excel file has been moved to trash.');      
        // スプレッドシートのURLを取得
        var sheetUrl = newFile.alternateLink;
        Logger.log('Sheet URL: ' + sheetUrl);
        return sheetUrl;
    } catch (e) {
        Logger.log('Error converting file to Sheet: ' + e.toString());
        throw new Error('Failed to convert file to Sheet: ' + e.message);
    }
}

// Pickerから選択されたファイルIDを受け取ってファイルを変換する関数
function handleFileSelection(fileId) {
    try {
        Logger.log('handleFileSelection called with fileId: ' + fileId);
        
        var scriptProperties = PropertiesService.getScriptProperties();
        var parentFolderId = scriptProperties.getProperty('PARENT_FOLDER_ID');
        
        Logger.log('parentFolderId: ' + parentFolderId);
        
        if (!parentFolderId) {
            throw new Error('PARENT_FOLDER_ID is undefined');
        }
        
        var sheetUrl = convertFileToSheet(fileId, parentFolderId);
        return 'ファイルをスプレッドシートに変換し、指定のフォルダに保存しました。スプレッドシートのURL: ' + sheetUrl;
    } catch (e) {
        Logger.log('Error handling file selection: ' + e.toString());
        return 'Error: ' + e.message;
    }
}

// アップロードされたファイルを処理する関数
function processFile(fileName, fileContent) {
    try {
        // バイナリデータをMS932文字コードで読み取る
        var decodedContent = Utilities.newBlob(fileContent, 'application/octet-stream', fileName).getDataAsString('MS932');
        
        // デコードされた内容をシートに書き込む（例としてシートに書き込み）
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        sheet.getRange(1, 1).setValue(decodedContent);
        
        return 'File processed successfully!';
    } catch (e) {
        Logger.log('Error processing file: ' + e.toString());
        return 'Error: ' + e.message;
    }
}

// エラーメッセージの表示
function messager(msg) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(msg);
}