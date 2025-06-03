function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('次月の指示書依頼一覧作成', 'main')
    .addItem('送り状作成', 'executeTask')
    .addToUi();
}