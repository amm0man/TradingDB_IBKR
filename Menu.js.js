function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IBKR')
    .addItem('1. Import from folder', 'importIBKRFromFolder')
    .addItem('2. Map Import → Master', 'mapIBKRImportToMaster')
    .addToUi();
}