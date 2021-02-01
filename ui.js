function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('Summary Report Sync')
    .addItem('Add Reports to Drive','Tymeshift_Summary_Email_Client.sendToDrive')
    .addSeparator()
    .addItem('Sync Reports','populateSheets')
    .addSeparator()
    .addItem('Update Summary Report', 'writeCompleteReport')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Tools')
      .addItem('Update active users','writeAgentsToSheet'))
    .addToUi();
}
