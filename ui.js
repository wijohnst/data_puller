function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('Summary Report Sync')
    .addItem('Add Reports to Drive','sendToDrive')
    .addSeparator()
    .addItem('Sync Reports','populateSheets')
    .addSeparator()
    .addItem('Update Summary Report', 'writeCompleteReport')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Tools')
      .addItem('Update active users','writeAgentsToSheet')
      .addItem('File historic data', 'promptForTargetMonth'))
    .addToUi();
}

function promptForTargetMonth(){
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('What is the target month for this data? (EX: January)');
  if(response.getSelectedButton() === ui.Button.OK){
    promptForTargetYear(response.getResponseText());
  }else{
    throwAlert('Historic data update cancelled.')
  }
}

function promptForTargetYear(targetMonth){
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('What is the target year for this data? (EX: 2021)');
  if(response.getSelectedButton() === ui.Button.OK){
    const reportTitle = {targetMonth: targetMonth, targetYear: response.getResponseText()}
    sendM2dToHistoricDataSpreadsheet(reportTitle)
  }else{
    throwAlert('Historic data update cancelled.')
  }
}

