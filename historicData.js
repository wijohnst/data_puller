function getHistoricDataSpreadsheet(){
  return SpreadsheetApp.openById('1voGpdx-JTQdjhOtu7SRmk6h-qpDcA9g0e9JLApMk9Tw');
}

function getMostRecentM2DData(){
  const m2d = SpreadsheetApp.getActive().getSheetByName('Month-To-Date ADH Data');
  return  m2d.getRange(1,1,m2d.getLastRow(),m2d.getLastColumn()).getValues();
}

function sendM2dToHistoricDataSpreadsheet({targetMonth, targetYear}){
  const reportTitle = `EOM Data - ${targetMonth} ${targetYear}`;
  const dataToWrite = getMostRecentM2DData();
  const targetSpreadsheet = getHistoricDataSpreadsheet();
  const newSheet = targetSpreadsheet.insertSheet(reportTitle);
  newSheet.getRange(1,1,dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
  confirmHistoricSync(reportTitle);
}