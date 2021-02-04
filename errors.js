function handleErrors({type, errorText}) {
  switch(type){
    case 'alert':
    throwAlert(errorText);
  }
}

function throwAlert(errorText){
  SpreadsheetApp.getUi().alert(errorText);
}

function confirmReportsLoaded(){
  SpreadsheetApp.getUi().alert(`Got the emails from your inbox and created your new reports. Next, select 'Sync Reports' from the Summary Reports Sync menu to sync this spreadsheet with the latest reports.` )
}

function confirmReportsSync(){
  SpreadsheetApp.getUi().alert('Your reports updated to the latest version in your Google Drive account.')
}

function confirmSummaryReportUpdated(){
  SpreadsheetApp.getUi().alert(`The Summary Report worksheet successfully updated using the latest data. See the Troubleshooting tab to confirm the file names of the reports that were used to generate this Summary Report.`)
}

function confirmHistoricSync(reportTitle){
  SpreadsheetApp.getUi().alert(`Historic Data Updated. See sheet: ${reportTitle}.`);
}

function handleGetColSuccess(searchParam, targetSheet){
  console.log(`Attempted to return ${searchParam} from ${targetSheet}...`)
}