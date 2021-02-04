/* GET COLUMN DATA BY HEADING */

function getColByHeading(targetSheet, targetHeading){

  const [ headers ] = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues(); 

  const targetIndex = headers.findIndex(header => header === targetHeading);

  const  targetData  = targetSheet.getRange(1,targetIndex + 1,targetSheet.getLastRow(),1).getValues(); 

  return targetData;
}

/* GET SUMMARY REPORT AS 'sheet' CLASS*/

function getSummaryReport(){
  const sheet = SpreadsheetApp.getActive();
  return sheet.getSheetByName('Summary Report');
}

/* GET AN ARRAY OF UNTRACKED TIMES DURATIONS*/

function getUntrackedTimes(targetSheet){
  return getColByHeading(targetSheet,'UNTRACKED');
}

/* GET VALUES FROM HEADER ROW OF TARGET SHEET */

function getHeaders(targetSheetName){
  const sheet = SpreadsheetApp.getActive().getSheetByName(targetSheetName);
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues().flat();
}

/*GET AN ARRAY OF WORKING TIMES DURATIONS */

function getWorkingTimes(targetSheet){
   return getColByHeading(targetSheet,'Working Time');
}

function getTimesInGeneral(targetSheet){
  return getColByHeading(targetSheet,'GENERAl TASKS').flat().filter(duration => duration !== 'GENERAl TASKS');
}

/* PARSE DURATIONS FOR CALCULATIONS */

function parseTimes(dateStringArr){
  return dateStringArr.map((cellValue,index) => {
    const JSDate = Date.parse(cellValue);
    return { time : getGoogleDate(JSDate), agentIndex: index}
  })
}

/* CONVERTS A JAVASCRIPT DATETIME OBJECT INTO A GOOGLE DATE FOR CALCULATIONS */

function getGoogleDate(JSDate){ 
  const D = new Date(JSDate);
    const epoch = new Date(Date.UTC(1899,11,30,0,0,0,0));
    return ((D.getTime() - epoch.getTime())/60000 - D.getTimezoneOffset())
}

function test(){
  writeAdhOccToSheet()
}

/* ALLOWS YOU TO LOOK UP AND RETURN DATA FROM A COLUMN USING A HEADER STRING AS A KEY */

function getColByHeading(targetSheet, targetHeading){

//This conditional allows you to pass in a string (<T : SheetName>) for your targetSheet instead of a sheet object, which is cumbersom sometimes, while preserving those calls that do pass a sheet object
  if(typeof targetSheet === 'string'){ 
    const spreadsheet = SpreadsheetApp.getActive();
    targetSheet = spreadsheet.getSheetByName(targetSheet);
  }

//If the conditional is skipped, targetSheet should be of class Spreadsheet.Sheet (<targetSheet : Spreadsheet.Sheet>) 
  const [ headers ] = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues();
  const targetIndex = headers.findIndex(header => header.trim() === targetHeading) + 1;
  const  targetData  = targetSheet.getRange(1,targetIndex,targetSheet.getLastRow(),1).getValues();

  return targetData; 

  /* The 'first header' bug. For some reason, the only way to access the first header index with findIndex() is the re-write the first header. For an example, see datapuller.gs >> writeDataToSheet() >> sheetMin.setValue('Name'). I have literally no idea why. - 1/27/21 - WJ */
}

/* RETURN AN ARRAY OF ACTIVE AGENTS BASED ON THE 'Reporting Data' SPREADSHEET*/

function getActiveAgents(){
    
    const sheet = SpreadsheetApp.getActive();
    const reportingData = sheet.getSheetByName('Reporting Data');
    const [ headers ] = reportingData.getRange(1,1,1,reportingData.getLastColumn()).getValues(); //Returns the heading for each column in the Reporting Data spreadsheet

    // const targetAgentsCol = headers.findIndex(header => header === 'Agent');
    const isAgent = getColByHeading(reportingData,'Is Agent');
    const allNames = getColByHeading(reportingData,'Agent')

    
    const activeAgents = isAgent.map((bool,index) => {
      if(typeof bool[0] !== 'string' && bool[0] === true){
        return allNames[index];
      }
    }).filter(value => value !== undefined).flat(); //Returns an error in GAS browser editor but still works

    return activeAgents;
}

/* RETURNS AN ARRAY OF ROUTES FOR MOVING ADHERENCE DATA FROM A SOURCE TO THE SUMMARY REPORT*/

function getRoutes(type){

if(type === 'ADH'){
  return [
    {
      sheetName: `7 Day ADH Data`,
      columnName: '7 DAY ADHERENCE'
    },
    {
      sheetName: '30 Day ADH Data',
      columnName: '30 DAY ADHERENCE'
    },
    {
      sheetName: 'Month-to-Date ADH Data',
      columnName: 'MONTH TO DATE ADHRENCE'
    }
  ]
}else if(type === 'OCC'){
  return [
    {
      sheetName: `7 Day ADH Data`,
      columnName: '7 DAY OCCUPANCY'
    },
    {
      sheetName: '30 Day ADH Data',
      columnName: '30 DAY OCCUPANCY'
    },
    {
      sheetName: 'Month-to-Date ADH Data',
      columnName: 'MONTH TO DATE OCCUPANCY'
    }
  ]
}else{
  console.error('INCORRECT REPORT TYPE')
}
  
}

function testWriteReportNamesToSheet(){
    const emailAttachments = getEmailAttachments().flat();
    const names = emailAttachments.map(attachment => attachment.getName());
    if(emailAttachments && emailAttachments.length > 0){
      writeReportNamesToSheet(names);
    }else{
      console.log('No attachements found...')
    }
    
}

function writeReportNamesToSheet(reportNames){
  
  const troubleshooting = SpreadsheetApp.getActive().getSheetByName('Troubleshooting');
  const headers = getHeaders('Troubleshooting');
  const targetColumnNum = headers.indexOf('CURRENT REPORT NAMES') + 1;

  const targetRange = troubleshooting.getRange(2,targetColumnNum, 3,1);
  targetRange.clearContent();

  reportNames.map((reportName,index) => {
    const targetCell = targetRange.getCell(index + 1, 1);
    targetCell.setValue(reportName);
  })

  
  
}


