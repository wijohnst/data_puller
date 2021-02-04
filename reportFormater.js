function writeCompleteReport(){
  writeUntrackedToSheet();
  writeGeneralTimesToSheet();
  writeAdhOccToSheet();
  confirmSummaryReportUpdated();
}

function writeAgentsToSheet(){

  const sheet = SpreadsheetApp.getActive();
  const summaryReport = sheet.getSheetByName('Summary Report');

  const agentNames = getActiveAgents();

  let targetRange = summaryReport.getRange(2,1,summaryReport.getLastRow());

  targetRange.clearContent();

  let writeRange = summaryReport.getRange(2,1,agentNames.length);

  agentNames.forEach((name,index) => {
    writeRange.getCell(index + 1, 1).setValue(name);
  })
  
}

/* writeUntrackedToSheet() & writeGeneralTimesToSheet() could be dried out and consolidated into a single function*/

function writeUntrackedToSheet(){

  
  const summaryReport = getSummaryReport();
  const targetSheet = 'Summary Report';

  let agents;
  const searchParam = 'AGENT';

  try{
    agents = getColByHeading(targetSheet, searchParam).flat().filter(agent => agent !== "" );
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText(searchParam, targetSheet, "reportFormater.gs >> writeUntrackedToSheet()")})
  }
  finally{
    handleGetColSuccess(searchParam,targetSheet)
  }

  const percentUntrackedArr = getPercentUntracked();
  const headers = summaryReport.getRange(1,1,1,summaryReport.getLastColumn()).getValues();
  const targetColNum = headers[0].indexOf('% TIME UNTRACKED') + 1;
  const targetColumn = summaryReport.getRange(2,targetColNum,summaryReport.getLastRow(),1);

  targetColumn.clearContent();

  percentUntrackedArr.map(({agent,data},index) => {

    const percentUntracked = data.percentUntracked;
    const targetRow = agents.indexOf(agent);
    
    if(targetRow > 0){
        const targetCell = targetColumn.getCell(targetRow,1);
          if(Number.isNaN(percentUntracked)){
            Logger.log(percentUntracked);
            targetCell.setValue('No data available.')
          }
          else{
            targetCell.setValue(percentUntracked);
          }
    }
  })
}

function writeGeneralTimesToSheet(){
  
  const sheet = SpreadsheetApp.getActive();
  const summaryReport = sheet.getSheetByName('Summary Report');
  
  const generalTimes = getGeneralTimes();
  const targetSheet = 'Summary Report';

  let agents;
  const searchParam = 'AGENT';

  try{
    agents = getColByHeading(targetSheet, searchParam).flat().filter(agent => agent !== "" );
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText(searchParam, targetSheet, "reportFormater.gs >> writeGeneralTimesToSheet()")})
  }
  finally{
    handleGetColSuccess(searchParam,targetSheet)
  }
  
  const headers = summaryReport.getRange(1,1,1,summaryReport.getLastColumn()).getValues();
  const targetColNum = headers[0].indexOf('% TIME GENERAL') + 1;
  const targetColumn = summaryReport.getRange(2,targetColNum,summaryReport.getLastRow(),1);

  targetColumn.clearContent();
  
  generalTimes.map(({agent,data},index) => {

    const percentGeneral = data.percentGeneral;
    const targetRow = agents.indexOf(agent);
    
    if(targetRow > 0){
        const targetCell = targetColumn.getCell(targetRow,1);
          if(Number.isNaN(percentGeneral)){
            Logger.log(percentGeneral);
            targetCell.setValue('No data available.')
          }
          else{
            targetCell.setValue(percentGeneral);
          }
    }
  }) 
}

function writeAdhOccToSheet(){ 
  const routesArr = [ 
    {routes: getRoutes('ADH') , dataTarget: 'Schedule adherence'}, 
    {routes: getRoutes('OCC'), dataTarget: 'Occupancy Rate'}
  ];

  routesArr.map(({ routes, dataTarget}) => writeReportToSheet(routes,dataTarget));
}

function writeReportToSheet(targetRoutes, dataTarget){

  targetRoutes.map(({ sheetName, columnName }) => {

  let agents;
  const agentSearchParam = 'Name';
  try{
    agents = getColByHeading(sheetName, agentSearchParam).flat().filter(agent => agent !== "" );
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText(agentSearchParam, sheetName, "reportFormater.gs >> writeReportToSheet()")})
  }
  finally{
    handleGetColSuccess(agentSearchParam,sheetName);
  }

  let activeAgents;
  const activeSearchParam = 'AGENT';
    try{
    activeAgents = getColByHeading('Summary Report', activeSearchParam).flat().filter(agent => agent !== 'AGENT' );
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText(activeSearchParam, 'Summary Report', "reportFormater.gs >> writeReportToSheet()")})
  }
  finally{
    handleGetColSuccess(activeSearchParam,'Summary Report');
  }

  let data;
    try{
    data = getColByHeading(sheetName, dataTarget).flat().filter(data => data !== dataTarget );
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText(dataTarget, sheetName, "reportFormater.gs >> writeReportToSheet()")})
  }
  finally{
    handleGetColSuccess(dataTarget,sheetName)
  }

  const summaryReport = getSummaryReport();
  const headers = getHeaders('Summary Report');
  const targetColNum = headers.indexOf(columnName) + 1;
  const targetRange = summaryReport.getRange(2,targetColNum, summaryReport.getLastRow(),1);

  targetRange.clearContent();

  const writeData = () => {
    return activeAgents.map(name => data[agents.indexOf(name)]).filter(num => num !== undefined);
  }
      
    writeData().map((value, index) => {
    
      const targetAgent = activeAgents[index];
      const targetRow = activeAgents.indexOf(targetAgent) + 1;
       
      if(targetRow >= 0){
        const targetCell = targetRange.getCell(targetRow, 1);
          targetCell.setValue(`${Math.round(value)}`);
      }
    })
  })
}

