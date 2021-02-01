function writeCompleteReport(){
  writeUntrackedToSheet();
  writeGeneralTimesToSheet();
  writeAdhOccToSheet();
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

  const sheet = SpreadsheetApp.getActive();
  const summaryReport = sheet.getSheetByName('Summary Report');

  const agents = getColByHeading(summaryReport,'AGENT').flat().filter(name => name !== "");
  const percentUntrackedArr = getPercentUntracked('7 Day ADH Data');

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
  const agents = getColByHeading(summaryReport, 'AGENT').flat().filter(name => name !== "");

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

    const agents = getColByHeading(sheetName,'Name').flat().filter(name => name !== 'Name');
    const activeAgents = getColByHeading('Summary Report', 'AGENT').flat().filter(agent => agent !== 'AGENT');
    const data = getColByHeading(sheetName, dataTarget).flat().filter(data => data !== dataTarget);
    const summaryReport = getSummaryReport();
    const headers = getHeaders('Summary Report');
    const targetColNum = headers.indexOf(columnName) + 1;
    const targetRange = summaryReport.getRange(2,targetColNum, summaryReport.getLastRow(),1);

    targetRange.clearContent();

  // let test = () => activeAgents.map(name => data[agents.indexOf(name)]);
  // console.log(test());

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

function writeADHToSheet(){

  const routes = getADHRoutes();

  routes.map(({ sheetName, columnName }) => {

    const agents = getColByHeading(sheetName, 'Name').flat().filter(name => name !== 'Name');
    const activeAgents = getColByHeading('Summary Report', 'AGENT').flat().filter(agent => agent !== 'AGENT');
    const targetADHData = getColByHeading(sheetName, 'Schedule adherence').flat().filter(adhData => adhData !== 'Schedule adherence');
    const summaryReport = getSummaryReport();
    const headers = getHeaders('Summary Report');
    const targetColNum = headers.indexOf(columnName) + 1;
    const targetRange = summaryReport.getRange(2,targetColNum, summaryReport.getLastRow(),1);

    targetRange.clearContent();

    const writeData = () => {
      return activeAgents.map(name => targetADHData[agents.indexOf(name)]).filter(num => Boolean(num));
    }
      
    writeData().map((adhNum, index) => {
    
      const targetAgent = activeAgents[index];
      const targetRow = activeAgents.indexOf(targetAgent) + 1;

      if(targetRow > 0){
        const targetCell = targetRange.getCell(targetRow, 1);
        targetCell.setValue(adhNum);
      }
    })
  })
}
