// function writeGeneralTimesToSheet(){
  
//   const sheet = SpreadsheetApp.getActive();
//   const summaryReport = sheet.getSheetByName('Summary Report');
  
//   const generalTimes = getGeneralTimes();
//   const agents = getColByHeading(summaryReport, 'AGENT').flat().filter(name => name !== "");

//   const headers = summaryReport.getRange(1,1,1,summaryReport.getLastColumn()).getValues();
//   const targetColNum = headers[0].indexOf('% TIME GENERAL') + 1;
//   const targetColumn = summaryReport.getRange(2,targetColNum,summaryReport.getLastRow(),1);

//   targetColumn.clearContent();
  
//   generalTimes.map(({agent,data},index) => {

//     const percentGeneral = data.percentGeneral;
//     const targetRow = agents.indexOf(agent);
    
//     if(targetRow > 0){
//         const targetCell = targetColumn.getCell(targetRow,1);
//           if(Number.isNaN(percentGeneral)){
//             Logger.log(percentGeneral);
//             targetCell.setValue('No data available.')
//           }
//           else{
//             targetCell.setValue(percentGeneral + '%');
//           }
//     }
//   }) 
// }
function getGeneralTimes(){

  const sheet = SpreadsheetApp.getActive();
  const targetSheet = sheet.getSheetByName('7 Day ADH Data');

  const generalTasksArr = getTimesInGeneral(targetSheet);

  const agents = getColByHeading(targetSheet,'Name').flat().filter(name => name !== "Name");
  const generalTimes = parseTimes(generalTasksArr);
  const workingTimes = parseTimes(getWorkingTimes(targetSheet).flat().filter(time => time !== 'Working Time'));

    const generalTimesArr = generalTimes.map(({time, agentIndex},index) => {

    const generalTime = time;
    const workingTime = workingTimes[index].time;
     
    return { 
      agent: agents[agentIndex],
      data: {
         agentIndex: agentIndex, 
         percentGeneral : Math.round((generalTime / workingTime) * 100) 
      } 
    }
  })
 return generalTimesArr;
}

