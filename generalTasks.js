function getGeneralTimes(){

  const sheet = SpreadsheetApp.getActive();
  const targetSheet = sheet.getSheetByName('7 Day ADH Data');

  const generalTasksArr = getTimesInGeneral(targetSheet);
    let agents;

  try{
    agents = getColByHeading(targetSheet,'Name').flat().filter(name => name !== "Name");
  }
  catch(err){
    throwAlert(`System cannot find a column named 'Name1' in the sheet '7 Day ADH Data'. If this query needs to be updated to match a new header name, please see generalTasks.gs >> getGeneralTimes() and updated the query parameter in the try block.`)
  }
  finally{
    console.log('Attempted to get a list of agents...')
  }
  
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

