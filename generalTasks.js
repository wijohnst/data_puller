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

