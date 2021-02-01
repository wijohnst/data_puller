function getPercentUntracked(sheetName){
  // Logger.log(`Error Trace: @getPercentUntracked()`);
  let sheet = SpreadsheetApp.getActive();
  let targetSheet = sheet.getSheetByName(sheetName || '7 Day ADH Data');
  
  const agents = getColByHeading(targetSheet,'Name').flat().filter(name => name !== 'Name' && name !== ''); //THIS IS WHERE THE POPULATION ERROR HAPPENS. ERASE 'Name' from '7 Day ADH Data' sheet, retype 'Name' and run function again.

  const workingTimes = parseTimes(getWorkingTimes(targetSheet).flat().filter(time => time !== 'Working Time'));
  const untrackedTimes = parseTimes(getUntrackedTimes(targetSheet).flat().filter(time => time !== 'UNTRACKED'));

  const percentUntrackedArr = untrackedTimes.map(({time, agentIndex},index) => {

    const untrackedTime = time;
    const workingTime = workingTimes[index].time;
     
    return { 
      agent: agents[agentIndex],
      data: {
         agentIndex: agentIndex, 
         percentUntracked : Math.round((untrackedTime / workingTime) * 100) 
      } 
    }
  })
 return percentUntrackedArr;
}
