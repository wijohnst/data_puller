function getPercentUntracked(){
  let sheet = SpreadsheetApp.getActive();
  let targetSheet = '7 Day ADH Data';

  let agents;
  const searchParam = 'Name';

  try{
    agents = getColByHeading(targetSheet,searchParam).flat().filter(name => name !== searchParam && name !== '')
  }
  catch(err){
    handleErrors({type: 'alert', errorText: getColHeadingErrorText('Name',targetSheet, "percentUntracked.gs >> getPercentUntracked()")})
  }
  finally{
    handleGetColSuccess(searchParam,targetSheet);
  }


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
