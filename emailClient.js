// sendToDrive() is the main function for the email client and is called from a button click event in the spreadsheet

function sendToDrive(){
  
  const emailAttachments = getEmailAttachments().flat();
  const names = emailAttachments.map(attachment => attachment.getName());

  if(emailAttachments && emailAttachments.length > 0){

    writeReportNamesToSheet(names);
      /*
        writeReportNamesToSheet() writes data to the 'Troubleshooting' spreadsheet. This function provides the 
        original file names for the current reports. Great for making sure your data source is accurate.
      */

    emailAttachments.map((attachment,index) => {

      const reportType = getReportType(parseName(names[index]));
      const file = attachment.copyBlob();
      const folder = DriveApp.getFolderById('1mPmevvful7jgIbBLJ1AT7fwX0e2X0FKS');
      const newReport = folder.createFile(file);

      newReport.setName(`${reportType}.csv`);   
    })
    confirmReportsLoaded()
  }
  else{
     handleErrors({type: 'alert', errorText: `No emails found under the Tymeshift Reports label in your Gmail inbox. Combine all three reports together in an email, send it to yourself, then label it 'Tymeshift Reports', then try this again.`})
  }
}

function getEmailAttachments(){
 
 const gmailLabel = GmailApp.getUserLabelByName('TymeShift Reports');
 const tsReportThreads = gmailLabel.getThreads();
 const tsReportMessages = tsReportThreads.map(thread => thread.getMessages());
 const attachments = tsReportMessages.map(message => message[0].getAttachments());

 tsReportThreads.map(thread => removeLabel(thread, gmailLabel));
 return attachments;
}

function parseName(attachmentName){
  const splitSubject = attachmentName.split('_');
  const startString = `${splitSubject[5]}/${splitSubject[6]}/${splitSubject[7]}:${splitSubject[8]}`;
  const endString = `${splitSubject[10]}/${splitSubject[11]}/${splitSubject[12]}:${splitSubject[13]}`;
  return {
    reportStart: new Date(startString),
    reportEnd : new Date(endString)
  }
}

function testReportType(){
  
  const emailAttachments = getEmailAttachments().flat();
  const names = emailAttachments.map(attachment => attachment.getName());

  emailAttachments.map((attachment,index) => {

      getReportType(parseName(names[index]));
  })
}

function getReportType({reportStart, reportEnd}){
  console.log(`START DATE: ${reportStart}, END TIME: ${reportEnd}`)
  const duration = Math.floor((reportEnd.getTime() - reportStart.getTime()) / (1000*60*60*24));
  const startMonth = reportStart.getMonth();
  const endMonth = reportEnd.getMonth();
  console.log(`DURATION: ${duration}`);
  if(duration === 6){
    console.log('7 day')
    return '7 day'
  }else if(duration > 7 && startMonth === endMonth || duration < 7 && startMonth === endMonth){
    console.log('m2d')
    return 'm2d'
  }else if(duration > 8 && startMonth !== endMonth){
    console.log('30 day')
    return '30 day'
  }
  else{
    console.error(`Can't determine report type. ERROR TRACE: startMonth: ${startMonth}, endMonth: ${endMonth}, duration: ${duration}`);
  }
}

function removeLabel(thread,label){
  thread.removeLabel(label);
}
