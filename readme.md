# Data Puller
_Tymeshift Report Email Client and Adherence Report Generator_

## Overview

_PROBLEM:_ "Retrieving, formatting, and generating a Candid adherence report is time consuming and prone to inaccuracy caused by human input error."

_SOLUTION:_ An automated process that:

    1. Retrieves Tymeshift Summary Reports from the user's email inbox
    2. Imports those reports to Google Drive
    3. Culls the data from those reports
    4. Formats that data as a Candid Summary Report
    5. Persists the end-of-month Summary Report for future records

## Project Structure

The Data Puller automation consists of 2 Google App Script projects that work together to accomplish the above solution.  

1. The Tymeshift Email Client (emailClient.js) 
2. The Data Puller (datapuller.js)

### Tymeshift Email Client

The purpose of the email client is to:

1. Identify emails that contain attached Tymeshift Summary Reports (.csv files)
2. Import those reports into a Google Drive repository
3. Rename those reports according to their duration (7 day, 30 day, or Month-to-Date)

#### Email Client Call Stack (emailClient.js)
- ```sendToDrive()``` 

    - This is the orchestrating, event-based function that triggers all fo the other calls
    - This function is triggered by a button click from the custom UI menu in the active spreadsheet (see ui.js)

    - ```getEmailAttachments()```

        - Returns an array of ```EmailAttachments``` from the user's inbox
            - For more info on the ```EmailAttachments``` class see the ```GmailApp``` API Documentation 
        - _These email containing the target attachments is referenced using the 'Tymeshift Reports' label in the user's inbox._

            - ```removeLabel()```

                - This function takes an ```Threads``` class object (see GmailApp API docs) and a gmail thread string as parameters and removes the threads from the users inbox
                - This is done after the threads have been consumed and prevents the user from having more than one email with attachments under the 'Tymeshift Reports' label in their inbox. This is done to prevent importing more than 3 reports at a time. 

    - ```getReportType()```

        - This function accepts a ```ReportDates``` object and returns a ```ReportType```
        - ```ReportType : enum ReportTypes```
            - ```enum ReportTypes { 7 Day = "7 day", ...}```

        - ```parseName()```

            - This function accepts an attachment name (equal to the file name supplied by Tymeshift) and returns a 
            ```ReportDates``` object
                - ```ReportDates = { reportStart : Date, reportEnd: Date}```

    - ```confirmReportsLoaded()```

        - see ```errors.js```

    - ```handleErrors()```

        - Sends UI alert to user if emails are found under the 'Tymeshift Reports' Label

## Data Puller

The purpose of the Data Puller (and related scripts) is to:

1. Locate Tymeshift Reports (.csv) imported from the user's email inbox
2. Extract the data from those .csv files
3. Write that data to a target worksheet within the ADH Reports spreadsheet
4. Perform basic calculations and formatting within the ADH Reports spreadsheet
5. Send formatted data to be persisted in a separate spreadsheet (Historic Reports)
6. Format the data in Historic Reports

## Data Puller Call Stack

- ```onLoad()```

    - Default event that populates the custom UI menu (ui.js) that allows for all other automation controls
    - [If this script does not automatically run when the spreadsheet loads, add a separate onLoad trigger that calls the onLoad function.](https://stackoverflow.com/q/13337599/10718682)

- ```populateSheets()```

    - Orchestrating function for retrieving reports from Google Drive and writing them to the ADH Reports spreadsheet
    - This function is called from a button click in the custom ui menu 
        - 'Sync Reports' option
    
    - ```getReports()```

        - This function returns an array of ```Report``` objects 

            - ```getReportObject```

               - Accepts a fileName string and return a ```Report``` object

                - ```Report = { fileName: string, fileId: string, lastUpdate: Date}```
    
    - ```writeDataToSheet()```

        - This function accepts a ```Report``` object and writes that report's data to the target worksheet in the ADH Reports spreadsheet

            - ```getCSVData()```

                - This function accepts a ```Report.fileID``` and returns that files .csv data as a blob (see the ```DriveApp``` API documentation for more)
                - uses the built in ```Utilities.parseCSV()``` function (see official docs for more)

    - ```confirmReportSync()```

        - This function alerts the user that their reports have synced correctly
        - see ```errors.js```

- ```writeCompleteReport()```

    - This is the main orchestration for the report formatting functionality 
    - Please refer to ```reportFormater.js``` 
    - This function is called from a button click in the custom ui menu 
        - 'Update Summary Report' option 

    - ```writeUntrackedToSheet()```

       - This function gets the last 7 days of untracked data from the corresponding report and writes it to the 'Summary Report' worksheet
       - See ```percentUntracked.js```
    
       - ```getSummaryReport()```

           - returns a ```Sheet``` object referenced by the name 'Summary Report'
           - See ```Sheet``` class in official Google App Script Documentation

        - ```getPercentUntracked()```

            - This function takes in a ```sheetName``` string and returns an array of ```PercentUntracked``` objects

                - ```PercentUntracked = { agent : string, data : Data}```

                - ```Data = { agentIndex : number, percentUntracked: number } ```

    - ```writeGeneralTimesToSheet()```

        - This function gets the last 7 days of % Time Spent in General data from the corresponding report and writes it to the 'Summary Report' worksheet
        - See ```generalTasks.js```

        - ```getTimesInGeneral()```

            - Returns an array representing the % of time spent in general tasks for each active agent
            - Please note the strange spelling of 'GENERAl TASKS' in the heading on each of the Tymeshift Reports. This spelling, with a lowercase 'l' character must be referenced with this strict casing, unless a change is made on Tymeshift. In that case this function should be updated to index the correct spelling. 


