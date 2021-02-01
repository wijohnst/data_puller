/*

    onLoad() -> // One of App Scripts default triggers. A function (in this case, populateSheets()) is called
                   whenever the spreadsheet is opened

        populateSheets() ->

            getReports() // returns an array of File Objects (custom type)
                interface FileObj = {
                    fileName: string, // File name in Google drive + file extension 
                    fileId: string, // Unique Id 
                    lastUpdate: ISO date
                }

                getReportObject(name) -> //Receives one of 3 report name choices and returns a file 
                    
                    This function uses the DriveApp API to generate a 'file iterator' (JS Generator), 
                    which is a list of all files in a Drive account that match a criteria, in this case
                    file name

                    The files are compared (eg: multiple files name '7 Day.csv'...) and the newest file
                    is returned
        
        Once the reports array has been generated we map over them and pass each report to the 
        following callback:
                
                writeDataToSheet(report) //takes a single File Object as a param
                    
                    First we get the data from the individual report by calling:
                    
                        getCsvData(report.fileID)

                            Again using the DriveApp API, but this time the target is a specific file,
                            not a list of files, using the fileID from the File Object as a reference

                            The function uses a built in utility function 'parseCsv' to prepare the data
                            to be written. This parsed CSV object is returned
                    
                    We also use the sheetName() switch case to find the specific spreadsheet we want to
                    write out data to. 

                    A note on the sheetMax constant:

                        This returns a Range (class.Range) that must match the dimensions of our parsed 
                        CSV object. So when we instantiate the range (sheet.getRange()), the 3rd and 4th
                        param are references to that CSV object's length.

                        We do this because the next method, setValues must be called on a range with the 
                        same dimensions as the data to be written

                    Finally we call setValues on our range and write out CSV data to the target sheet

*/
function test(){
  Tymeshift_Summary_Email_Client.sendToDrive();
}
function populateSheets(){

    let reports = getReports();

    reports.map(report => writeDataToSheet(report));
    confirmReportsSync();
}

function getReports(){

    const reportNames = ['7 day.csv', '30 day.csv', 'm2d.csv']; //Valid reports must be named one of these choices (basically an ENUM)

    const reports = reportNames.map(name => getReportObject(name));

    return reports;
}

function getReportObject(fileName){

    const files = DriveApp.getFilesByName(fileName);

    const fileObjs = [];

    while(files.hasNext()){
        const file = files.next();
        const fileObj = {
            fileName : file.getName(),
            fileID : file.getId(),
            lastUpdate : file.getLastUpdated()
        }
        fileObjs.push(fileObj);
    }

    const newestFile = fileObjs.reduce((a,b) => {
        return new Date(a.lastUpdate) > new Date(b.lastUpdate) ? a : b
    });

    return newestFile;
}

function writeDataToSheet(report){ 

    let data = getCsvData(report.fileID);

    const sheetName = ({fileName}) => {
        switch (fileName) {
            case '7 day.csv':
                return '7 Day ADH Data'
            case '30 day.csv':
                return '30 Day ADH Data'
            case 'm2d.csv':
                return 'Month-to-Date ADH Data'
        }
    }
 
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName(report));
    sheet.clearContents(); // 1/22/21 -> Throwing an error on execution...
    let sheetMax = sheet.getRange(1,1,data.length, data[0].length);
    let sheetMin = sheet.getRange(1,1);
    
    sheetMax.setValues(data);
    sheetMin.setValue('Name');
}

function getCsvData(fileId){

    const file = DriveApp.getFileById(fileId);

    return Utilities.parseCsv(file.getBlob().getDataAsString());
}






