var rideCreateForm = FormApp.openById('1IkvrC1L4sDu1Xp7SpDSGF02_Fs526QKNLjBvUmyD5Eo');
var responseSpreadsheet = SpreadsheetApp.openById('17FWr4QvI1nC1N9MgDZhhE7ejZZFpfBuNYr_eKNr0jMw')

function doGet(e) {
    Logger.log(JSON.stringify(e));
    const rideID = e.parameter.rideID;
    Logger.log(rideID);

    const rideDetails = getRideDetails(rideID);
    Logger.log(JSON.stringify(rideDetails));

    const filteredData = getListofThoseSignedUp(rideID);

    var template =  HtmlService.createTemplateFromFile('index');
    template.rideTitle = " " + rideDetails.eventTitle.value;
    const JSeventDate = new Date(rideDetails.eventDate.value);
    template.rideDate = JSeventDate.toLocaleDateString();
    const JSEventTime = new Date(rideDetails.eventTime.value);
    template.rideTime = JSEventTime.toLocaleTimeString().substring(0,5);
    template.data = filteredData;
    var html = template.evaluate();
    return html;
}

function test(){
    var rideID = 'kl9vihqut7ee0c50tam3bafc5c';
    Logger.log(rideID);

    var rideDetails = getRideDetails(rideID);
    Logger.log(JSON.stringify(rideDetails))
}

function getRideDetails(rideID){
    var rideInfo = {
        eventTitle: {questionID: 1812650511, value: ''},
        eventDate: {questionID: 932951434, value: ''},
        eventTime: {questionID: 1947966956, value: ''},
    };
    const sheetRow = findRowOnResponseSheet(rideID);
    const rideInfoKeys = Object.keys(rideInfo);
    for (var i = 0; i < rideInfoKeys.length; i++) {
        var rideInfoKey = rideInfoKeys[i];
        if (rideInfo[rideInfoKey].hasOwnProperty('questionID')) {
            var sheetColumn = getColumnOnResponseSheetForQuestion(rideInfo[rideInfoKey].questionID);
            rideInfo[rideInfoKey].value=responseSpreadsheet.getActiveSheet().getRange(sheetRow,sheetColumn).getValue()
        }
    }
    return rideInfo;
}

function getColumnOnResponseSheetForQuestion(QuestionID){
    var questionTitle = rideCreateForm.getItemById(QuestionID).getTitle();
    const titleRow = responseSpreadsheet.getActiveSheet().getRange(1, 1, 1, responseSpreadsheet.getActiveSheet().getLastColumn()).getValues();
    for (var i = 0; i < titleRow[0].length; i++) {
        if (titleRow[0][i] === questionTitle) {
            return i+1;
        }
    }
}

function findRowOnResponseSheet(rideID){
    const allRides = responseSpreadsheet.getActiveSheet().getDataRange().getValues()
    for (var i = 1; i < allRides.length; i++) {
        if (allRides[i][0] == rideID){
            return i+1;
        }
    }
}
function getListofThoseSignedUp(rideID){
    var rawData = SpreadsheetApp
        .openById('1ywrjto0vUK5eho8lbycsggraag4VwfhWCrZaUlofCow')
        .getActiveSheet()
        .getDataRange()
        .getValues();
    var filteredRows = [['Name','Signed Up On','Member?']]
    for (var i = 1; i < rawData.length; i++) {
        Logger.log("is " + rawData[i][1] + " equal to " + rideID + "?")
        if (rawData[i][1]==rideID){
            var signedUpRow = []
            signedUpRow.push(rawData[i][3])
            const JSSignedUpDate = new Date(rawData[i][0])
            signedUpRow.push(JSSignedUpDate.toLocaleString())
            signedUpRow.push(rawData[i][6])
            filteredRows.push(signedUpRow)
        }
    }
    return filteredRows
}