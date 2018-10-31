var globalVariables = (function () {
    var signupSpreadsheet = SpreadsheetApp.getActive();
    var signupSheetName = 'Form Responses';
    var responseSheet = signupSpreadsheet.getSheetByName(signupSheetName);
    var linkedFormUrl = responseSheet.getFormUrl();
    var signupForm = FormApp.openByUrl(linkedFormUrl);
    return {
        signupForm: signupForm,
        signupSpreadsheet: signupSpreadsheet,
        responseSheet: responseSheet,
    };
}());

function onSubmit(e) {
    var eventInfo = {
        //these fields are mapped directly from the responses on the form:
        rideID: {QuestionID: 296382551, value: ''},
        VCGHMember: {QuestionID: 1752987048, value: ''},
        emailAddress: {QuestionID: 1812650511, value: ''},
        name: {QuestionID: 1812650511, value: ''}
    };

    //get all the ride info from the form:
    const eventInfoKeys = Object.keys(eventInfo);
    Logger.log('____form response received____');
    Logger.log(JSON.stringify(e));

    for (var i = 0; i < eventInfoKeys.length; i++) {
        var eventInfoKey = eventInfoKeys[i];
        if (eventInfo[eventInfoKey].hasOwnProperty('QuestionID')) {
            var questionTitle = globalVariables.signupForm.getItemById(eventInfo[eventInfoKey].QuestionID).getTitle();
            eventInfo[eventInfoKey].value = e.namedValues[questionTitle].toString();
            Logger.log(eventInfoKey + ': ' + eventInfo[eventInfoKey].value)
        }
    }

    eventInfo = getRideIdIfAlreadyExists(e, eventInfo);
    Logger.log('ride ID: ' + eventInfo.rideID.value);
    Logger.log('is new ride: ' + eventInfo.isNewRide.value);
}