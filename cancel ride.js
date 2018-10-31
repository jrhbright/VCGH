var globalVariables = (function () {
    var cancelSpreadsheet = SpreadsheetApp.getActive();
    var cancelRideSheetName = 'Form Responses';
    var responseSheet = cancelSpreadsheet.getSheetByName(cancelRideSheetName);
    var linkedFormUrl = responseSheet.getFormUrl();
    var rideCancelForm = FormApp.openByUrl(linkedFormUrl);
    return {
        rideCancelForm: rideCancelForm,
        cancelSpreadsheet: cancelSpreadsheet,
        responseSheet: responseSheet,
    };
}());

function onSubmit(e) {
    var eventInfo = {
        //these fields are mapped directly from the responses on the form:
        rideID: {QuestionID: 296382551, value: ''},
        cancellationReason: {QuestionID: 1752987048, value: ''},
        emailThoseSignedUp: {QuestionID: 1812650511, value: ''}
        //these fields are calculated:
    };

    //get all the ride info from the form:
    const eventInfoKeys = Object.keys(eventInfo);
    Logger.log('____form response received____');
    Logger.log(JSON.stringify(e));

    for (var i = 0; i < eventInfoKeys.length; i++) {
        var eventInfoKey = eventInfoKeys[i];
        if (eventInfo[eventInfoKey].hasOwnProperty('QuestionID')) {
            var questionTitle = globalVariables.rideCancelForm.getItemById(eventInfo[eventInfoKey].QuestionID).getTitle();
            eventInfo[eventInfoKey].value = e.namedValues[questionTitle].toString();
            Logger.log(eventInfoKey + ': ' + eventInfo[eventInfoKey].value)
        }
    }

    eventInfo = getRideIdIfAlreadyExists(e, eventInfo);
    Logger.log('ride ID: ' + eventInfo.rideID.value);
    Logger.log('is new ride: ' + eventInfo.isNewRide.value);
}

function cancelEvent(RideID){
 //hello git
}

function cancelCalendarEvent(){

}

function emailThoseSignedUp(){

}