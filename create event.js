var globalVariables = (function () {
    var mainDatabaseSS = SpreadsheetApp.getActive();
    var createRideResponseSheetName = 'Form Responses';
    var responseSheet = mainDatabaseSS.getSheetByName(createRideResponseSheetName);
    var linkedFormUrl = responseSheet.getFormUrl();
    var rideCreateForm = FormApp.openByUrl(linkedFormUrl);
    var signUpForm = FormApp.openById('15WC_GAoeVziEZpbu63cJ_MCfgqoMDs8pBCvDBikatU4');
    var seeWhoSignedUpMacroID = 'AKfycbxnMNlqpRutQ3MrP-ksNhsXuQSX3ZvJXfFjU_BSEjOl';
    var rideSignUpFormID = '1FAIpQLSeXClehvRgnRx2Q4ENRu5OPCRMCwzhmI3s9zEUTYOyHufQ2WA';
    var repeatRideFormID = '1FAIpQLSc2JLZCALMt6k_h7RbEA4y8ub1CiyOu2WAqLwGw9XMWGa0pyQ';
    var postRideFormID = '1FAIpQLSdq8p7hq7La3pV1r1PWQU5wctD8sA5o05vbVnzqDjZp5PRJDQ';
    return {
        rideCreateForm: rideCreateForm,
        mainDatabaseSS: mainDatabaseSS,
        responseSheet: responseSheet,
        signUpForm: signUpForm,
        seeWhoSignedUpMacroID: seeWhoSignedUpMacroID,
        rideSignUpFormID: rideSignUpFormID,
        repeatRideFormID: repeatRideFormID,
        postRideFormID: postRideFormID
    };
}());

function getIDsOnForm(){
    var items = globalVariables.rideCreateForm.getItems();
    for (var i in items) {
        Logger.log(items[i].getTitle() + ': ' + items[i].getId());
    }
}

function testing() {

}

function onSubmit(e) {
    var eventInfo = {
        //these fields are mapped directly from the responses on the form:
        leaderEmail: {QuestionID: 296382551, EditFormID: 457481492, value: ''},
        leaderName: {QuestionID: 1752987048, EditFormID: 617361207, value: ''},
        rideTitle: {QuestionID: 1812650511, EditFormID: 1681098133, value: ''},
        description: {QuestionID: 259055813, EditFormID: 1666226205, value: ''},
        rideLevel: {QuestionID: 1070500107, EditFormID: 1553132989, value: ''},
        rideType: {QuestionID: 500128296, EditFormID: 788554312, value: ''},
        rideLocation: {QuestionID: 814127789, EditFormID: 940217310, value: ''},
        routeLink: {QuestionID: 2128299947, EditFormID: 1120941765, value: ''},
        date: {QuestionID: 932951434, EditFormID: 427532083, value: ''},
        startTime: {QuestionID: 1947966956, EditFormID: 1564705344, value: ''},
        endTime: {QuestionID: 1321142776, EditFormID: 1555197267, value: ''},
        suitableForYouth: {QuestionID: 964539934, EditFormID: 1762405609, value: ''},
        ladiesOnly: {QuestionID: 1930527488, EditFormID: 1115287845, value: ''},
        //these fields are calculated:
        isNewRide: {value: ''},
        rideStart: {value: ''},
        rideEnd: {value: ''},
        editResponseURL: {value: '', spreadsheetCol: 3},
        repeatRideURL: {value: '', spreadsheetCol: 5},
        rideID: {value: '', spreadsheetCol: 1},
        rideSignUpURL: {value: '', spreadsheetCol: 4},
        seeWhoIsSignedUpURL: {value: '', spreadsheetCol: 6},
        postRideFormURL: {value: '', spreadsheetCol: 2}
    };

    //get all the ride info from the form or from the spreadsheet if it's an edited ride
    const eventInfoKeys = Object.keys(eventInfo);
    Logger.log('____form response received____');
    Logger.log(JSON.stringify(e));

    for (var i = 0; i < eventInfoKeys.length; i++) {
        var eventInfoKey = eventInfoKeys[i];
        if (eventInfo[eventInfoKey].hasOwnProperty('QuestionID')) {
            var questionTitle = globalVariables.rideCreateForm.getItemById(eventInfo[eventInfoKey].QuestionID).getTitle();
            eventInfo[eventInfoKey].value = e.namedValues[questionTitle].toString();
            if (eventInfo[eventInfoKey].value === '') {
                eventInfo[eventInfoKey].value = getResponseItemFromSheet(questionTitle, e.range.getRow());
            }
            Logger.log(eventInfoKey + ': ' + eventInfo[eventInfoKey].value)
        }
    }

    eventInfo = getRideIdIfAlreadyExists(e, eventInfo);
    Logger.log('ride ID: ' + eventInfo.rideID.value);
    Logger.log('is new ride: ' + eventInfo.isNewRide.value);

    //get the ride start and ride end by joining the date and start time.
    eventInfo.rideStart.value = parseDateAndTime(eventInfo.date.value, eventInfo.startTime.value);
    eventInfo.rideEnd.value = parseDateAndTime(eventInfo.date.value, eventInfo.endTime.value);

    //get the URL to edit the response and add it to the sheet
    eventInfo.editResponseURL.value = getResponseEditURL(e.namedValues['Timestamp'].toString());
    addExtraFieldToSheet(e, eventInfo, 'editResponseURL');

    //get the  URL to repeat this ride and add it to the sheet
    eventInfo.repeatRideURL.value = getRepeatRideURL(eventInfo);
    addExtraFieldToSheet(e, eventInfo, 'repeatRideURL');

    if (eventInfo.isNewRide.value) {
        Logger.log('____creating calendar appointment and generating the rideID____');
        eventInfo.rideID.value = createCalendarEvent(eventInfo);
        addExtraFieldToSheet(e, eventInfo, 'rideID');
    } else {
        Logger.log('____updating calendar appointment____');
        updateCalendarAppointment(eventInfo);
    }

    //now we've got the ride ID we can get the 'post-ride form' URL, the 'ride Sign up' URL, the 'See who's signed up' URL
    eventInfo.postRideFormURL.value = getPreFilledURLforPostRideForm(eventInfo.rideID.value);
    eventInfo.rideSignUpURL.value = getRideSignUpURL(eventInfo);
    eventInfo.seeWhoIsSignedUpURL.value = getSeeWhoSignedUpURL(eventInfo.rideID.value);

    if (eventInfo.isNewRide.value) { //these need adding to the sheet if it's a new ride
        addExtraFieldToSheet(e, eventInfo, 'postRideFormURL');
        addExtraFieldToSheet(e, eventInfo, 'rideSignUpURL');
        addExtraFieldToSheet(e, eventInfo, 'seeWhoIsSignedUpURL');
        Logger.log('____Signing up ride leader____');
        signUpRideLeader(eventInfo);
    }

    Logger.log('____Adding details of the event to the calendar item____');
    createCalendarEventDetails(eventInfo);

    Logger.log('____Emailing ride Leader____');
    emailRideLeader(eventInfo)
}

function getResponseItemFromSheet(questionTitle, responseRowNo) {
    Logger.log("couldn't find '" + questionTitle + "' in response from form, so getting it from sheet");
    const titleRow = globalVariables.responseSheet.getRange(1, 1, 1, globalVariables.responseSheet.getLastColumn()).getValues();
    const responseRow = globalVariables.responseSheet.getRange(responseRowNo, 1, 1, globalVariables.responseSheet.getLastColumn()).getValues();
    for (var i = 0; i < titleRow[0].length; i++) {
        if (titleRow[0][i] === questionTitle) {
            var responseFromSheet = responseRow[0][i];
            //hack to get the date and time of the ride back as string (like you get from the form) rather than as javascript date object
            if (responseFromSheet instanceof Date) {
                if (questionTitle.indexOf('Date') > 0) {
                    //found the date - now need to convert it to the right format
                    Logger.log(responseFromSheet.toLocaleString());
                    var year = responseFromSheet.getFullYear().toString();
                    Logger.log(year);
                    var month = responseFromSheet.getMonth() + 1;
                    Logger.log(month);
                    month = ((month > 9) ? "" + month : "0" + month);
                    Logger.log(month);
                    var day = responseFromSheet.getDate();
                    Logger.log(day);
                    day = ((day > 9) ? "" + day : "0" + day);
                    Logger.log(day);

                    return day + '/' + month + '/' + year;
                }
                if (questionTitle.indexOf('Time') > 0) {
                    //found the start time or end time - now need to convert it to the right format
                    var hour = responseFromSheet.getHours();
                    hour = ((hour > 9) ? "" + hour : "0" + hour);
                    var minutes = responseFromSheet.getMinutes();
                    minutes = ((minutes > 9) ? "" + minutes : "0" + minutes);
                    return hour + ":" + minutes + ':00';
                }
            }
            return responseRow[0][i].toString();
        }
    }
    return ''
}

function emailRideLeader(eventInfo) {
    var messageArray = [];
    var firstName = eventInfo.leaderName.value.split(" ")[0];
    messageArray.push('<h3> Your ride has been successfully ' + ((eventInfo.isNewRide.value) ? 'created' : 'updated') + '</h3>');
    messageArray.push(firstName + ', thanks for creating this ride for the VCGH community.  Please check that the details below are correct and edit the ride if necessary.');
    messageArray.push("If you've made a mistake then " + createHyperlinkString('edit the ride details', eventInfo.editResponseURL.value) + '.');
    //messageArray.push("You can " + createHyperlinkString('cancel ride completely',eventInfo.editResponseURL) + ' (which will let those signed up know).')
    //todo: make cancelling the ride a web app job
    messageArray.push("After the ride, please don't forget to complete the " + createHyperlinkString('post ride form', eventInfo.postRideFormURL.value) + '.');
    messageArray.push("You may want to quickly " + createHyperlinkString('re-create this ride', eventInfo.repeatRideURL.value) + ' on a different date?');
    messageArray.push("You've been signed up automatically but to see who else has signed up: " + createHyperlinkString("Click here to see who's signed up", eventInfo.seeWhoIsSignedUpURL.value) + '.');
    messageArray.push("<strong>Event Details:</strong>");
    messageArray.push("Title: " + eventInfo.rideTitle.value);
    messageArray.push("Description: " + eventInfo.description.value);
    messageArray.push("Start: " + eventInfo.rideStart.value.toLocaleString());
    messageArray.push("End: " + eventInfo.rideEnd.value.toLocaleString());
    messageArray.push("Location: " + eventInfo.rideLocation.value);
    messageArray.push("Ride Level: " + eventInfo.rideLevel.value);
    messageArray.push("Event Type: " + eventInfo.rideType.value);
    messageArray.push('<img src="http://www.vcgh.co.uk/_/rsrc/1515424568089/config/customLogo.gif?revision=7" alt="VCGH">');

    var subject = ((eventInfo.isNewRide.value) ? 'VCGH Ride Created' : 'VCGH Ride Updated');

    subject = subject + ' (' + (eventInfo.date.value) + ')'
    sendEmail(eventInfo.leaderEmail.value, subject, messageArray)
}

function sendEmail(emailAddress, subject, messageArray) {
    var html = '<body>';
    for (var i = 0; i < messageArray.length; i++) {
        html = html + '<p>' + messageArray[i].toString() + '</p>'
    }
    html = html + '</body>';
    MailApp.sendEmail(emailAddress, subject, 'HTML message content hidden', {htmlBody: html});
}

function getRideIdIfAlreadyExists(e, eventInfo) { //check if form submission is editing a ride rather than creating one
    Logger.log('____checking if rideId already exists____');
    const responseRowNo = e.range.getRow();
    const rideIdCell = globalVariables.responseSheet.getRange(responseRowNo, 1).getCell(1, 1);
    if (rideIdCell.isBlank()) {
        Logger.log('ride ID not found so must be a new ride');
        eventInfo.isNewRide.value = true;
    } else {
        Logger.log('ride ID found so must be an edited ride');
        eventInfo.isNewRide.value = false;
        eventInfo.rideID.value = rideIdCell.getValue();
    }
    return eventInfo
}

function getLinkedCalendar() {

}

function createCalendarEvent(eventInfo) {
    const linkedCalendar = CalendarApp.getDefaultCalendar();
    const event = linkedCalendar.createEvent(eventInfo.rideTitle.value, eventInfo.rideStart.value, eventInfo.rideEnd.value);
    const eventID = event.getId();
    return eventID.slice(0, eventID.indexOf("@"));
}

function createCalendarEventDetails(eventInfo) {
    const linkedCalendar = CalendarApp.getDefaultCalendar();
    const eventID = eventInfo.rideID.value + '@google.com';
    const event = linkedCalendar.getEventById(eventID);
    var descriptionArray = [];
    descriptionArray.push(eventInfo.description.value);
    //TODO: descriptionArray.push(getRideLevelDescription())
    descriptionArray.push(makeBold("Sign Up:") + '\n' + createHyperlinkString('Click here to sign up to this ride', eventInfo.rideSignUpURL.value));
    descriptionArray.push(makeBold("See who's signed up:") + '\n' + createHyperlinkString("Click here to see who's signed up", eventInfo.seeWhoIsSignedUpURL.value));
    if (!(eventInfo.routeLink.value === '')) {
        descriptionArray.push(makeBold("Route:") + '\n' + createHyperlinkString('View the proposed route', eventInfo.routeLink.value))
    }
    descriptionArray.push('Please email the ride leader ' + eventInfo.leaderName.value + '(' + eventInfo.leaderEmail.value + ') with any queries.');
    if (!eventInfo.isNewRide.value) {
        descriptionArray.push('(these Ride details have been edited from the original)')
    }
    var eventDescription = '';
    for (var i = 0; i < descriptionArray.length; i++) {
        eventDescription = eventDescription + descriptionArray[i].toString() + '\n\n'
    }
    event.setDescription(eventDescription);

    event.setLocation(eventInfo.rideLocation.value)
}

function updateCalendarAppointment(eventInfo) {
    const linkedCalendar = CalendarApp.getDefaultCalendar();
    Logger.log('RideID: ' + eventInfo.rideID.value)
    const eventID = eventInfo.rideID.value + '@google.com';
    Logger.log('event ID: ' + eventID);
    const event = linkedCalendar.getEventById(eventID);
    Logger.log('event: ' + event);
    event.setTitle(eventInfo.rideTitle.value + ' (edited)');
    event.setTime(eventInfo.rideStart.value, eventInfo.rideEnd.value)
}

function getPreFilledURLforPostRideForm(rideID) {
    return 'https://docs.google.com/forms/d/e/' + globalVariables.postRideFormID + '/viewform?usp=pp_url&entry.606826787=' + rideID
}

function signUpRideLeader(eventInfo) {

    var newData = [{
        index: 0,
        getAs: 'asTextItem',
        value: eventInfo.rideID.value,
    }, {
        index: 1,
        getAs: 'asTextItem',
        value: eventInfo.rideStart.value,
    }, {
        index: 2,
        getAs: 'asTextItem',
        value: eventInfo.rideTitle.value,
    }, {
        index: 3,
        getAs: 'asMultipleChoiceItem',
        value: 'Yes',
    }, {
        index: 4,
        getAs: 'asTextItem',
        value: eventInfo.leaderEmail.value,
    }, {
        index: 5,
        getAs: 'asTextItem',
        value: eventInfo.leaderName.value + ' (ride leader)',
    }];

    var formResponse = globalVariables.signUpForm.createResponse();
    var items,
        formItem,
        response;

    items = globalVariables.signUpForm.getItems();
    for (var i = 0; i < newData.length; i++) {
        formItem = items[newData[i].index][newData[i].getAs]();
        response = formItem.createResponse(newData[i].value);
        formResponse.withItemResponse(response);
    }

    formResponse.submit();
}

function getResponseEditURL(timestamp) {
    eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());

    var responses = globalVariables.rideCreateForm.getResponses().reverse();
    var thisResponsesMoment = moment(timestamp, "DD/MM/YYYY HH:mm:ss", 'en');
    var thisResponsesUnixSec = parseInt(thisResponsesMoment.format("X"));
    for (var i = 0; i < responses.length; i++) {
        var responseMoment = moment(responses[i].getTimestamp());
        var responseMomentUnixSec = parseInt(responseMoment.format("X"));
        if (responseMomentUnixSec === thisResponsesUnixSec || responseMomentUnixSec === thisResponsesUnixSec - 1) {
            return responses[i].getEditResponseUrl()
        }
    }
}

function getRepeatRideURL(eventInfo) {
    var URL = 'https://docs.google.com/forms/d/e/' + globalVariables.repeatRideFormID + '/viewform?usp=pp_url';
    const eventInfoKeys = Object.keys(eventInfo);
    for (var i = 0; i < eventInfoKeys.length; i++) {
        var eventInfoKey = eventInfoKeys[i];
        if (eventInfo[eventInfoKey].hasOwnProperty('EditFormID')) {
            var URLToAdd = '&entry.ID=URI'.replace('ID', eventInfo[eventInfoKey].EditFormID).replace('URI', encodeURI(eventInfo[eventInfoKey].value));
            URL = URL + URLToAdd
        }
    }
    return URL
}

function getRideSignUpURL(eventInfo) {
    var prefilledURL = 'https://docs.google.com/forms/d/e/' + globalVariables.rideSignUpFormID + '/viewform?usp=pp_url&entry.853358763=rideID&entry.435466124=rideDate&entry.1250084608=rideTitle';
    prefilledURL = prefilledURL.replace('rideID', encodeURI(eventInfo.rideID.value));
    prefilledURL = prefilledURL.replace('rideTitle', encodeURI(eventInfo.rideTitle.value));
    prefilledURL = prefilledURL.replace('rideDate', encodeURI(eventInfo.rideStart.value));
    return prefilledURL
}

/*
function getRideLevelDescription() {
    Logger.log('Getting Ride Speed')
    var rideLevelsSheet = mainDatabaseSS.getSheetByName('rideLevels')
    var rideLevelsData = rideLevelsSheet.getDataRange();
    for (var i = 1; i <= rideLevelsData.getNumRows(); i++) {
        if (rideLevelsData.getCell(i, 1).getValue() == eventInfo.rideLevel) {
            var rideSpeedDescription = makeBold('Ride Level:') + '\n' + eventInfo.rideLevel + ' ride at ' + rideLevelsData.getCell(i, 4).getValue() + 'mph for ' + rideLevelsData.getCell(i, 5).getValue() + 'miles [' + rideLevelsData.getCell(i, 2).getValue() + 'km/h for ' + rideLevelsData.getCell(i, 3).getValue() + 'km] <a href=' + '"' + rideLevelsData.getCell(i, 6).getValue() + '"' + '>Click Here for more info</a>'
            Logger.log(rideSpeedDescription)
            return rideSpeedDescription
        }
    }
}
*/

function getSeeWhoSignedUpURL(rideID) {
    const appURL = 'https://script.google.com/a/vcgh.co.uk/macros/s/' + globalVariables.seeWhoSignedUpMacroID + '/dev?rideID=';
    return appURL + rideID
}

function createHyperlinkString(title, link) {
    return '<a href="url">linkText</a>'.replace('url', link).replace('linkText', title)
}

function makeBold(input) {
    return '<strong>' + input + '</strong>'
}

function addExtraFieldToSheet(e, eventInfo, eventInfoKey) {
    const responseRow = e.range.getRow();
    const spreadsheetColumn = eventInfo[eventInfoKey].spreadsheetCol;
    const valueToSet = eventInfo[eventInfoKey].value;
    globalVariables.responseSheet.getRange(responseRow, spreadsheetColumn).setValue(valueToSet);
}

function parseDateAndTime(dateString, timeString) {
    Logger.log('____Parsing date and time_____');
    const day = dateString.split('/')[0];
    Logger.log('day: ' + day);
    const month = dateString.split('/')[1];
    Logger.log('month: ' + month);
    const year = dateString.split('/')[2];
    Logger.log('year: ' + year);
    const hours = timeString.split(':')[0];
    Logger.log('hour: ' + hours)
    const minutes = timeString.split(':')[1];
    Logger.log('minutes: ' + minutes);
    const dateForReturn = new Date(year, month - 1, day, hours, minutes);
    Logger.log('full date: ' + dateForReturn.toISOString())
    return dateForReturn
}
