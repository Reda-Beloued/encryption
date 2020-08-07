
//var mailboxItem;
var infoMsgKey = "SSL365-Info-Message";
var currentEvent;

Office.initialize = function (reason) {
    currentMail = Office.context.mailbox.item;
};

function sendUsingSSL365(event) {

    if (currentMail === undefined) {
        event.completed({ allowEvent: true });
        return;
    }

    const queryString = window.location.search;

    loadSettings();

    if (settings.autoSendKeywordList !== null && settings.autoSendKeywordList.length > 0) {
        currentMail.subject.getAsync(
            function callback(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var subject = result.value;
                    if (isSubjectConatinsKeyword(subject)) {

                        currentMail.notificationMessages.addAsync(infoMsgKey, {
                            type: "informationalMessage",
                            message: "Sending message via SSL365..",
                            persistent: false,
                            icon: "about16"
                        });

                        currentEvent = event;
                        isUIless = true;

                        prepareData();
                        sendMessage();
                    } else {
                        event.completed({ allowEvent: true });
                    }
                }
            });
    } else {
        event.completed({ allowEvent: true });
    }
}

function isSubjectConatinsKeyword(subject) {

    var keywords = settings.autoSendKeywordList.split(",").map(item => item.trim());

    for (var i = 0; i < keywords.length; i++) {

        if (subject.indexOf(keywords[i]) >= 0) {
            if(settings.removeKeyword)
                cleanSubject = subject.replace(keywords[i], '').trim();
            return true;
        }
    }

    return false;
}