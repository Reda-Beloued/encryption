
//var mailboxItem;
var infoMsgKey = "Encryption-Addin-Info-Message";
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

    currentMail.notificationMessages.addAsync(infoMsgKey, {
        type: "informationalMessage",
        message: "Checking settings..",
        persistent: false,
        icon: "about16"
    });

    loadSettings();
    loadKeywords(() => {

        if (settings.autoSendKeywordList !== null && settings.autoSendKeywordList.length > 0) {
            currentMail.subject.getAsync(
                function callback(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        var subject = result.value;
                        if (isTextConatinsKeyword(subject)) {
                            prepareSend(event);
                        } else {
                            currentMail.body.getAsync('text',
                                function callback(result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                        var body = result.value;
                                        if (isTextConatinsKeyword(body)) {
                                            prepareSend(event);
                                        } else {
                                            event.completed({ allowEvent: true });
                                        }
                                    }
                                });
                        }
                    }
                });
        } else {
            event.completed({ allowEvent: true });
        }
    });
}

function prepareSend(event) {
    currentMail.notificationMessages.addAsync(infoMsgKey, {
        type: "informationalMessage",
        message: "Sending message via "+appInfo.name+" encryption addin..",
        persistent: false,
        icon: "about16"
    });

    currentEvent = event;
    isUIless = true;

    prepareData();
    sendMessage();
}
function isTextConatinsKeyword(text) {

    var keywords = settings.autoSendKeywordList.split(",").map(item => item.trim());

    for (var i = 0; i < keywords.length; i++) {

        if (text.indexOf(keywords[i]) >= 0) {
            if(settings.removeKeyword)
                cleanSubject = text.replace(keywords[i], '').trim();
            return true;
        }
    }

    return false;
}