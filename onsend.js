
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

    currentEvent = event;
    const queryString = window.location.search;

    
    loadSettings();

    if (settings === null || settings ===undefined ||
        settings.useAutoSend === false ) {

        currentEvent.completed({ allowEvent: true });
        
    } else {

        if (appInfo.configUrl === null) {
            startSendingMessage();
        } else {

            currentMail.notificationMessages.addAsync(infoMsgKey, {
                type: "informationalMessage",
                message: "Checking settings..",
                persistent: false,
                icon: "about16"
            });

            loadKeywords(() => {
                startSendingMessage();
            });
        }
    }
}

var keywordsText = "";
var keywords;
function startSendingMessage() {
   
    if (settings.autoSendKeywordList !== undefined &&
        settings.autoSendKeywordList !== null &&
        settings.autoSendKeywordList.length > 0) {
        keywordsText = settings.autoSendKeywordList;
    }
    if (settings.customSendKeywordList !==undefined &&
        settings.customSendKeywordList !== null &&
        settings.customSendKeywordList.length > 0) {
        if (keywordsText.length > 0)
            keywordsText += "," + settings.customSendKeywordList;
        else
            keywordsText = settings.customSendKeywordList;
    }

    if (keywordsText.trim().length > 0) {
        keywords = keywordsText.split(",").map(item => item.trim());
        currentMail.subject.getAsync(
            function callback(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var subject = result.value;
                    if (isTextConatinsKeyword(subject)) {
                        prepareSend();
                    } else {
                        currentMail.body.getAsync('text',
                            function callback(result) {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    var body = result.value;
                                    if (isTextConatinsKeyword(body)) {
                                        cleanSubject = undefined;
                                        prepareSend(currentEvent);
                                    } else {
                                        currentEvent.completed({ allowEvent: true });
                                    }
                                }
                            });
                    }
                }
            });
    } else {
        currentEvent.completed({ allowEvent: true });
    }
}
function prepareSend() {
    currentMail.notificationMessages.addAsync(infoMsgKey, {
        type: "informationalMessage",
        message: "Sending message via "+appInfo.name+" encryption addin..",
        persistent: false,
        icon: "about16"
    });

   
    isUIless = true;

    prepareData();
    sendMessage();
}
function isTextConatinsKeyword(text) {

    for (var i = 0; i < keywords.length; i++) {
        if (text.indexOf(keywords[i]) >= 0) {
            return true;
        } else {//check if uppercase exists as well..
            var cap = keywords[i].charAt(0).toUpperCase() + keywords[i].slice(1);
            if (text.indexOf(cap) >= 0) {
                return true;
            }
        }
    }

    return false;
}