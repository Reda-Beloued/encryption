
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

    var txt = text.toLowerCase();
    if (keywords.find(word => word === "*")) {
        return true;
    } else {
        var asteriks = keywords.filter(word => word.endsWith("*"));
        var found=false;

        asteriks.every((asterik) => {
            var asterikFree = asterik.substr(0, asterik.length - 1).toLowerCase();
            if (txt.indexOf(asterikFree) === 0 || txt.indexOf(" " + asterikFree) > 0) {
                found=true;
                return false;
            }
            
            return true;
        });
        
        if(found){
            return true;
        }
    }

    for (var i = 0; i < keywords.length; i++) {
        var keyword = keywords[i];
        if (keyword.length > 0 && keyword !== " ") {
            if (text.indexOf(keyword) >= 0) {
                return true;
            } else {//check if uppercase exists as well..
                var cap = keyword.charAt(0).toUpperCase() + keyword.slice(1);
                if (text.indexOf(cap) >= 0) {
                    return true;
                }
            }
        }
    }
 
    return false;
}
