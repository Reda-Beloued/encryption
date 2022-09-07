
const OUTLOOK_WEB_APP = "OutlookWebApp";
const OK_TEXT = "ok";
const WARNING_TEXT = "warning";
const WARNING_OK_TEXT = "warning-ok";
const ERROR_TEXT = "error";
const MAX_ATTEMPTS = 5;

const SENT_MAIL_FLAG =
    [
        {
            PropertyId: 'Integer 0x0E07',
            Value: '1'
        }
    ];

var cleanSubject = undefined;
var encryptedMessage = null;

var currentMail;
var newMail =
{
    Sender: null,
    From: null,
    Subject: null,
    Body: null,
    ToRecipients: null,
    CcRecipients: null,
    SingleValueExtendedProperties: null,
    Attachments: null
};
var accessToken;
var currentMailID;
var attachements = [];
var restHost;// = Office.context.mailbox.restUrl;
var formData;

var isUIless = false;

function prepareData() {

    restHost = Office.context.mailbox.restUrl;
    currentMail = Office.context.mailbox.item;
    currentMailID = undefined;

    currentMail.saveAsync(function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            //
            currentMailID = getItemRestId(result.value);
        }
    });

    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            accessToken = result.value;
        }
    });
}

function closeTaskpane() {
    Office.context.ui.closeContainer();
}

function sendMessage() {

    if (settings === null || settings.username === undefined || settings.username.length === 0
        || settings.password === undefined || settings.password.length === 0) {

        showNotification(WARNING_TEXT, "Credentials Missing!",
            "Please use 'Settings' to enter your credentials then try again");

        return;
    }

    var to, subject, body;

    formData = new FormData();
    formData.append("email", settings.username);
    formData.append("password", settings.password);
    formData.append("send_mode", "L");
    //formData.append("passwordless", "0");
    formData.append("download_all", "1");

    //if (appInfo.id === "helloflex") {
    //    formData.append("passwordless", "1");
    //}

     if (settings.requirePassword === true) {
        formData.append("passwordless", "0");
    } else {
        formData.append("passwordless", "1");
    }

    currentMail.to.getAsync(
        function callback(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {

                if (result.value.length === 0 || (to = getEmailAddresses(result.value)) === "") {
                    showNotification(WARNING_TEXT, "Recipient Missing!",
                        "Please, make sure you have provided a valid recipient email in 'To' field.");
                    return;
                }

                formData.append("to", to);

                currentMail.subject.getAsync(
                    function callback(result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            subject = result.value;
                            if (subject.length === 0) {
                                showNotification(WARNING_TEXT, "Subject Missing!",
                                    "Please, make sure you have provided a subject for the email.");
                                return;
                            }
                            formData.append("subject", subject);

                            currentMail.body.getAsync("text",
                                function callback(result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {

                                        body = result.value.trim();
                                        if (body.length === 0) {
                                            showNotification(WARNING_TEXT, "Empty Message!",
                                                "Please, make sure you have provided some text as content for the email.");
                                            return;
                                        }

                                        formData.append("unencrypted_body", body);
                                        
                                        if (encryptedMessage !== null &&
                                            encryptedMessage.length > 0) {
                                            formData.append("body", encryptedMessage);
                                        }

                                        var options = { asyncContext: { currentItem: currentMail } };
                                        currentMail.getAttachmentsAsync(options, function callback(result) {
                                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                                var attchs = result.value;
                                                var attchCount = attchs.length;
                                                if (attchCount > 0) {
                                                    showSpinner(true);
                                                    showSpinnerText("Preparing attachments..");
                                                    var loadedAttchCount = 0;
                                                    var attch;
                                                    var fileNames = "";
                                                    for (var i = 0; i < attchCount; i++) {
                                                        attch = attchs[i];
                                                        var fileName = attch.name;


                                                        result.asyncContext.currentItem.getAttachmentContentAsync(attch.id, { asyncContext: { currentItem: fileName } }, function callback(result) {
                                                            if (result.status === Office.AsyncResultStatus.Succeeded) {

                                                                var fileName = result.asyncContext.currentItem;
                                                                loadedAttchCount++;
                                                                showSpinnerText("Loading attachment.. (" + loadedAttchCount + "/" + attchCount + "): " + fileName);
                                                                var file = result.value;
                                                                if (file.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {

                                                                    attachements.push({
                                                                        "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                                                                        Name: fileName,
                                                                        ContentBytes: file.content
                                                                    });

                                                                    var buffer = decodeBase64(file.content);
                                                                    var blob = new Blob([buffer], { type: "octet-stream" });
                                                                    formData.append("file[]", blob, fileName);


                                                                    if (loadedAttchCount === attchCount)
                                                                        processCC();
                                                                }
                                                            } else {
                                                                //result.asyncContext.currentItem.getAttachmentContentAsync FAILED
                                                            }
                                                        });
                                                    }
                                                } else {// no attachments
                                                    processCC();
                                                }
                                            } else {
                                                //email.getAttachmentsAsync FAILED
                                            }
                                        });
                                    }// body
                                });
                        } //subject
                    });
            } //to
        });
}


function sendRequestX(data) {
    getSavedMessage();
}

function processCC() {
    var cc;
    currentMail.cc.getAsync(
        function callback(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                if (result.value.length > 0 && (cc = getEmailAddresses(result.value)) !== "") {
                    formData.append("cc", cc);
                }
            }
            sendRequest();
        });
}

function sendRequest() {

    var xhr = new XMLHttpRequest();

    //xhr.open('POST', 'https://secure.sslpost.com/app/xml/encrypt/');
    //xhr.open('POST', 'https://portal.mailadoc.co.uk/app/xml/encrypt/');

    xhr.open('POST', appInfo.endpoint);

    xhr.onreadystatechange = function () {
        //xhr.onload = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {

            var stat = xhr.status;

            if (stat === 200) {

                getSavedMessage();

            } else {
                if (xhr.responseXML !== null) {
                    showNotification(WARNING_TEXT, "Message Not Sent!",
                        xhr.responseXML.firstChild.firstChild.firstChild.textContent);
                } else
                    showNotification(ERROR_TEXT, "Error!", "An unkown error had occurred");
            }

        }
    };

    showSpinner(true);

    showSpinnerText("Sending Message..");
    xhr.send(formData);
}
function getSavedMessage() {

    showSpinnerText("Moving message to 'Sent Items' folder..");


    if (currentMailID !== undefined) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', restHost + '/v2.0/me/messages/' + currentMailID + '?$select=Sender,From,Subject,Body,ToRecipients,CcRecipients');
        xhr.setRequestHeader("Authorization", "Bearer " + accessToken);

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {

                var stat = xhr.status;
                if (stat === 200 || stat === 201) {

                    var mail;
                    try {
                        mail = JSON.parse(xhr.responseText);
                    } catch (err) {
                        //showNotification(WARNING_OK_TEXT, "getSavedMessage", "Cannot process response!");
                        showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
                        return;
                    }
                    generateNewMail(mail);
                } else if (stat === 404) {
                    attempts = 0;
                    waitForMessage();
                } else {
                    showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
                }
            }
        };

        xhr.send();
    } else {

        attempts = 0;
        waitForMessage();
    }
}

var attempts;
function waitForMessage() {

    var xhr = new XMLHttpRequest();

    attempts++;
    showSpinnerText("Moving message.. " + (MAX_ATTEMPTS - attempts));

    if (currentMailID === undefined || accessToken === undefined) {
        showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
        return;
    }

    xhr.open('GET', restHost + '/v2.0/me/messages/' + currentMailID + '?$select=Sender,From,Subject,Body,ToRecipients,CcRecipients');
    xhr.setRequestHeader("Authorization", "Bearer " + accessToken);

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {

            var stat = xhr.status;
            if (stat === 200 || stat === 201) {
                var mail;
                try {
                    mail = JSON.parse(xhr.responseText);
                } catch (err) {
                    //showNotification(WARNING_OK_TEXT, "getSavedMessage2", "Cannot process response!");
                    showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
                    return;
                }
                generateNewMail(mail);
            } else if (attempts >= MAX_ATTEMPTS) {
                showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
            } else {
                setTimeout(waitForMessage, 1000);
            }
        }
    };

    xhr.send();

}

//var isSending;
function generateNewMail(mail) {

    //if (isSending) {
    //    return;
    //}

    newMail.Sender = mail.Sender;
    newMail.From = mail.From;
    newMail.Subject = mail.Subject;
    newMail.Body = mail.Body;
    newMail.ToRecipients = mail.ToRecipients;
    newMail.CcRecipients = mail.CcRecipients;
    newMail.SingleValueExtendedProperties = SENT_MAIL_FLAG;
    newMail.Attachments = attachements;

    var txt = JSON.stringify(newMail);
    createMailCopy(txt);
}

var SUCCESS_MSG = "SUCCESS! Your email has been sent securely via " + appInfo.name + " encryption add-in. Please disregard the message below, it was auto-generated by Outlook.";
//var SUCCESS_MSG = "Your email has been sent.";

function createMailCopy(emailJson) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', restHost + '/v2.0/me/MailFolders/sentitems/messages/');

    xhr.setRequestHeader("Authorization", "Bearer " + accessToken);
    xhr.setRequestHeader("Content-Type", "application/json");

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {

            var stat = xhr.status;
            if (isUIless) {//from auto-send
                if (stat === 200 || stat === 201) {
                    if (Office.context.mailbox.diagnostics.hostName === 'Outlook' &&
                        navigator.platform !== null && navigator.platform.toLowerCase().indexOf("mac") >= 0) {
                        SUCCESS_MSG = SUCCESS_MSG.replace('below', 'above');
                    }
                    //showNotification(OK_TEXT, "", SUCCESS_MSG);
                    deleteCurrentMail();
                } else {
                    showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
                }
            } else {
                if (stat === 200 || stat === 201) {
                    //showNotification(OK_TEXT, "SUCCESS!",
                    //    "Your email has been sent securely via the " + appInfo.name + " encryption add-in.");
                    deleteCurrentMail();
                } else {
                    showNotification(WARNING_OK_TEXT, "Message Sent!", "But couldn't be moved to 'Sent Items' folder.");
                }
            }
        }
    };

    xhr.send(emailJson);
}

function deleteCurrentMail() {
    var xhr = new XMLHttpRequest();
    xhr.open('DELETE', restHost + '/v2.0/me/messages/' + currentMailID);

    xhr.setRequestHeader("Authorization", "Bearer " + accessToken);
    xhr.setRequestHeader("Content-Type", "application/json");

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {

            if (isUIless) {//from auto-send
                showNotification(OK_TEXT, "", SUCCESS_MSG);
            } else {
                showNotification(OK_TEXT, "SUCCESS!",
                    "Your email has been sent securely via the " + appInfo.name + " encryption add-in.");
            }
        }
    };

    xhr.send();
}

function getItemRestId(id) {
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        return id;
    } else {
        // Convert to an item ID for API v2.0.
        return Office.context.mailbox.convertToRestId(id, Office.MailboxEnums.RestVersion.v2_0);
    }
}

function getEmailAddresses(recipient) {
    var emailAddresses = "";

    for (var i = 0; i < recipient.length; i++) {
        if (recipient[i].emailAddress !== undefined)
            emailAddresses += recipient[i].emailAddress + ",";
    }


    return emailAddresses;
}

function decodeBase64(base64) {
    "use strict";

    var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

    // Use a lookup table to find the index.
    var lookup = new Uint8Array(256);
    for (var i = 0; i < chars.length; i++) {
        lookup[chars.charCodeAt(i)] = i;
    }


    var bufferLength = base64.length * 0.75,
        len = base64.length, p = 0,
        encoded1, encoded2, encoded3, encoded4;

    if (base64[base64.length - 1] === "=") {
        bufferLength--;
        if (base64[base64.length - 2] === "=") {
            bufferLength--;
        }
    }

    var arraybuffer = new ArrayBuffer(bufferLength),
        bytes = new Uint8Array(arraybuffer);

    for (i = 0; i < len; i += 4) {
        encoded1 = lookup[base64.charCodeAt(i)];
        encoded2 = lookup[base64.charCodeAt(i + 1)];
        encoded3 = lookup[base64.charCodeAt(i + 2)];
        encoded4 = lookup[base64.charCodeAt(i + 3)];

        bytes[p++] = (encoded1 << 2) | (encoded2 >> 4);
        bytes[p++] = ((encoded2 & 15) << 4) | (encoded3 >> 2);
        bytes[p++] = ((encoded3 & 3) << 6) | (encoded4 & 63);
    }

    return arraybuffer;

}


var isPopupVisible;
function showPopup(show, spinner) {

    if (show) {
        if (!isPopupVisible) {

            if (spinner) {
                document.getElementById("progress-indicator").style.display = "block";
                document.getElementById("notification-popup").style.display = "none";

            } else {
                document.getElementById("notification-popup").style.display = "block";
                document.getElementById("progress-indicator").style.display = "none";
            }

            document.getElementById("overlay").style.display = "block";
            isPopupVisible = true;
        }
    }
    else {//hide
        if (isPopupVisible) {
            document.getElementById("overlay").style.display = "none";
            isPopupVisible = false;
        }
    }
}

function showSpinner(show) {
    if (isUIless)
        return;

    showPopup(show, true);
}

function showSpinnerText(text) {

    if (isUIless) {//running on auto-send
        showMessageNotification("ok", text);
    } else {
        $("#spin-msg").text(text);
    }
}

function showMessageNotification(type, content) {

    try {
        if (type === "ok") {
            currentMail.notificationMessages.replaceAsync(infoMsgKey, {
                type: "informationalMessage",
                message: content,
                persistent: false,
                icon: "about16"
            });
        } else {
            currentMail.notificationMessages.replaceAsync(infoMsgKey, {
                type: "errorMessage",
                message: content
            });
        }
    } catch (err) {
        var ex = err;
    }
}
function showNotification(type, header, content) {

    if (isUIless) {//running on auto-send
        if (header.length === 0) {
            showMessageNotification(type, content);
        } else {
            showMessageNotification(type, header + " " + content);
        }
        currentEvent.completed({ allowEvent: false });
    } else {
        showSpinner(false);
        $("#notification-title").text(header);
        $("#notification-msg").text(content);
        $("#notification-icon").attr("src", "images/" + type + ".png");
        showPopup(true, false);
    }
}




