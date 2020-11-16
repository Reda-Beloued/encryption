
var settings =
{
    username: "",
    password: "",
    useAutoSend:false,
    removeKeyword: false,
    autoSendKeywordList: ""
};


function loadSettings() {

    var settingStorage = Office.context.roamingSettings;
    var settingsStr = settingStorage.get(appInfo.id + "Settings");

    if (settingsStr !== undefined && settingsStr !== null) {

        try {
            settings = JSON.parse(settingsStr);
        } catch (err) {
            settings = null;
        }
    } else {
        settings = null;
    }
}


function loadKeywords(onSettingsLoadedCallback) {

    var xhr = new XMLHttpRequest();
    xhr.open('GET', appInfo.configUrl);

    xhr.onreadystatechange = function () {

        if (xhr.readyState === XMLHttpRequest.DONE) {

            var stat = xhr.status;
            if (stat === 200 && xhr.responseText !== null && xhr.responseText.length > 0) {
                settings.autoSendKeywordList = xhr.responseText;
            } 

            onSettingsLoadedCallback();
        }
    };

    xhr.send();
}

function enableDiv() {
    $('#auto-encrypt-div').children().prop("disabled",
        !$('#chk-auto-encrypt').prop("checked") === true);
}

function showSettings(showKeywords) {
    if (settings === null)
        return;

    $('#input-username').val(settings.username);
    $('#input-password').val(settings.password);
    $('#chk-auto-encrypt').prop("checked", settings.useAutoSend === true);
    $('#auto-encrypt-div').children().prop("disabled", !settings.useAutoSend === true);

    $('#chk-remove-keyword').prop("checked", settings.removeKeyword);
    if (showKeywords)
        $('#input-keywords').val(settings.autoSendKeywordList);
    else
        $('#input-keywords').prop("placeholder", "");


}

function updateSettings() {

    if (settings.autoSendKeywordList !== null &&
        settings.autoSendKeywordList.length > 0) {
        $('#input-keywords').val(settings.autoSendKeywordList);
    } else {
        $('#input-keywords').val("error loading keywords!");
        $('#input-keywords').css('color', 'red');
    }
}


function saveSettings() {

    settings =
    {
        username: "",
        password: "",
        useAutoSend: false,
        removeKeyword: false,
        autoSendKeywordList: ""
    };

    settings.username = $('#input-username').val();
    settings.password = $('#input-password').val();
    settings.useAutoSend = $('#chk-auto-encrypt').prop("checked");
    settings.removeKeyword = $('#chk-remove-keyword').prop("checked");
    settings.autoSendKeywordList = $('#input-keywords').val();


    var settingStorage = Office.context.roamingSettings;

    settingStorage.set(appInfo.id + "Settings", JSON.stringify(settings));

    Office.context.roamingSettings.saveAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            Office.context.ui.closeContainer();
        }
    });
}


