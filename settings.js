
var settings =
{
    username: "",
    password: "",
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


