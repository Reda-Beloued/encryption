var appInfo = { id: "", name: "", service: "", endpoint: "" };
var appData =
    [
        { id: "sslpost", name: "SSLP365", service: "SSLPost", endpoint: "https://secure.sslpost.com/app/xml/encrypt/" },
        { id: "mailadoc", name: "MailaDoc365", service: "MailaDoc", endpoint: "https://portal.mailadoc.co.uk/app/xml/encrypt/" },
        { id: "securedd", name: "SecureDD365", service: "SecureDD", endpoint: "https://secure.sslposteurope.com/app/xml/encrypt/" }
    ];

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const appid = urlParams.get('appid');

appInfo= appData.find(ap => ap.id === appid);
