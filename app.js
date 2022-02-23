var appInfo = { id: "", name: "", service: "", accountUrl:"", endpoint: "", configUrl:null };

var appData =
    [
        {
            id: "sslpost", name: "SSLP365", service: "SSLPost",
            endpoint: "https://secure.sslpost.com/app/xml/encrypt/",
            configUrl: null,
            accountUrl:"https://secure.sslpost.com/app/"
        },
        { 
            id: "mailadoc", name: "MailaDoc365", service: "MailaDoc",
            endpoint: "https://portal.mailadoc.co.uk/app/xml/encrypt/",
            configUrl: null,
            accountUrl: "https://portal.mailadoc.co.uk/app/"
        },
        {
            id: "securedd", name: "SecureDD365", service: "SecureDD",
            endpoint: "https://secure.sslposteurope.com/app/xml/encrypt/",
            configUrl: null,
            accountUrl: "https://secure.sslposteurope.com/app/"
        },
        {
            id: "sslpdev", name: "SSLP Dev365", service: "SSLP Dev365",
            endpoint: "https://jollynoyce.dev.sslpost.com/app/xml/encrypt/",
            configUrl: null,
            accountUrl:"https://jollynoyce.dev.sslpost.com/app/"
        }
    ];

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const appid = urlParams.get('appid');

appInfo= appData.find(ap => ap.id === appid);
