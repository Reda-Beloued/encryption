var appInfo = { id: "", name: "", service: "", accountUrl:"", endpoint: "", configUrl:null };


var appData =
    [
        //"https://yacdn.org/proxy/https://drive.google.com/uc?export=download&id=1BClUOjdvamSVH-3okpO-XYhApLz_tYY6"
        {
            id: "sslpost", name: "SSLP365", service: "SSLPost",
            endpoint: "https://secure.sslpost.com/app/xml/encrypt/",
            configUrl: null,
            accountUrl:"https://secure.sslpost.com/app/"
        },
        { //gFQdA2N4LrD94vS+

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
            id: "helloflex", name: "HelloFlex365", service: "HelloFlex group",
            endpoint: "https://esafe.helloflexgroup.com/app/xml/encrypt/",
            configUrl: null,
            accountUrl: "https://esafe.helloflexgroup.com/app/"
        }
    ];

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const appid = urlParams.get('appid');

appInfo= appData.find(ap => ap.id === appid);
