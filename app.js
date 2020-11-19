var appInfo = { id: "", name: "", service: "", endpoint: "", configUrl:null };

//CORS services:
//https://cors-anywhere.herokuapp.com/
//https://cors-proxy.htmldriven.com/?url=
//http://www.whateverorigin.org/get?url=
//http://alloworigin.com/get?url=
//https://api.allorigins.win/get?url=

//https://yacdn.org/proxy/

var appData =
    [
        //"https://yacdn.org/proxy/https://drive.google.com/uc?export=download&id=1BClUOjdvamSVH-3okpO-XYhApLz_tYY6"
        {
            id: "sslpost", name: "SSLP365", service: "SSLPost",
            endpoint: "https://secure.sslpost.com/app/xml/encrypt/",
            configUrl: "https://yacdn.org/proxy/https://drive.google.com/uc?export=download&id=1BClUOjdvamSVH-3okpO-XYhApLz_tYY6"
        },
        { //reda.exclusive@gmail.com//gFQdA2N4LrD94vS+

            id: "mailadoc", name: "MailaDoc365", service: "MailaDoc",
            endpoint: "https://portal.mailadoc.co.uk/app/xml/encrypt/",
            configUrl: null
        },
        {
            id: "securedd", name: "SecureDD365", service: "SecureDD",
            endpoint: "https://secure.sslposteurope.com/app/xml/encrypt/",
            configUrl: null
        }
    ];

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);
const appid = urlParams.get('appid');

appInfo= appData.find(ap => ap.id === appid);
