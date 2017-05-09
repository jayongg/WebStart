declare var JSADDINS:any;

var agaveList:any = {
    MD: {
        url: "https://www.onenote.com/"
    },

};

function getParameterByName(name:string, url:string) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
}

var agaveId = getParameterByName("onenoteagave", window.location.href);

if (!agaveList[agaveId]) {
}

if (navigator.onLine) {
    // we are online
    var agave = agaveList[agaveId];
    var destUrl = agave.url;
    var qparams = window.location.search;

    if (destUrl.indexOf("?") >= 1) {
        if (qparams.length && qparams[0] == "?") {
            qparams = "&" + qparams.substring(1);
        }
    }

    destUrl += qparams;

    // navigate to destination URL
    window.location.href = destUrl; 
} else {
    // show an error otherwise
    var lang = getParameterByName("lang", window.location.href);

    Localization.setLanguageStringsAsync(lang)
        .then(function () {
            var noInternet = JSADDINS.Strings[agaveId + "_NoInternet"];
            var retryString = JSADDINS.Strings[agaveId + "_Retry"];

            var messageElement = document.getElementById("message");
            messageElement.innerText = noInternet;
            var retryElement = document.getElementById("retryButton");
            retryElement.innerText = retryString;
            retryElement.onclick = () => {
                window.location.reload(true);
            };
            var mainElement = document.getElementById("main");
            mainElement.className = "";
        })
}