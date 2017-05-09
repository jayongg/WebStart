export module Localization {

    const defaultLang = "en-us";

	export function getLanguageFileWithFallback(lang: string): void {
        if (!lang) {
            lang = defaultLang;
        }

        var head = document.getElementsByTagName('head')[0];
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.onload = function () {
            console.log("yeah!");
        };
        script.onerror = function () {
            if (lang != defaultLang) {
                // try English
                getLanguageFileWithFallback(defaultLang);
            }
        }
        script.src = lang + "/onenote_strings.js";
        head.appendChild(script);        
	}
}
