// Should we be referencing node.d.ts?
// declare function require(name: string):any;f
export module Localization {
	export enum FontFamily {
		Regular,
		Bold,
		Light,
		Semibold,
		Semilight
	}

	export function getLanguageFileWithFallback(lang: string): string {
		if (!lang) {
			throw new Error("stringId must be a non-empty string, but was: " + lang);
		}

        var head = document.getElementsByTagName('head')[0];
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.onload = function () {
            console.log("yeah!");
        };
        script.onerror = function () {
            if (lang != "en-us") {
                // try English
                getLanguageFileWithFallback("en-us");
            }
        }
        script.src = lang + "/onenote_strings.js";
        head.appendChild(script);        
	}
}
