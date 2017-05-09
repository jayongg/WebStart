namespace Localization {

    const defaultLang = "en-us";

	export function setLanguageStringsAsync(lang: string): Promise<any> {
        if (!lang) {
            lang = defaultLang;
        }

        var promise = new Promise<any>(function(resolve:any, reject:any) {
            var head = document.getElementsByTagName('head')[0];
            var script = document.createElement('script');
            script.type = 'text/javascript';
            script.onload = function (val:any) {
                resolve(val);
            };
            script.onerror = function (error) {
                if (lang != defaultLang) {
                    // try English
                    setLanguageStringsAsync(defaultLang)
                        .then(resolve)
                        .catch(reject);
                }
                else {
                    reject(error);
                }
            }
            script.src = "../" + lang + "/onenote_strings.js";
            head.appendChild(script);        
        });

        return promise;
	}
}
