export module SecondModule {

	export function getLocalizedString(stringId: string): string {
		if (!stringId) {
			throw new Error("stringId must be a non-empty string, but was: " + stringId);
		}


		throw new Error("getLocalizedString could not find a localized or fallback string: " + stringId);
	}
}