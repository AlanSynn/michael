// Language-code → display name. Shared by the taskpane flows and the commands
// function-file.

/** @type {ReadonlyArray<{value: string, label: string}>} */
export const LANGUAGE_OPTIONS = Object.freeze([
  { value: "ko", label: "Korean" },
  { value: "en", label: "English" },
  { value: "ja", label: "Japanese" },
  { value: "zh_cn", label: "Chinese" },
  { value: "es", label: "Spanish" },
  { value: "fr", label: "French" },
  { value: "de", label: "German" },
  { value: "it", label: "Italian" },
]);

const LANGUAGE_TEXT = Object.freeze(
  Object.fromEntries(LANGUAGE_OPTIONS.map((option) => [option.value, option.label]))
);

/**
 * @param {string} languageCode
 * @returns {string}
 */
export function getLanguageText(languageCode) {
  return LANGUAGE_TEXT[languageCode] || "English";
}
