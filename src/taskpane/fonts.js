// Font-size token → CSS value.

/** @type {ReadonlyArray<{value: string, css: string}>} */
export const FONT_SIZE_OPTIONS = Object.freeze([
  { value: "small", css: "0.875rem" },
  { value: "medium", css: "1rem" },
  { value: "large", css: "1.125rem" },
]);

const FONT_SIZE_MAP = Object.freeze(
  Object.fromEntries(FONT_SIZE_OPTIONS.map((option) => [option.value, option.css]))
);

/**
 * @param {string} size
 * @returns {string}
 */
export function getFontSizeValue(size) {
  return FONT_SIZE_MAP[size] || "1rem";
}
