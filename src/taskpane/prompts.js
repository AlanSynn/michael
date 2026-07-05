// Prompt placeholder substitution. Replaces the {subject}/{content}/...
// chain that was repeated across every email flow and the commands file.
//
// Single-pass regex with a function callback: each placeholder is resolved
// against the original template (no cascading) and `$` in values is literal
// (string-replacement `$` interpretation does not apply to function callbacks).

const PLACEHOLDER = /\{(subject|content|language|languageInstructions)\}/g;

function stringify(value) {
  return value == null ? "" : String(value);
}

/**
 * Fill a prompt template by substituting {key} placeholders from `vars`.
 * Unknown placeholders are left intact; null/undefined become "".
 * @param {string} template
 * @param {Record<string, string>} vars
 * @returns {string}
 */
export function fillTemplate(template, vars = {}) {
  if (typeof template !== "string") {
    return "";
  }
  return template.replace(PLACEHOLDER, (_, key) => stringify(vars[key]));
}
