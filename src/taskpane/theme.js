// Pure theme resolution — no DOM, no Office.
// Single source of truth for the dark/light fallback (the prior code used
// "dark" in applyCurrentTheme and "system" in onSettingsChanged).

export const DEFAULT_THEME_FALLBACK = "dark";

/**
 * Determine if a hex color is dark by computing perceived brightness.
 * @param {string} color - hex color code
 * @returns {boolean}
 */
export function isDarkTheme(color) {
  if (!color || typeof color !== "string") {
    return false;
  }

  try {
    const hex = color.replace("#", "");
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);

    if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
      return false;
    }

    return 0.299 * r + 0.587 * g + 0.114 * b < 128;
  } catch {
    return false;
  }
}

/**
 * Resolve the effective theme ("dark" | "light").
 * - explicit "light"/"dark" wins
 * - "system" (or empty) follows the Office theme when available, else the fallback
 * @param {string|undefined} savedTheme
 * @param {{ bodyBackgroundColor?: string }=} officeTheme
 * @returns {"dark" | "light"}
 */
export function resolveTheme(savedTheme, officeTheme) {
  if (savedTheme === "light" || savedTheme === "dark") {
    return savedTheme;
  }

  const bg = officeTheme ? officeTheme.bodyBackgroundColor : undefined;
  if (bg) {
    return isDarkTheme(bg) ? "dark" : "light";
  }

  return DEFAULT_THEME_FALLBACK;
}
