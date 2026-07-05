// Pure leaf: builds a "today is ..." context string for calendar prompts, in
// the user's LOCAL timezone. Lives here (not in calendar.js) so it is trivially
// unit-testable without dragging in Office/DOM/generation. Calendar prompts
// inject this so the model can resolve relative dates ("tomorrow", "next
// Friday", "in two weeks") into absolute local YYYY-MM-DDTHH:mm:ss.

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/**
 * Render the current-date context for calendar prompts.
 *
 * @param {Date} [now=new Date()] injection point for tests; in production the
 *   default `new Date()` resolves to the Office host's local time.
 * @returns {string} e.g. "Today is Friday, 2026-07-05. Current local time:
 *   14:30 (Asia/Seoul, UTC+09:00)."
 */
export function getCurrentDateContext(now = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  const hm = `${pad(now.getHours())}:${pad(now.getMinutes())}`;
  const weekday = WEEKDAYS[now.getDay()];

  // IANA tz name (e.g. "Asia/Seoul") — best-effort; older runtimes lack it.
  let tzName = "local";
  try {
    tzName = Intl.DateTimeFormat().resolvedOptions().timeZone || tzName;
  } catch {
    // keep "local"
  }

  // getTimezoneOffset() is positive for zones BEHIND UTC (e.g. UTC-5 → 300),
  // so flip the sign to produce the conventional UTC+/- offset.
  const offMin = -now.getTimezoneOffset();
  const sign = offMin >= 0 ? "+" : "-";
  const abs = Math.abs(offMin);
  const offset = `UTC${sign}${pad(Math.floor(abs / 60))}:${pad(abs % 60)}`;

  return `Today is ${weekday}, ${ymd}. Current local time: ${hm} (${tzName}, ${offset}).`;
}
