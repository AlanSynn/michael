<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# assets

## Purpose
Static images referenced by `manifest.xml` (add-in icons) and the taskpane HTML
(brand banner). Copied verbatim into both `dist/` entries by webpack `CopyWebpackPlugin`
+ `asset/resource`; no processing.

## Key Files
| File | Description |
|------|-------------|
| `icon-*.png` / `*.ico` | Manifest add-in icons (16/32/80/128). Referenced by `<bt:Image>` resources. |
| `meet-michael-{black,white}.png` | Brand banner shown on the landing screen; black for dark theme, white for light. Resized to 640×310 (~50KB) to cut payload. |
| `*.svg` | Vector marks where raster was unnecessary. |

## For AI Agents

### Working In This Directory
- **Two banner variants must stay in sync** (black + white) — both referenced by
  theme-switching code. If you resize one, resize the other (`sips --resampleWidth 640`).
- **Manifest icon sizes are contractual** — Outlook shows wrong/garbled icons if a declared
  `<bt:Image size="N">` is missing or mismatched. Validate the manifest after icon changes.
- Keep files lossless-optimized; large PNGs ship to every taskpane load. Target tens of KB, not hundreds.

### Testing Requirements
- No tests. Verify visually after build: open `dist/taskpane.html` + validate `dist/manifest.prod.xml`.

### Common Patterns
- Filenames kebab-case; theme-paired assets use `-black`/`-white` (or `-dark`/`-light`) suffixes.

## Dependencies

### Internal
- Referenced by `manifest.xml`, `src/taskpane/taskpane.html`, webpack copiers.

### External
- `sips` (macOS) for one-off resize; otherwise none.

<!-- MANUAL: -->
