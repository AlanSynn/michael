<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# workflows

## Purpose
GitHub Actions pipeline(s). Builds the add-in, runs lint + unit tests, then deploys
the prod build to GitHub Pages (`alansynn.com/michael/` is served from Pages).

## Key Files
| File | Description |
|------|-------------|
| `ci.yml` | On push to `main`: install, lint, unit test, build (prod), deploy `dist/` to Pages. Manifest prod transform + validate happen in the build step. |

## For AI Agents

### Working In This Directory
- **Pipeline order matters:** lint → test → build → deploy. Do not let deploy run before test/build pass.
- **Pages deploys can stall.** If `gh-pages` hasn't updated after a green run, trigger a rebuild:
  `gh api -X POST repos/AlanSynn/michael/pages/builds`.
- **No secrets needed for the prod build** — the API key is runtime/user-supplied, not CI-injected.
  Keep it that way; never add the Z.AI key as a CI secret that lands in the bundle.
- Validate the emitted prod manifest in CI if a validate step exists (`dist/manifest.prod.xml`).

### Testing Requirements
- Workflows are tested by running them. After edits, push to a branch + watch the Actions tab,
  or act locally: `npm run lint && npm test && npm run build && npx office-addin-manifest validate dist/manifest.prod.xml`.

### Common Patterns
- Pin action versions (`@vN`) to avoid surprise breakage.
- Cache `~/.npm` + `node_modules` for speed.

## Dependencies

### Internal
- Drives the root `package.json` scripts (`lint`, `test`, `build`).
- Deploys `dist/` (webpack prod output).

### External
- GitHub Actions, GitHub Pages (peaceiris-style deploy action).

<!-- MANUAL: -->
