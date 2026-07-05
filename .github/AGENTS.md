<!-- Parent: ../AGENTS.md -->
<!-- Generated: 2026-07-05 | Updated: 2026-07-05 -->

# .github

## Purpose
Container for GitHub-specific config — CI/CD workflows and (if added) issue/PR templates.
No runtime code; the only meaningful child is `workflows/`.

## Subdirectories
| Directory | Purpose |
|-----------|---------|
| `workflows/` | GitHub Actions pipelines. See `workflows/AGENTS.md`. |

## For AI Agents

### Working In This Directory
- This directory is a container; edit files under `workflows/`, not here.
- Secrets (if any) are referenced via `secrets.*` — never inline tokens.

## Dependencies

### External
- GitHub Actions runner.

<!-- MANUAL: -->
