# Repository working rules

This project automates standard e-mail subscriptions for paginated (RDL) reports.
Read `CONTRIBUTING.md`, the affected files and the existing validation workflow before changing anything.
Respond in the user's language.

## Git and pull requests

- Inspect the current branch and working tree before editing and before committing. Preserve unrelated work.
- If this folder has no Git repository, say so. Do not invent a branch, commit or pull request.
- After repository bootstrap, never develop on or push directly to `main`. Use a focused `feature/*`, `fix/*`
  or `docs/*` branch and submit the change through a pull request.
- External contributions require an issue and a maintainer invitation before a pull request.
- Use Conventional Commits and commit only changes belonging to the current task.
- Push or open a pull request only when the user requests it. A completed local change is not permission to publish.
- The `validate` check must pass before merge. Do not disable checks, bypass the ruleset, force-push, or delete
  the default branch.
- The maintainer decides when to merge and release. Never enable automatic merging without an explicit request.

## Implementation boundaries

- Preserve preview by default, explicit `-Apply`, duplicate handling, the pre-check error gate and read-back
  verification. Never introduce subscription deletion as an incidental cleanup.
- Do not contact a report server, run `-Apply` or `-Overwrite -Apply`, or send test mail without explicit approval
  for that server and operation. Local validation must work without server access.
- Maintain Windows PowerShell 5.1 compatibility. Keep `.ps1` and `.psd1` files ASCII-only.
- Keep changes small. Add focused regression cases for behavior changes and update directly related documentation.
- Treat the workbook generator as the template source. Do not overwrite a user's working subscription workbook.
- Do not read `.env` files, credentials, private workbooks or production logs. Use synthetic values in examples,
  issue reports and tests; never commit production recipients, server names or parameter values.

## Evidence

Run the relevant checks in `.github/workflows/validate.yml` before a pull request.
Parsing and linting do not prove SOAP behavior, report execution or e-mail delivery.
Record which product/version was actually exercised; do not describe Power BI Report Server as live-tested
based only on an SSRS run.

The GitHub-side protection is defined in `.github/rulesets/main.json`. That file is not automatically enforced;
follow `.github/REPOSITORY_SETUP.md` when the repository is created.
