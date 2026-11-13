# Contributing

## Start with an issue

Report bugs and ask questions through the issue templates, using fake report paths, recipients and parameters.
For external contributors, pull requests are accepted after the maintainer agrees on the issue and invites a fix.
Unsolicited pull requests may be closed. Report vulnerabilities privately as described in `SECURITY.md`.

## Change workflow

1. Check the current branch and working tree; leave unrelated files untouched.
2. Create a `feature/<topic>`, `fix/<topic>` or `docs/<topic>` branch. Do not work directly on `main`.
3. Make one focused change. Preserve Windows PowerShell 5.1 support and preview-first behavior.
4. Run the relevant validation below. For behavior changes, include a regression case using synthetic input
   and document any server-side checks that still need to be performed.
5. Commit using Conventional Commits, for example `fix(schedule): handle monthly boundaries`.
6. Open a pull request, link the agreed issue where applicable, and complete its checklist.
7. Wait for the required `validate` check and the maintainer's review. Only the maintainer merges and releases.

Every change to the default branch goes through a pull request after the initial repository bootstrap,
including changes authored by the owner. Zero required approvals avoids requiring a sole maintainer to
approve their own pull request; it does not remove the PR or CI requirement.

## Local validation

Use the checks in [the workflow](.github/workflows/validate.yml):

- ASCII check for all `.ps1` and `.psd1` files.
- Syntax parsing with Windows PowerShell 5.1 and PowerShell 7.
- PSScriptAnalyzer with `PSScriptAnalyzerSettings.psd1`; findings must be resolved.

If needed, install the development analyzer for the current user, then run it from the project root:

```powershell
Install-Module PSScriptAnalyzer -Scope CurrentUser
$results = @(Invoke-ScriptAnalyzer -Path . -Recurse -Settings .\PSScriptAnalyzerSettings.psd1)
$results | Format-Table -AutoSize
if ($results.Count -gt 0) { throw 'Resolve analyzer findings before opening a pull request.' }
```

CI is offline with respect to report servers: it must not need production credentials or create subscriptions.
An approved live check should distinguish preview, create, overwrite, read-back and actual scheduled delivery.
SSRS evidence does not establish that the same run has been tested on Power BI Report Server.

## Files and compatibility

- Change `tools/New-SubscriptionTemplate.ps1` to update the distributable workbook. Never regenerate over
  a workbook containing real subscriptions. Example rows must remain disabled and use fictitious values.
- Keep PowerShell source ASCII-only; Windows PowerShell 5.1 can interpret BOM-less Unicode incorrectly.
- Preserve the script's ordinal comparisons, SOAP null handling and existing public parameters.
- Describe the pre-check gate accurately: a runtime failure can leave earlier rows applied; there is no
  cross-subscription transaction or automatic rollback.
- Keep Markdown concise: fewer than 500 lines per file and fewer than 200 lines in `AGENTS.md`.
- Update `CHANGELOG.md` for user-visible changes. Changes remain unreleased until the maintainer creates a release.

## License and repository administration

Contributions use this project's MIT license. Do not copy code or documentation from a differently licensed
project without checking its terms.

The workflow and governance conventions were informed by
[`analytics-agent-harness`](https://github.com/uzuntasgokberk/analytics-agent-harness); its domain tools and
hardening pack are not installed here.

The repository name is `ssrs-reportserver-subscriptions`. GitHub-side setup is documented in
[`.github/REPOSITORY_SETUP.md`](.github/REPOSITORY_SETUP.md).
