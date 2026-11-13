# GitHub repository setup

Repository: [`uzuntasgokberk/ssrs-reportserver-subscriptions`](https://github.com/uzuntasgokberk/ssrs-reportserver-subscriptions).
This is the provisioning and verification procedure. Files in Git do not activate GitHub protection by themselves.

## Initial publication (maintainer only)

After the user authorizes the repository name, visibility and publication:

1. Initialize Git if needed and inspect every file to be published. Include only the synthetic distributable
   workbook; exclude production workbooks, logs and secrets.
2. Create the GitHub repository without an automatically generated README or license. Publish the initial
   baseline to `main`, and set it as the default branch. This bootstrap precedes the PR requirement;
   subsequent development must use branches and pull requests.
3. Ensure the README badge/clone links and the script's help link use the canonical repository above.
   After bootstrap, make any further corrections through a PR.
4. Let the `Validate` workflow finish successfully. Its stable required job/check name is `validate`.
5. Enable private vulnerability reporting in **Settings > Code security** when available.
6. Enable automatic deletion of merged branches in **Settings > General**.
7. Import `.github/rulesets/main.json` in **Settings > Rules > Rulesets > Import a ruleset**.
   Keep the ruleset active and verify that it targets the default branch. Ruleset availability depends on
   repository visibility and the GitHub plan; if unavailable, stop and resolve that limitation rather than
   claiming enforcement.
8. Set default Actions token permissions to read-only and disallow Actions from approving pull requests.
   Enable secret scanning and push protection. Keep automatic merging disabled.

## Protection contract

| Setting | Required value |
|---|---|
| Target | Default branch (`main` initially) |
| Enforcement | Active |
| Bypass actors | None |
| Pull request | Required, including owner-authored changes |
| Required status check | `validate` |
| Required approving reviews | `0` (single-maintainer workflow) |
| Required code-owner approval | Off; `CODEOWNERS` still identifies the reviewer |
| Force pushes / branch deletion | Blocked |
| Merge methods | Merge, squash or rebase |
| Merged branch cleanup | Automatic |
| Update branch button | Enabled |
| Actions token / PR approval | Read-only / not allowed |
| Secret scanning / push protection | Enabled |
| Private vulnerability reports | Enabled |
| Automatic merging | Disabled |
| Wiki / Projects / Discussions | Disabled |

This follows the PR/CI protections observed in `analytics-agent-harness` on 2026-11-13.
Its Windows/Ubuntu checks are not copied: this Windows PowerShell project uses its own `validate` check.
Issue-first, invitation-only external contributions and maintainer-only merging are contribution policies;
they are not automatically enforced by this branch ruleset.

## Verify after activation

Read the live ruleset back and compare it with the protection contract. Check that the CI check appears on a
test pull request and that its successful completion allows a maintainer merge without a self-approval.
Confirm that the repository's rules UI reports PR and status-check requirements with no bypass actors.
Do not test protection by force-pushing or deleting `main`.

A ruleset JSON file in Git is only a specification. Do not report the repository as protected until GitHub
confirms that the rule is active.
