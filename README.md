# SSRS and Power BI Report Server Subscriptions

[![Validate](https://github.com/uzuntasgokberk/ssrs-reportserver-subscriptions/actions/workflows/validate.yml/badge.svg)](https://github.com/uzuntasgokberk/ssrs-reportserver-subscriptions/actions/workflows/validate.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Windows PowerShell 5.1](https://img.shields.io/badge/Windows%20PowerShell-5.1-blue)

Create SQL Server Reporting Services (SSRS) e-mail subscriptions in bulk from one Excel file: one row is one subscription. It's a single PowerShell script that talks to the report server's own SOAP API (`ReportService2010.asmx`).

The scope is paginated (RDL) reports, including those hosted on Power BI Report Server; Power BI (`.pbix`) report subscriptions are not supported. SSRS has been live-tested; Power BI Report Server compatibility has not yet been verified by a live run.

The default run is a **preview**: every row is checked against the report server and nothing is changed until you add `-Apply`.

| In the Excel file | On the report server |
|---|---|
| `ReportPath` = `/Sales/Daily Sales` | The report the subscription belongs to |
| `ScheduleType` = `Weekly`, `ScheduleDetail` = `Mon,Thu`, `RunTime` = `07:30` | Runs every Monday and Thursday at 07:30 |
| `To` / `Cc` / `Bcc`, `Subject`, `Comment` | The e-mail |
| `RenderFormat` = `EXCELOPENXML` | The report attached as an Excel file |
| `Parameters` = `Region=North,South` | Report parameter values; empty = the report defaults |

## Features

- **Preview first.** Without `-Apply` the script only reads from the report server and shows what it would do.
- **All or nothing.** With `-Apply`, if even one row fails the checks, no subscription is created or changed.
- **Nothing is deleted.** The script never deletes a subscription.
- **Safe to run again.** A subscription with the same name on the same report is skipped. With `-Overwrite` it is updated from the Excel row (the subscription ID is kept), or reported as unchanged if it already matches.
- **Checked before writing.** The script checks the following, read-only:
  - the report exists
  - the e-mail delivery extension is configured
  - the data sources use stored credentials
  - the report parameters are valid, evaluated live on the server the way the web portal does it; for a missing or invalid value the log lists the valid values
- **Verified after writing.** Every created or updated subscription is read back from the server and compared with its Excel row, field by field.
- **Five schedule types.** `Daily`, `Weekly` (chosen days), `Monthly` (days of the month), `Hourly` (every N hours) and `Once`.
- **Built for debugging.** Every run writes a detailed log and a CSV with one row per Excel row, and returns an exit code.

## Requirements

| Requirement | Details |
|---|---|
| Report server | SSRS in native mode with the `ReportService2010` SOAP endpoint (SSRS 2008 R2 and later), in an edition with subscriptions (Express has none). **Tested with SSRS 2019.** Power BI Report Server has the same endpoint but has not been tested |
| E-mail | E-mail delivery configured on the report server: Report Server Configuration Manager → **E-mail Settings** |
| SQL Server Agent | Running on the instance that hosts the report server database. SSRS runs subscription schedules as Agent jobs |
| PowerShell | **Windows PowerShell 5.1.** PowerShell 7 is not supported: it has no `New-WebServiceProxy` |
| Excel module | [ImportExcel](https://github.com/dfinke/ImportExcel). If it is missing, the script installs it for the current user from the PowerShell Gallery. Microsoft Excel itself is not needed |
| Permissions | Permission to create subscriptions on the reports (the **Browser** role has "Manage individual subscriptions"). To send to other people's addresses you usually need "Manage all subscriptions" (the **Content Manager** role) |
| Data sources | The reports must use **stored credentials** (or none). Subscriptions cannot run with Windows integrated or prompted credentials |

Sign-in uses the current Windows account.

## Quick start

```powershell
git clone https://github.com/uzuntasgokberk/ssrs-reportserver-subscriptions.git
cd ssrs-reportserver-subscriptions
```

1. Open `Subscriptions.xlsx` and replace the example rows with your subscriptions. The example rows have `Enabled = FALSE`; set `Enabled` to `TRUE` for each row that is ready. The **How to use** sheet in the workbook explains every column.
2. Preview. This changes nothing:

   ```powershell
   .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer"
   ```

3. Read the result in the console, or open the CSV in the `Logs` folder. Fix the rows marked `ERROR` and preview again.
4. Apply. At the end of a clean preview the script prints this command, ready to copy:

   ```powershell
   .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer" -Apply
   ```

`-ReportServerUrl` is the **web service** address, which usually ends with `/ReportServer`. It is not the web portal address (`/Reports`). You can find it in Report Server Configuration Manager → **Web Service URL**.

If PowerShell refuses to run the script ("running scripts is disabled on this system"), allow it for the current window only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Unblock-File .\Import-SsrsSubscriptions.ps1      # only needed if the file was downloaded as a ZIP
```

## Usage

| Parameter | Description | Default |
|---|---|---|
| `-ReportServerUrl` | Web service address, e.g. `http://reportserver/ReportServer` | Required |
| `-ExcelPath` | Input Excel file | `Subscriptions.xlsx` next to the script |
| `-WorksheetName` | Worksheet with the subscription rows | `Subscriptions` |
| `-Apply` | Make the changes. Without it the run is a preview | Off |
| `-Overwrite` | Update existing subscriptions with the same name from their Excel row | Off |
| `-AllowOwnerMismatch` | Continue when `ExpectedOwner`, or the owner of an existing subscription, is not the account running the script | Off |

```powershell
$Server = "http://reportserver/ReportServer"

# Preview, then create
.\Import-SsrsSubscriptions.ps1 -ReportServerUrl $Server
.\Import-SsrsSubscriptions.ps1 -ReportServerUrl $Server -Apply

# Update existing subscriptions from the Excel file: preview, then apply
.\Import-SsrsSubscriptions.ps1 -ReportServerUrl $Server -Overwrite
.\Import-SsrsSubscriptions.ps1 -ReportServerUrl $Server -Overwrite -Apply

# Another workbook or worksheet
.\Import-SsrsSubscriptions.ps1 -ReportServerUrl $Server -ExcelPath "D:\ssrs\finance.xlsx" -WorksheetName "Finance"
```

## Excel format

One row per subscription on the `Subscriptions` sheet. Column names matter, their order does not. Unknown columns are ignored with a warning, and missing optional columns count as empty. Empty rows are skipped.

| Column | Required | Description | Example |
|---|---|---|---|
| `Enabled` | Yes | `TRUE` = process the row. `FALSE` or empty = skip it | `TRUE` |
| `SubscriptionName` | Yes | Name of the subscription ("Description" in SSRS). **Must be unique per report**: re-runs match on `ReportPath` + `SubscriptionName` (case-insensitive) | `Daily Sales - Management` |
| `ReportPath` | Yes | Path of the report on the server, not a URL | `/Sales/Daily Sales` |
| `ScheduleType` | Yes | `Daily`, `Weekly`, `Monthly`, `Hourly` or `Once`. See [Schedules](#schedules) | `Weekly` |
| `ScheduleDetail` | Depends | Days for `Weekly`, days of the month for `Monthly`, hours for `Hourly`. Empty for `Daily` and `Once` | `Mon,Thu` |
| `RunTime` | Yes | Time of day, in the **time zone of the report server** | `08:00` |
| `StartDate` | No | First day of the schedule. Empty = today, or tomorrow if `RunTime` has already passed | `2027-01-04` |
| `To` | Yes | Recipients, separated with `;` or `,` | `a@example.com;b@example.com` |
| `Cc` | No | Copy recipients | `team@example.com` |
| `Bcc` | No | Blind copy recipients | `archive@example.com` |
| `Subject` | Yes (column) | Mail subject. An empty cell uses the SSRS default `@ReportName was executed at @ExecutionTime` | `Daily Sales Report` |
| `RenderFormat` | Yes | `EXCELOPENXML`, `PDF`, `CSV`, `WORDOPENXML`, `MHTML`, `IMAGE`, `XML`, `PPTX` or `ATOM` | `EXCELOPENXML` |
| `Comment` | No | Text in the mail body | `The report is attached.` |
| `IncludeLink` | No | `TRUE` = also include a link to the report. Empty = `FALSE` | `FALSE` |
| `ExpectedOwner` | No | `DOMAIN\user` that must run the script. If another account runs it, the row fails (or warns with `-AllowOwnerMismatch`) | `CONTOSO\svc-reports` |
| `Parameters` | No | Report parameters. Empty = the report defaults. See [Report parameters](#report-parameters) | `Region=North,South; Year=2026` |

Dates can be Excel dates or text (`yyyy-MM-dd`, `dd.MM.yyyy`). Times can be Excel times or text (`HH:mm`).

## Schedules

| `ScheduleType` | `ScheduleDetail` | Example | Runs |
|---|---|---|---|
| `Daily` | empty | `RunTime` = `08:00` | Every day at 08:00 |
| `Weekly` | Days: `Mon`, `Tue`, `Wed`, `Thu`, `Fri`, `Sat`, `Sun` or full names; `Weekdays`; `Weekend` | `Mon,Thu` | Every Monday and Thursday |
| `Monthly` | Days of the month, separated with `;`; ranges with `-` | `1;15` or `1-5` | On these days of every month |
| `Hourly` | Every N hours, 1-24 (empty = 1) | `2` | Every 2 hours, starting at `RunTime` |
| `Once` | empty | `StartDate` = `2027-01-04`, `RunTime` = `09:00` | Once, on that date and time |

- **Time zone.** The schedule runs in the time zone of the report server. In step 4 the log shows it, e.g. `Server time zone = UTC+03:00 (schedules run in this time zone)`.
- **Past start dates.** A `StartDate` in the past is moved to the next run, so a re-run does not create schedules that start in the past. A `Once` time in the past is an error for a new subscription; an existing subscription with that name is left alone.
- **Days 29-31.** A `Monthly` day that some months do not have (for example 31) is skipped in those months. The row gets a warning.
- **Decimal commas.** In Excel with a comma as decimal separator, typing `1,15` stores the number 1.15. Separate days with `;`, or keep the cell formatted as Text. The template's `ScheduleDetail` column is already formatted as Text. A decimal number in `ScheduleDetail` is reported as an error.
- **Aliases.** Values are case-insensitive. Turkish day and type names are accepted as well (for example `Pzt`, `Cuma`, `HaftaIci`, `Haftalik`).

## Report parameters

Syntax: `Name=value; Name2=value1,value2`

- Empty column = all parameters use the report defaults, like "Use default" in the web portal.
- Use the parameter **value**, not its label: for example `Region=N`, not `Region=North` if the label of `N` is "North". Parameter names are case-insensitive.
- Separate the values of a multi-value parameter with `,`. Values cannot contain `;` or `,`.
- Write dates as `yyyy-MM-dd`. A fixed date is sent on **every** run, so it gets a warning. Leave the parameter out to keep a relative default such as "yesterday".

Before anything is written, the values are validated on the report server the way the web portal does it: cascading parameters and lists of valid values included. A parameter that is required and has no default, a value that is not valid, an unknown name or several values for a single-value parameter make the row an `ERROR`. The log then lists the values the report accepts:

```text
2026-11-11 09:15:02.114 ERROR [Row 3]   [MISSING/INVALID] Region
2026-11-11 09:15:02.115 INFO  [Row 3]     Valid values for 'Region' (first 3 of 3). Put the value on the left into the Excel file:
2026-11-11 09:15:02.115 INFO  [Row 3]       N   (North)
2026-11-11 09:15:02.115 INFO  [Row 3]       S   (South)
2026-11-11 09:15:02.116 INFO  [Row 3]       W   (West)
2026-11-11 09:15:02.118 ERROR [Row 3] Pre-check error: Parameter 'Region': the value from the Excel row is not valid on SSRS.
```

## Safety rules

1. **Preview is the default.** Nothing is written to the report server without `-Apply`.
2. **All or nothing.** With `-Apply`, if a single row is `ERROR` in the pre-check, no subscription is created or updated. The valid rows are reported as `BLOCKED`. Fix the row or set `Enabled` to `FALSE`.
3. **Never deletes.**
4. **Same name on the same report:**
   - Default: `SKIP_EXISTING`, left alone. Fields that differ from the Excel row are logged as warnings.
   - `-Overwrite`: updated from the Excel row if anything differs (the subscription ID is kept), otherwise `UNCHANGED`.
5. **With `-Overwrite` the Excel row is the source of truth.** All mail settings, the schedule and the parameters are written as they are in the Excel row. A fixed parameter value on the server that is not in the Excel row is removed, so the parameter goes back to the report default.
6. **`-Overwrite` refuses (`ERROR`) when:**
   - more than one subscription on the report has that name
   - the subscription is data-driven
   - the subscription belongs to another user (unless `-AllowOwnerMismatch`)
7. **Last-second check.** Right before each write the state is read again. If it changed since the pre-check, the row is not applied.
8. **Read-back verification.** Every applied subscription is read back and compared with its Excel row. Differences are marked `*_MISMATCH`.

The all-or-nothing rule is a **pre-check gate**, not a transaction across subscriptions. If an error occurs during apply, earlier successful rows remain applied; there is no automatic rollback. Check the run's CSV before retrying.

## The Excel file and the server

The script adds and (with `-Overwrite`) updates subscriptions. It does not mirror the Excel file onto the server:

| In the Excel file | On the server | What to do |
|---|---|---|
| Delete a row, or set `Enabled` to `FALSE` | The subscription stays and keeps sending mail | Delete it in the web portal: report → **Manage** → **Subscriptions** |
| Change `SubscriptionName` | A new subscription is created; the old one keeps running, so **recipients get the mail twice** | Delete the subscription with the old name in the web portal. If you renamed a subscription in the portal, use the same name in the Excel file |
| Change `ReportPath` | A new subscription is created on the new report; the one on the old report stays | Delete the subscription on the old report in the web portal |
| A bulk `-Apply` was a mistake | The created subscriptions stay | The CSV of that run lists every `CREATED` row with its subscription ID; delete those in the web portal |

Before a large run, read the CSV of the preview: it shows what will be created and what will be skipped.

## Ownership

A subscription belongs to the account that runs the script. If that account is disabled later, its subscriptions stop running. For production subscriptions, run the script as a service account, and use `ExpectedOwner` so that a run under the wrong account fails instead of creating subscriptions that belong to a person.

## Output

Every run writes two files to the `Logs` folder next to the script:

- `SSRS_Subscription_<yyyyMMdd_HHmmss>_<MODE>.log`: everything shown on the console, plus debug lines (raw cell values, SOAP details, the full chain of inner exceptions with the SSRS error code and the script line number).
- `SSRS_Subscription_<yyyyMMdd_HHmmss>_<MODE>.csv`: one row per Excel row. Columns: `ExcelRow`, `Action`, `ReportPath`, `SubscriptionName`, `To`, `Cc`, `Bcc`, `Subject`, `RenderFormat`, `IncludeLink`, `Parameters`, `Schedule`, `FirstRun`, `SubscriptionID`, `Owner`, `Message`, `Warnings`. The CSV uses the list separator of your Windows culture, so it opens directly in Excel.

`<MODE>` is `PREVIEW`, `APPLY`, `PREVIEW_OVERWRITE` or `APPLY_OVERWRITE`.

Log levels: `STEP`, `INFO`, `OK`, `WARN`, `ERROR`, `DEBUG`. A run has 7 steps: Environment → Reading Excel → Input check → SSRS connection → Pre-check → Apply and verify → Summary.

```text
INFO  [Row 2]   Report      : /Sales/Daily Sales
INFO  [Row 2]   Subscription: Daily Sales - Management
INFO  [Row 2]   To          : sales.manager@example.com
INFO  [Row 2]   Cc          : sales.team@example.com
INFO  [Row 2]   Subject     : Daily Sales Report
INFO  [Row 2]   Format/Link : EXCELOPENXML / IncludeLink=False
INFO  [Row 2]   Parameters  : Region=North,South
INFO  [Row 2]   Schedule    : daily at 08:00 | starting 2026-11-12 08:00
INFO  [Row 2]   First run   : 2026-11-12 08:00 (Thu, server time)
```

The log and the CSV contain e-mail addresses. Review them before you share them.

### Result codes (`Action`)

| Code | Meaning |
|---|---|
| `DISABLED` | `Enabled` is `FALSE` or empty; the row was skipped |
| `ERROR` | Input, pre-check or apply error. Short message in the CSV, details in the log |
| `PREVIEW_CREATE` | Preview: `-Apply` will create it |
| `PREVIEW_OVERWRITE` | Preview: `-Apply -Overwrite` will update it (the fields that differ are in the message) |
| `BLOCKED` | Valid, but not applied because other rows have errors |
| `SKIP_EXISTING` | A subscription with this name exists; left alone |
| `UNCHANGED` | With `-Overwrite`: the existing subscription already matches the Excel row |
| `CREATED` | Created and verified |
| `OVERWRITTEN` | Updated and verified |
| `CREATED_MISMATCH` / `OVERWRITTEN_MISMATCH` | Applied, but a value on the server differs from the Excel row. See the `[DIFF]` lines in the log |
| `CREATED_UNVERIFIED` / `OVERWRITTEN_UNVERIFIED` | Applied, but could not be read back; check it in the web portal |

### Exit codes

| Code | Meaning |
|---|---|
| `0` | No problems |
| `1` | At least one `ERROR`, `BLOCKED`, `*_MISMATCH` or `*_UNVERIFIED` row, or a fatal error |

## Troubleshooting

| Message in the log | Cause | Fix |
|---|---|---|
| `This script must be run in Windows PowerShell 5.1` | Started with `pwsh` (PowerShell 7) | Open "Windows PowerShell" from the Start menu |
| `Could not connect to SSRS` | Network / VPN, wrong address, or HTTP 401 | Open `http://<server>/ReportServer/ReportService2010.asmx` in a browser |
| `... looks like the web portal address` | The portal address (`/Reports`) was given | Use the web service address (`/ReportServer`) |
| `Report not found, or not a report (ItemType=Unknown)` | Wrong `ReportPath`, or no permission on the report | Check the path in the web portal (case does not matter) |
| `Parameter 'X' is required and has no default` | The report parameter has no default value | Add `X=<value>` to `Parameters`, using a value listed in the log |
| `Parameter 'X': the value from the Excel row is not valid on SSRS` | The value is not in the list the report accepts | Use a value listed in the log: the value, not the label |
| `[rsParameterTypeMismatch]` | Wrong type, e.g. a date format | Write dates as `yyyy-MM-dd` |
| Warning `CredentialRetrieval=Integrated` or `Prompt`, or `[rsInvalidDataSourceCredentialSetting]` | The data source does not use stored credentials | Set the report's or shared data source's credentials to "stored credentials" in the web portal |
| `[rsAccessDenied]` | No permission | Ask the SSRS administrator for subscription permissions on the folder |
| `Owner mismatch` | Another account than `ExpectedOwner` runs the script | Run it as that account, or use `-AllowOwnerMismatch` |
| `[DIFF] TO` (`CREATED_MISMATCH`) | Without "Manage all subscriptions", SSRS can replace the recipient | Ask for the permission, or check the recipient |
| `ActiveState.MissingParameterValue = True` | Created, but a parameter value is missing, so it will not run | Fill in `Parameters` and run with `-Overwrite -Apply` |
| `SSRS has no 'Report Server Email' delivery extension` | E-mail delivery is not configured | Report Server Configuration Manager → E-mail Settings |
| `Could not read the Excel file` | The file is open in Excel, or the sheet name is different | Close the file; check the sheet name (`-WorksheetName`) |
| `N subscriptions have this name` (`-Overwrite`) | The report has several subscriptions with this name | Delete the extra ones in the web portal, or use another name in the Excel file |
| `Could not install the ImportExcel module` | No access to the PowerShell Gallery | Install it manually (`Install-Module ImportExcel -Scope CurrentUser`), or copy it from another machine with `Save-Module` |

When reporting a problem, the `.log` file of the run has all the technical details.

## How it works

1. **Excel.** The rows are read with ImportExcel and validated: required fields, e-mail addresses, date and time, render format, schedule, `Parameters` syntax, duplicates inside the file.
2. **Connection.** `New-WebServiceProxy` connects to `<ReportServerUrl>/ReportService2010.asmx` with the current Windows account. The server version, edition and time zone are logged.
3. **Pre-check (read-only).** For each row: `GetItemType`, `ListSubscriptions` and `GetSubscriptionProperties` (same-name subscriptions), `GetItemDataSources` and `GetDataSourceContents` (credentials), `GetItemParameters` with `ForRendering = false` and then `true` (live parameter validation).
4. **Apply.** `CreateSubscription`, or `SetSubscriptionProperties` for `-Overwrite`, with a `TimedSubscription` event and the `Report Server Email` delivery extension. The schedule is a `ScheduleDefinition` (`DailyRecurrence`, `WeeklyRecurrence`, `MonthlyRecurrence`, `MinuteRecurrence` or none for `Once`). `StartDateTime` has no UTC offset, so the server reads it as its local time.
5. **Verify.** `GetSubscriptionProperties` again, compared field by field with the Excel row.

## Limitations

- The interval is always 1: every day, every week, every month. "Every 2 weeks" is not possible; `Hourly` takes 1-24 hours.
- Patterns such as "first Monday of the month" or "last day of the month", and end dates, are not supported. An existing subscription that uses them is changed to the Excel pattern by `-Overwrite`; the change is shown in the `[DIFF] Schedule` line.
- E-mail delivery and standard (not data-driven) subscriptions only.
- Every subscription gets its own schedule; shared schedules are not used.
- The priority is always `NORMAL` and the report is always attached.
- `Parameters` values cannot contain `;` or `,`.

## Notes for contributors

- `Subscriptions.xlsx` is generated by [`tools/New-SubscriptionTemplate.ps1`](tools/New-SubscriptionTemplate.ps1). To change the template, change that script and run it again rather than saving the workbook in Excel; Excel would add the author name (and possibly a company or sensitivity label) to the file.
- The scripts are **ASCII only**: Windows PowerShell 5.1 reads a file without BOM as ANSI, so a non-ASCII character can break it. The CI checks this.
- Functions read script-level variables with the `$script:` prefix. PowerShell uses dynamic scoping, so a local variable with the same name in a calling function would shadow them.
- In Windows PowerShell 5.1, `@($list)` on a `List[object]` can fail with "Argument types do not match"; the script uses the pipeline instead.
- `GetItemParameters` needs a real `null` for `HistoryID`: PowerShell passes `$null` to a string parameter as `""`, which SSRS rejects. The script passes `[NullString]::Value`.
- SSRS 2019 returns the parameter state in `ParameterStateName` and the type in `ParameterTypeName`. Subscriptions that use the report defaults return no parameter values.
- String comparisons are ordinal, so the Turkish "I/i" casing rules on a tr-TR system do not change the result.

## Related projects

- [ssrs-report-export](https://github.com/uzuntasgokberk/ssrs-report-export): back up SSRS reports, data sources and datasets as files, using the REST API.
- [fabric-powerbi-export](https://github.com/uzuntasgokberk/fabric-powerbi-export): the same for the Power BI service / Microsoft Fabric (PBIP, TMDL and RDL), using the Fabric CLI.

## Contributing

Open an issue first; external pull requests require a maintainer invitation. After initial repository bootstrap, all changes to `main` must arrive through a feature/fix/docs branch and a pull request with the required `validate` check passing. The maintainer merges and releases.

See [CONTRIBUTING.md](CONTRIBUTING.md) for the workflow and local checks, [AGENTS.md](AGENTS.md) for coding-agent rules, and [SECURITY.md](SECURITY.md) for private vulnerability reporting. Use synthetic examples instead of production workbooks or raw logs.

[Repository setup](.github/REPOSITORY_SETUP.md) documents the GitHub protections for this repository. The [ruleset specification](.github/rulesets/main.json) must also be activated in GitHub; committing the file alone does not enforce it.

## License

[MIT](LICENSE) © uzuntasgokberk

This project is not affiliated with or endorsed by Microsoft. SQL Server, SQL Server Reporting Services and Excel are trademarks of the Microsoft group of companies.
