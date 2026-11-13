#requires -Version 5.1

<#
.SYNOPSIS
    Creates SQL Server Reporting Services (SSRS) e-mail subscriptions in bulk from an Excel file:
    preview first, all-or-nothing apply, every result read back and verified.

.DESCRIPTION
    One Excel row = one subscription. The default run is a PREVIEW: the script only reads from the
    report server, checks every row and logs what it would do. Nothing on the report server is
    changed unless -Apply is given.

    Flow (shown as "[STEP x/7]" in the log):
      1) Environment : Windows PowerShell 5.1, ImportExcel module
      2) Excel       : columns, row count
      3) Input       : required fields, e-mail addresses, time and date, format, owner,
                       Parameters syntax, schedule, duplicates inside the Excel file
      4) SSRS        : connection, server info, e-mail delivery extension
      5) Pre-check   : the report exists, a subscription with the same name exists (SKIP / -Overwrite),
                       data source credentials, report parameters validated live on the server
                       (READ-ONLY)
      6) Apply       : only with -Apply. If even ONE row has an ERROR in the pre-check, NOTHING is
                       changed. Every created or updated subscription is read back from the server
                       and compared with its Excel row.
      7) Summary     : console + Logs\*.log + Logs\*.csv

    If the report already has a subscription with the same SubscriptionName:
      - default    : SKIP (left alone; fields that differ from the Excel row are logged as warnings)
      - -Overwrite : UPDATED from the Excel row if anything differs (the subscription ID is kept),
                     otherwise UNCHANGED.
    The script never deletes a subscription.

    Requires Windows PowerShell 5.1 (PowerShell 7 has no New-WebServiceProxy).

.PARAMETER ReportServerUrl
    Web service address of the report server, e.g. "http://reportserver/ReportServer".
    This is the web service (".../ReportServer"), not the web portal (".../Reports").

.PARAMETER ExcelPath
    Input Excel file. Default: Subscriptions.xlsx next to the script.

.PARAMETER WorksheetName
    Worksheet with one row per subscription. Default: Subscriptions.

.PARAMETER Apply
    Without it the script runs as a PREVIEW and changes nothing on the report server.

.PARAMETER Overwrite
    Updates existing subscriptions with the same name from their Excel row. Without -Apply it only
    shows what would change.

.PARAMETER AllowOwnerMismatch
    Continue even if ExpectedOwner in the Excel row, or the owner of an existing subscription, is not
    the account that runs the script.

.EXAMPLE
    .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer"
    Preview: checks every row and shows what would be created. Nothing is changed.

.EXAMPLE
    .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer" -Apply
    Creates the subscriptions once the preview is clean.

.EXAMPLE
    .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer" -Overwrite
    .\Import-SsrsSubscriptions.ps1 -ReportServerUrl "http://reportserver/ReportServer" -Overwrite -Apply
    Updates existing subscriptions from the Excel file: preview first, then apply.

.LINK
    https://github.com/uzuntasgokberk/ssrs-reportserver-subscriptions
#>

# PositionalBinding = $false: an unnamed or extra argument fails loudly instead of being bound to
# the wrong parameter.
[CmdletBinding(PositionalBinding = $false)]
param(
    [switch]$Apply,
    [switch]$Overwrite,
    [switch]$AllowOwnerMismatch,
    [Parameter(Mandatory = $true)]
    [string]$ReportServerUrl,
    [string]$ExcelPath,
    [string]$WorksheetName = 'Subscriptions'
)

$script:ScriptName = Split-Path -Leaf $PSCommandPath

# ---------------------------------------------------------------------------
# PowerShell edition check (the log file is not open yet, so this writes to the console only)
# ---------------------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Host 'ERROR: This script must be run in Windows PowerShell 5.1.' -ForegroundColor Red
    Write-Host '       PowerShell 7 (pwsh) has no New-WebServiceProxy, which the SSRS SOAP connection needs.' -ForegroundColor Red
    Write-Host '       Open "Windows PowerShell" from the Start menu and run:' -ForegroundColor Yellow
    Write-Host ('         Set-Location "{0}"' -f $PSScriptRoot) -ForegroundColor Yellow
    Write-Host '         Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force' -ForegroundColor Yellow
    Write-Host ('         .\{0} -ReportServerUrl "{1}"' -f $script:ScriptName, $ReportServerUrl) -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# NOTE: functions read script variables with the $script: prefix. PowerShell uses dynamic scoping, so
#       a local variable with the same name (case-insensitive) in a calling function could otherwise
#       shadow the value.
# ---------------------------------------------------------------------------
$script:Inv                = [Globalization.CultureInfo]::InvariantCulture
$script:IgnoreCase         = [StringComparison]::OrdinalIgnoreCase   # Ordinal: avoids the tr-TR 'I/i' casing problem
$script:EmailExtensionName = 'Report Server Email'
$script:CurrentUser        = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$AllowedRenderFormats      = @('EXCELOPENXML','PDF','CSV','WORDOPENXML','MHTML','IMAGE','XML','PPTX','ATOM')
$RequiredColumns           = @('Enabled','SubscriptionName','ReportPath','ScheduleType','RunTime','To','Subject','RenderFormat')
$OptionalColumns           = @('StartDate','Cc','Bcc','Comment','IncludeLink','ExpectedOwner','Parameters','ScheduleDetail')
# Schedule: element names and order of the SSRS ScheduleDefinition (taken from the server's WSDL types)
$script:DayOrder   = @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')
$script:MonthOrder = @('January','February','March','April','May','June','July','August','September','October','November','December')
$script:DayNames   = @{ Sunday = 'Sun'; Monday = 'Mon'; Tuesday = 'Tue'; Wednesday = 'Wed'; Thursday = 'Thu'; Friday = 'Fri'; Saturday = 'Sat' }
# These SSRS error codes mean that a parameter value from the Excel file is wrong (a definite error)
$script:ParamValueErrorCodes = @('rsParameterTypeMismatch','rsReportParameterTypeMismatch','rsUnknownReportParameter',
                                 'rsInvalidReportParameter','rsReportParameterValueNotSet','rsInvalidParameter')
# Results that set the exit code to 1
$script:BadActions = @('ERROR','BLOCKED','CREATED_MISMATCH','CREATED_UNVERIFIED','OVERWRITTEN_MISMATCH','OVERWRITTEN_UNVERIFIED')

$RunMode = if ($Apply) { 'APPLY' } else { 'PREVIEW' }
if ($Overwrite) { $RunMode += '_OVERWRITE' }
if (-not $ExcelPath) { $ExcelPath = Join-Path $PSScriptRoot 'Subscriptions.xlsx' }

# Command line used in the hints at the end of a run (ready to copy and paste)
$script:CommandBase = '.\{0} -ReportServerUrl "{1}"' -f $script:ScriptName, $ReportServerUrl
if ($PSBoundParameters.ContainsKey('ExcelPath'))     { $script:CommandBase += (' -ExcelPath "{0}"' -f $ExcelPath) }
if ($PSBoundParameters.ContainsKey('WorksheetName')) { $script:CommandBase += (' -WorksheetName "{0}"' -f $WorksheetName) }

$RunStamp = (Get-Date).ToString('yyyyMMdd_HHmmss', $script:Inv)
$LogDir   = Join-Path $PSScriptRoot 'Logs'
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$script:LogFile = Join-Path $LogDir ('SSRS_Subscription_{0}_{1}.log' -f $RunStamp, $RunMode)
$script:CsvFile = Join-Path $LogDir ('SSRS_Subscription_{0}_{1}.csv' -f $RunStamp, $RunMode)

$script:LogWriter = New-Object System.IO.StreamWriter -ArgumentList $script:LogFile, $false, (New-Object System.Text.UTF8Encoding -ArgumentList $true)
$script:LogWriter.AutoFlush = $true   # the log stays on disk even if the script is interrupted
$script:RowTag   = ''
$script:Proxy    = $null
$script:Ns       = $null
$script:Fatal    = $false
$script:ExitCode = 0
$script:Plans    = New-Object System.Collections.Generic.List[object]

# ===========================================================================
# LOGGING
# ===========================================================================
function Write-RunLog {
    param(
        [ValidateSet('INFO','STEP','OK','WARN','ERROR','DEBUG')][string]$Level = 'INFO',
        [AllowEmptyString()][AllowNull()][string]$Message = ''
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff', [Globalization.CultureInfo]::InvariantCulture)
    $tag  = if ($script:RowTag) { '[' + $script:RowTag + '] ' } else { '' }
    $line = '{0} {1,-5} {2}{3}' -f $ts, $Level, $tag, $Message
    $color = switch ($Level) {
        'STEP'  { 'Cyan' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
        default { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
    if ($script:LogWriter) { $script:LogWriter.WriteLine($line) }
}

function Write-Step {
    param([int]$No, [string]$Title)
    $script:RowTag = ''
    Write-RunLog STEP ''
    Write-RunLog STEP ('==================== [STEP {0}/7] {1} ====================' -f $No, $Title)
}

function Get-SoapErrorCode {
    # SSRS SOAP faults carry an <ErrorCode> in their Detail (e.g. rsItemNotFound)
    param($Exception)
    $ex = $Exception
    while ($null -ne $ex) {
        $p = $ex.PSObject.Properties['Detail']
        if ($p -and $null -ne $p.Value) {
            try {
                $n = $p.Value.SelectSingleNode("//*[local-name()='ErrorCode']")
                if ($n) { return $n.InnerText }
            } catch { Write-Verbose ('Could not read the SOAP fault detail: {0}' -f $_.Exception.Message) }
        }
        $ex = $ex.InnerException
    }
    return $null
}

function Get-ShortError {
    # One-line error message for the CSV and the summary (innermost exception + SSRS error code)
    param($Exception)
    $ex = $Exception
    while ($null -ne $ex.InnerException) { $ex = $ex.InnerException }
    $msg  = (([string]$ex.Message) -split "`r?`n")[0].Trim()
    $code = Get-SoapErrorCode $Exception
    if ($code) { return ('[{0}] {1}' -f $code, $msg) }
    return $msg
}

function Write-ErrorDetail {
    # Logs an error with all inner exceptions, the SSRS error code and the script line number
    param($ErrorRecord, [string]$Context)
    Write-RunLog ERROR ('{0}: {1}' -f $Context, (Get-ShortError $ErrorRecord.Exception))
    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($null -ne $ex -and $depth -lt 10) {
        Write-RunLog DEBUG ('  Exception[{0}] {1}: {2}' -f $depth, $ex.GetType().FullName, (([string]$ex.Message) -replace "`r?`n", ' | '))
        $dp = $ex.PSObject.Properties['Detail']
        if ($dp -and $null -ne $dp.Value -and $dp.Value.PSObject.Properties['OuterXml']) {
            Write-RunLog DEBUG ('  SOAP Detail: {0}' -f $dp.Value.OuterXml)
        }
        $ex = $ex.InnerException
        $depth++
    }
    $invInfo = $ErrorRecord.InvocationInfo
    if ($invInfo -and $invInfo.ScriptLineNumber -gt 0) {
        Write-RunLog DEBUG ('  Script location: line {0} -> {1}' -f $invInfo.ScriptLineNumber, ([string]$invInfo.Line).Trim())
    }
}

function Add-PlanWarning {
    param($Plan, [string]$Text)
    $Plan.Warnings.Add($Text)
    Write-RunLog WARN $Text
}

# ===========================================================================
# CONVERSION / VALIDATION
# ===========================================================================
function ConvertTo-CleanText {
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function Test-TextEqual {
    param($A, $B, [switch]$CaseSensitive)
    $cmp = if ($CaseSensitive) { [StringComparison]::Ordinal } else { $script:IgnoreCase }
    return [string]::Equals((ConvertTo-CleanText $A), (ConvertTo-CleanText $B), $cmp)
}

function Get-Cell {
    param($Row, [string]$Name)
    $p = $Row.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Format-Raw {
    # For debugging: raw type and value of a cell
    param($Value)
    if ($null -eq $Value) { return '<empty>' }
    return ('[{0}] {1}' -f $Value.GetType().Name, $Value)
}

function ConvertTo-Bool {
    param($Value, [bool]$Default, [string]$Field)
    if ($Value -is [bool]) { return $Value }
    $t = (ConvertTo-CleanText $Value).ToLowerInvariant()
    if ($t -eq '') { return $Default }
    if (@('true','1','yes','evet','e','y','x') -contains $t) { return $true }
    if (@('false','0','no','hayir','h','n') -contains $t) { return $false }
    throw ("{0} value not recognized: '{1}' (expected true/false)" -f $Field, $Value)
}

function ConvertTo-DateOnly {
    param($Value)
    if ($null -eq $Value -or (ConvertTo-CleanText $Value) -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.Date }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        # Excel stores dates as numbers (OADate): 46337 = 2026-11-11
        return [datetime]::FromOADate([double]$Value).Date
    }
    $t       = ConvertTo-CleanText $Value
    $formats = [string[]]@('yyyy-MM-dd','dd.MM.yyyy','d.M.yyyy','dd/MM/yyyy','d/M/yyyy','yyyy.MM.dd')
    $parsed  = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, $formats, $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    $num = 0.0
    if ([double]::TryParse($t, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$num)) { return [datetime]::FromOADate($num).Date }
    throw ("StartDate not recognized: '{0}' (example: 2026-11-11 or 11.11.2026)" -f $t)
}

function ConvertTo-TimeOfDay {
    param($Value)
    $t = ConvertTo-CleanText $Value
    if ($t -eq '') { throw 'RunTime is empty. Example: 09:00' }
    if ($Value -is [datetime]) { return $Value.TimeOfDay }
    if ($Value -is [timespan]) { return $Value }
    if ($Value -is [double] -or $Value -is [decimal]) {
        # A time-formatted Excel cell: 09:00 -> 0.375 (fraction of a day)
        $frac    = [double]$Value - [math]::Floor([double]$Value)
        $minutes = ([int][math]::Round($frac * 1440)) % 1440
        return [timespan]::FromMinutes($minutes)
    }
    $formats = [string[]]@('H:mm','HH:mm','H:mm:ss','HH:mm:ss','H.mm','HH.mm')
    $parsed  = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, $formats, $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.TimeOfDay }
    throw ("RunTime not recognized: '{0}' (example: 09:00)" -f $t)
}

function ConvertTo-EmailList {
    param($Value, [string]$Field, [bool]$Required)
    $t = ConvertTo-CleanText $Value
    if ($t -eq '') {
        if ($Required) { throw ('{0} is required.' -f $Field) }
        return ''
    }
    $parts = @($t -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $bad   = @($parts | Where-Object { $_ -notmatch '^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$' })
    if ($bad.Count -gt 0) { throw ('{0} contains an invalid e-mail address: {1}' -f $Field, ($bad -join ', ')) }
    return ($parts -join ';')
}

function Get-NormalizedEmailList {
    param([string]$Value)
    if (-not $Value) { return '' }
    $p = @($Value -split '[;,]' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ -ne '' } | Sort-Object)
    return ($p -join ';')
}

function ConvertTo-ParameterList {
    # "Region=North,South; Year=2026"  ->  list: { Name='Region'; Values=@('North','South') }, { Name='Year'; Values=@('2026') }
    param($Value)
    $list = New-Object System.Collections.Generic.List[object]
    $t = ConvertTo-CleanText $Value
    if ($t -eq '') { return ,$list }
    foreach ($seg in ($t -split ';')) {
        $s = $seg.Trim()
        if ($s -eq '') { continue }
        $eq = $s.IndexOf('=')
        if ($eq -lt 1) { throw ("Parameters syntax error: '{0}'. Example: Region=North,South; Year=2026" -f $s) }
        $name = $s.Substring(0, $eq).Trim()
        $vals = @($s.Substring($eq + 1) -split ',' | ForEach-Object { $_.Trim() })
        if (@($vals | Where-Object { $_ -eq '' }).Count -gt 0) { throw ("Parameter '{0}' has an empty value: '{1}'" -f $name, $s) }
        if (@($list | Where-Object { [string]::Equals($_.Name, $name, $script:IgnoreCase) }).Count -gt 0) {
            throw ("Parameter '{0}' is listed more than once in Parameters." -f $name)
        }
        $list.Add([pscustomobject]@{ Name = $name; Values = [string[]]$vals })
    }
    return ,$list
}

function Format-ParameterList {
    param($List)
    if ($null -eq $List -or $List.Count -eq 0) { return '' }
    # NOTE: in PS 5.1, @($List) on a List[object] fails with "Argument types do not match"; use the pipeline.
    return (($List | ForEach-Object { '{0}={1}' -f $_.Name, ($_.Values -join ',') }) -join '; ')
}

function Format-ParamMap {
    # Text for a parameter dictionary read from SSRS (Name -> List[string])
    param($Map)
    if ($null -eq $Map -or $Map.Count -eq 0) { return '(none - report defaults)' }
    return ((@($Map.Keys) | ForEach-Object { '{0}={1}' -f $_, (@($Map[$_]) -join ',') }) -join '; ')
}

function ConvertTo-ComparableValue {
    # SSRS can return values in another format (e.g. the date '1/1/2026 12:00:00 AM').
    # For comparing: number -> number, date -> ISO, anything else -> lower-case text.
    param($Value)
    $v = ConvertTo-CleanText $Value
    $num = 0.0
    if ([double]::TryParse($v, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$num)) { return 'N:' + $num.ToString('R', $script:Inv) }
    if ($v -match '\d') {
        $dt = [datetime]::MinValue
        foreach ($culture in @($script:Inv, (Get-Culture))) {
            if ([datetime]::TryParse($v, $culture, [Globalization.DateTimeStyles]::None, [ref]$dt)) { return 'D:' + $dt.ToString('yyyy-MM-ddTHH:mm:ss', $script:Inv) }
        }
    }
    return 'S:' + $v.ToLowerInvariant()
}

function Get-NormalizedValueSet {
    param($Values)
    $n = @(@($Values) | Where-Object { $null -ne $_ } | ForEach-Object { ConvertTo-ComparableValue $_ } | Sort-Object)
    return ($n -join '|')
}

# ===========================================================================
# SCHEDULE (Daily / Weekly / Monthly / Hourly / Once)
# ===========================================================================
function ConvertTo-AsciiLower {
    # Maps Turkish letters to ASCII and lower-cases the text, so the Turkish aliases match however they
    # are typed. The characters are written as code points to keep this file ASCII; the tr-TR I/i
    # casing rules are not relied on.
    param([string]$Text)
    $map = @{
        0x00E7 = 'c'; 0x00C7 = 'c'; 0x015F = 's'; 0x015E = 's'; 0x011F = 'g'; 0x011E = 'g'
        0x00FC = 'u'; 0x00DC = 'u'; 0x00F6 = 'o'; 0x00D6 = 'o'; 0x0131 = 'i'; 0x0130 = 'i'; 0x0307 = ''
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in ([string]$Text).ToCharArray()) {
        $code = [int]$ch
        if ($map.ContainsKey($code)) { [void]$sb.Append($map[$code]) } else { [void]$sb.Append($ch) }
    }
    return $sb.ToString().ToLowerInvariant()
}

function ConvertTo-ScheduleType {
    # English values, plus Turkish aliases
    param($Value)
    $t = (ConvertTo-AsciiLower (ConvertTo-CleanText $Value)) -replace '[\s_-]', ''
    $aliases = @{
        daily = 'Daily';     gunluk = 'Daily';    hergun = 'Daily'
        weekly = 'Weekly';   haftalik = 'Weekly'
        monthly = 'Monthly'; aylik = 'Monthly'
        hourly = 'Hourly';   saatlik = 'Hourly'
        once = 'Once';       birkez = 'Once';     teksefer = 'Once'
    }
    if ($t -ne '' -and $aliases.ContainsKey($t)) { return $aliases[$t] }
    throw ("ScheduleType='{0}' is not valid. Valid values: Daily, Weekly, Monthly, Hourly, Once" -f (ConvertTo-CleanText $Value))
}

function ConvertTo-WeekDayList {
    # "Mon,Fri" / "Monday;Friday" / "Weekdays" -> @('Monday','Friday') (SSRS order: Sunday..Saturday).
    # Turkish day names are accepted too.
    param([string]$Text)
    $aliases = @{
        pzt = 'Monday';    pts = 'Monday';    pazartesi = 'Monday';   mon = 'Monday';    monday = 'Monday'
        sal = 'Tuesday';   sali = 'Tuesday';  tue = 'Tuesday';        tuesday = 'Tuesday'
        car = 'Wednesday'; crs = 'Wednesday'; carsamba = 'Wednesday'; wed = 'Wednesday'; wednesday = 'Wednesday'
        per = 'Thursday';  prs = 'Thursday';  persembe = 'Thursday';  thu = 'Thursday';  thursday = 'Thursday'
        cum = 'Friday';    cuma = 'Friday';   fri = 'Friday';         friday = 'Friday'
        cmt = 'Saturday';  cts = 'Saturday';  cumartesi = 'Saturday'; sat = 'Saturday';  saturday = 'Saturday'
        paz = 'Sunday';    pazar = 'Sunday';  sun = 'Sunday';         sunday = 'Sunday'
    }
    $t = (ConvertTo-AsciiLower $Text) -replace 'hafta[\s_-]*ici', 'haftaici' -replace 'hafta[\s_-]*sonu', 'haftasonu'
    $tokens = @($t -split '[;,/\s]+' | Where-Object { $_ -ne '' })
    if ($tokens.Count -eq 0) { throw 'Weekly needs the days in ScheduleDetail. Example: Mon,Fri  or  Weekdays' }
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    $bad = @()
    foreach ($tk in $tokens) {
        if (@('weekdays','weekday','haftaici') -contains $tk) { foreach ($d in @('Monday','Tuesday','Wednesday','Thursday','Friday')) { [void]$set.Add($d) }; continue }
        if (@('weekend','weekends','haftasonu') -contains $tk) { foreach ($d in @('Saturday','Sunday')) { [void]$set.Add($d) }; continue }
        if ($aliases.ContainsKey($tk)) { [void]$set.Add($aliases[$tk]) } else { $bad += $tk }
    }
    if ($bad.Count -gt 0) { throw ("Unknown day in ScheduleDetail: {0}. Use Mon,Tue,Wed,Thu,Fri,Sat,Sun (or full names), Weekdays or Weekend" -f ($bad -join ', ')) }
    return ($script:DayOrder | Where-Object { $set.Contains($_) })
}

function ConvertTo-MonthDayList {
    # "1;15" / "1,15" / "1-5;20" -> 1,2,3,4,5,20 (sorted, no duplicates)
    param([string]$Text)
    $t = ConvertTo-CleanText $Text
    if ($t -eq '') { throw 'Monthly needs the days of the month in ScheduleDetail. Example: 1  or  1;15  or  1-5' }
    $set = New-Object 'System.Collections.Generic.SortedSet[int]'
    foreach ($part in ($t -split '[;,]')) {
        $p = $part.Trim()
        if ($p -eq '') { continue }
        if ($p -match '^(\d{1,2})\s*-\s*(\d{1,2})$') { $a = [int]$Matches[1]; $b = [int]$Matches[2] }
        elseif ($p -match '^\d{1,2}$')              { $a = [int]$p; $b = $a }
        else { throw ("Day of month not recognized in ScheduleDetail: '{0}'. Example: 1;15 or 1-5" -f $p) }
        if ($a -lt 1 -or $b -gt 31 -or $a -gt $b) { throw ("Day of month must be between 1 and 31: '{0}'" -f $p) }
        for ($d = $a; $d -le $b; $d++) { [void]$set.Add($d) }
    }
    return ($set | ForEach-Object { [int]$_ })
}

function ConvertTo-HourInterval {
    # Hourly: "2" / "2h" / "2 hours" -> 120 minutes. Empty = every hour.
    param([string]$Text)
    $t = ConvertTo-AsciiLower (ConvertTo-CleanText $Text)
    if ($t -eq '') { return 60 }
    if ($t -notmatch '^(\d{1,2})\s*(hours|hour|hrs|hr|h|saat|sa|s)?$') { throw ("Hourly needs the interval in hours in ScheduleDetail (1-24). Value: '{0}'" -f $Text) }
    $h = [int]$Matches[1]
    if ($h -lt 1 -or $h -gt 24) { throw ("The hourly interval must be 1-24 hours. Value: {0}" -f $h) }
    return ($h * 60)
}

function New-ScheduleXml {
    # SSRS ScheduleDefinition XML.
    # StartDateTime has no UTC offset, so the report server reads it as its own local time.
    param($Plan)
    $s = $Plan.FirstRun.ToString('s', $script:Inv)
    $rec = ''
    switch ($Plan.ScheduleType) {
        'Daily'   { $rec = '<DailyRecurrence><DaysInterval>1</DaysInterval></DailyRecurrence>' }
        'Weekly'  {
            $days = ($Plan.WeekDays | ForEach-Object { '<{0}>true</{0}>' -f $_ }) -join ''
            $rec  = '<WeeklyRecurrence><WeeksInterval>1</WeeksInterval><DaysOfWeek>' + $days + '</DaysOfWeek></WeeklyRecurrence>'
        }
        'Monthly' {
            $months = ($script:MonthOrder | ForEach-Object { '<{0}>true</{0}>' -f $_ }) -join ''
            $rec    = '<MonthlyRecurrence><Days>' + ($Plan.MonthDays -join ',') + '</Days><MonthsOfYear>' + $months + '</MonthsOfYear></MonthlyRecurrence>'
        }
        'Hourly'  { $rec = '<MinuteRecurrence><MinutesInterval>' + $Plan.IntervalMinutes + '</MinutesInterval></MinuteRecurrence>' }
        'Once'    { $rec = '' }
    }
    return ('<ScheduleDefinition><StartDateTime>{0}</StartDateTime>{1}</ScheduleDefinition>' -f $s, $rec)
}

function Get-ScheduleInfo {
    # Parses the MatchData XML returned by SSRS: type, start, interval, days
    param([string]$MatchData)
    $info = [pscustomobject]@{ Type = 'Unknown'; Start = $null; Interval = ''; WeekDays = @(); MonthDays = @(); MonthsCount = 0 }
    try {
        $x = [xml]$MatchData
        $n = $x.SelectSingleNode("//*[local-name()='StartDateTime']")
        if ($n) { $info.Start = [DateTimeOffset]::Parse($n.InnerText, $script:Inv).DateTime }
        $iv = $null
        if ($x.SelectSingleNode("//*[local-name()='DailyRecurrence']")) {
            $info.Type = 'Daily'; $iv = $x.SelectSingleNode("//*[local-name()='DaysInterval']")
        } elseif ($x.SelectSingleNode("//*[local-name()='WeeklyRecurrence']")) {
            $info.Type = 'Weekly'; $iv = $x.SelectSingleNode("//*[local-name()='WeeksInterval']")
            $dw = $x.SelectSingleNode("//*[local-name()='WeeklyRecurrence']/*[local-name()='DaysOfWeek']")
            $set = @()
            if ($dw) { foreach ($c in $dw.ChildNodes) { if ([string]::Equals(([string]$c.InnerText).Trim(), 'true', $script:IgnoreCase)) { $set += $c.LocalName } } }
            $info.WeekDays = @($script:DayOrder | Where-Object { $set -contains $_ })
        } elseif ($x.SelectSingleNode("//*[local-name()='MonthlyDOWRecurrence']")) {
            $info.Type = 'MonthlyDOW'
        } elseif ($x.SelectSingleNode("//*[local-name()='MonthlyRecurrence']")) {
            $info.Type = 'Monthly'
            $dn = $x.SelectSingleNode("//*[local-name()='MonthlyRecurrence']/*[local-name()='Days']")
            if ($dn) { $info.MonthDays = @(ConvertTo-MonthDayList $dn.InnerText) }
            $mo = $x.SelectSingleNode("//*[local-name()='MonthlyRecurrence']/*[local-name()='MonthsOfYear']")
            if ($mo) { $info.MonthsCount = @($mo.ChildNodes | Where-Object { [string]::Equals(([string]$_.InnerText).Trim(), 'true', $script:IgnoreCase) }).Count }
        } elseif ($x.SelectSingleNode("//*[local-name()='MinuteRecurrence']")) {
            $info.Type = 'Hourly'; $iv = $x.SelectSingleNode("//*[local-name()='MinutesInterval']")
        } elseif ($n) {
            $info.Type = 'Once'
        }
        if ($iv) { $info.Interval = ([string]$iv.InnerText).Trim() }
    } catch { Write-Verbose ('Could not parse the schedule XML: {0}' -f $_.Exception.Message) }
    return $info
}

function Format-ScheduleText {
    # Readable schedule text (for the log and the CSV)
    param([string]$Type, $Start, [string]$Interval, $WeekDays, $MonthDays, [int]$MonthsCount = 12)
    if ($null -eq $Start) { return ('{0} (no start)' -f $Type) }
    $st   = ([datetime]$Start)
    $hhmm = $st.ToString('HH:mm', $script:Inv)
    $from = $st.ToString('yyyy-MM-dd HH:mm', $script:Inv)
    switch ($Type) {
        'Daily'   { $txt = 'daily at {0}' -f $hhmm; if ($Interval -and $Interval -ne '1') { $txt = 'every {0} days at {1}' -f $Interval, $hhmm } }
        'Weekly'  { $txt = 'weekly on {0} at {1}' -f ((@($WeekDays) | ForEach-Object { $script:DayNames[$_] }) -join ','), $hhmm
                    if ($Interval -and $Interval -ne '1') { $txt += (' (every {0} weeks)' -f $Interval) } }
        'Monthly' { $txt = 'monthly on day {0} at {1}' -f (@($MonthDays) -join ','), $hhmm
                    if ($MonthsCount -ne 12) { $txt += (' ({0}/12 months)' -f $MonthsCount) } }
        'Hourly'  { $m = 0; [void][int]::TryParse($Interval, [ref]$m)
                    $txt = if ($m -gt 0 -and $m % 60 -eq 0) { 'every {0} hour(s)' -f ($m / 60) } else { 'every {0} minutes' -f $Interval } }
        'Once'    { return ('once at {0}' -f $from) }
        default   { $txt = $Type }
    }
    return ('{0} | starting {1}' -f $txt, $from)
}

function Format-PlanSchedule {
    param($Plan)
    $iv = if ($Plan.ScheduleType -eq 'Hourly') { [string]$Plan.IntervalMinutes } else { '1' }
    return (Format-ScheduleText $Plan.ScheduleType $Plan.FirstRun $iv $Plan.WeekDays $Plan.MonthDays 12)
}

function Get-NextRun {
    # First actual run after the schedule start (Weekly/Monthly: the first matching day)
    param($Plan)
    $start = [datetime]$Plan.FirstRun
    if ($Plan.ScheduleType -eq 'Weekly') {
        for ($i = 0; $i -lt 8; $i++) { $d = $start.AddDays($i); if (@($Plan.WeekDays) -contains $d.DayOfWeek.ToString()) { return $d } }
        return $null
    }
    if ($Plan.ScheduleType -eq 'Monthly') {
        for ($i = 0; $i -lt 400; $i++) { $d = $start.AddDays($i); if (@($Plan.MonthDays) -contains $d.Day) { return $d } }
        return $null
    }
    return $start
}

function Get-ExtensionValueMap {
    # ExtensionSettings.ParameterValues -> dictionary (Name -> Value), case-insensitive
    param($ExtensionSettings)
    $d = New-Object 'System.Collections.Generic.Dictionary[string,string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    if ($null -ne $ExtensionSettings) {
        foreach ($pv in @($ExtensionSettings.ParameterValues)) {
            if ($null -ne $pv -and $pv.PSObject.Properties['Name']) { $d[[string]$pv.Name] = [string]$pv.Value }
        }
    }
    return ,$d
}

function Get-DictValue {
    param($Dict, [string]$Key)
    if ($null -ne $Dict -and $Dict.ContainsKey($Key)) { return $Dict[$Key] }
    return ''
}

function Get-ParamTypeName {
    # ReportService2010 returns the type in 'ParameterTypeName' (older APIs: 'Type')
    param($P)
    if ($P.PSObject.Properties['ParameterTypeName']) { return [string]$P.ParameterTypeName }
    return [string]$P.Type
}

function Get-ParamState {
    # Parameter state from the live validation.
    # ReportService2010: 'ParameterStateName' (text: HasValidValue, MissingValidValue, ...)
    # Older APIs: 'State' (enum) + 'StateSpecified'. Returns '' if neither exists (state unknown).
    param($P)
    if ($null -eq $P) { return '' }
    if ($P.PSObject.Properties['ParameterStateName']) { return [string]$P.ParameterStateName }
    if ($P.PSObject.Properties['State'] -and ((-not $P.PSObject.Properties['StateSpecified']) -or $P.StateSpecified)) { return [string]$P.State }
    return ''
}

function Import-RequiredModule {
    param([string]$Name)
    $m = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $m) {
        Write-RunLog WARN "$Name module not found. Installing it from the PowerShell Gallery (Scope=CurrentUser)..."
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        } catch {
            throw ("Could not install the {0} module: {1}. Install it manually: Install-Module {0} -Scope CurrentUser" -f $Name, $_.Exception.Message)
        }
        $m = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    }
    Import-Module -Name $Name -RequiredVersion $m.Version -Force -ErrorAction Stop
    Write-RunLog OK ('{0} v{1} loaded ({2})' -f $Name, $m.Version, $m.ModuleBase)
}

# ===========================================================================
# ROW -> PLAN
# ===========================================================================
function New-Plan {
    param([int]$ExcelRow)
    return [pscustomobject]@{
        ExcelRow         = $ExcelRow
        Action           = 'PENDING'
        ReportPath       = ''
        SubscriptionName = ''
        To               = ''
        Cc               = ''
        Bcc              = ''
        Subject          = ''
        RenderFormat     = ''
        Comment          = ''
        IncludeLink      = $false
        ExpectedOwner    = ''
        Parameters       = (New-Object System.Collections.Generic.List[object])
        ParametersText   = ''
        ScheduleType     = 'Daily'      # Daily | Weekly | Monthly | Hourly | Once
        WeekDays         = @()          # Weekly: 'Monday',...
        MonthDays        = @()          # Monthly: 1..31
        IntervalMinutes  = 0            # Hourly: minutes
        ScheduleText     = ''
        RunTime          = [timespan]::Zero
        FirstRun         = $null        # SSRS StartDateTime (start of the schedule)
        NextRun          = $null        # first actual run (Weekly/Monthly: the first matching day)
        StartIsFuture    = $false
        OncePast         = $false
        ScheduleXml      = ''
        Mode             = 'Create'     # Create | Overwrite
        ExistingId       = ''
        HasDiff          = $false
        DiffText         = ''
        SubscriptionID   = ''
        Owner            = ''
        Message          = ''
        Warnings         = (New-Object System.Collections.Generic.List[string])
        ExtSettings      = $null
        ParamValues      = $null
    }
}

function Set-PlanFromRow {
    # Validates and normalizes an Excel row. Throws on invalid input.
    param($Plan, $Row)

    $Plan.SubscriptionName = ConvertTo-CleanText (Get-Cell $Row 'SubscriptionName')
    $path                  = ConvertTo-CleanText (Get-Cell $Row 'ReportPath')
    # Set the path first so the CSV shows which row failed even if validation throws
    $Plan.ReportPath = $path
    if (-not $Plan.SubscriptionName) { throw 'SubscriptionName is empty.' }
    if (-not $path)                  { throw 'ReportPath is empty.' }
    if ($path -match '^https?://')   { throw ("ReportPath must be an SSRS path, not a URL. Example: /Sales/Daily Sales  (value: {0})" -f $path) }
    $path = $path.Replace('\', '/').TrimEnd('/')
    if (-not $path.StartsWith('/')) {
        $path = '/' + $path
        Add-PlanWarning $Plan ("ReportPath did not start with '/'; added: {0}" -f $path)
    }
    $Plan.ReportPath = $path

    # --- Schedule type and detail ---
    $Plan.ScheduleType = ConvertTo-ScheduleType (Get-Cell $Row 'ScheduleType')
    $rawDetail = Get-Cell $Row 'ScheduleDetail'
    if (($rawDetail -is [double] -or $rawDetail -is [decimal]) -and ([double]$rawDetail -ne [math]::Floor([double]$rawDetail))) {
        # With a comma decimal separator, Excel stores "1,15" as the decimal number 1.15
        throw ("Excel stored ScheduleDetail as a decimal number ({0}). Separate the values with ';' (e.g. 1;15) or format the cell as Text." -f $rawDetail)
    }
    $detail = ConvertTo-CleanText $rawDetail
    switch ($Plan.ScheduleType) {
        'Daily'   { if ($detail) { throw ("ScheduleDetail must be empty for Daily (value: '{0}'). For specific days use ScheduleType=Weekly." -f $detail) } }
        'Once'    { if ($detail) { throw ("ScheduleDetail must be empty for Once (value: '{0}'). The date and time come from StartDate + RunTime." -f $detail) } }
        'Weekly'  { $Plan.WeekDays = @(ConvertTo-WeekDayList $detail) }
        'Monthly' {
            $Plan.MonthDays = @(ConvertTo-MonthDayList $detail)
            $late = @($Plan.MonthDays | Where-Object { $_ -gt 28 })
            if ($late.Count -gt 0) { Add-PlanWarning $Plan ('Day of month {0}: the subscription does not run in months without this day (e.g. February).' -f ($late -join ',')) }
        }
        'Hourly'  { $Plan.IntervalMinutes = ConvertTo-HourInterval $detail }
    }

    # --- Start ---
    $runTime   = ConvertTo-TimeOfDay (Get-Cell $Row 'RunTime')
    $startDate = ConvertTo-DateOnly  (Get-Cell $Row 'StartDate')
    $now       = Get-Date
    $Plan.RunTime       = $runTime
    $Plan.StartIsFuture = ($null -ne $startDate) -and ($startDate.Add($runTime) -gt $now)
    if ($null -eq $startDate) {
        $startDate = $now.Date
        Write-RunLog DEBUG 'StartDate is empty -> using today.'
    }
    $first = $startDate.Add($runTime)
    if ($first -le $now) {
        $original = $first.ToString('yyyy-MM-dd HH:mm', $script:Inv)
        if ($Plan.ScheduleType -eq 'Once') {
            # No error here: if a subscription with this name already exists (a Once that already ran),
            # the row must end as SKIP/UNCHANGED. A new create or overwrite is rejected in the pre-check.
            $Plan.OncePast = $true
            Write-RunLog INFO ('The Once time is in the past ({0}). An existing subscription is left alone; a new one cannot be created.' -f $original)
        } else {
            if ($Plan.ScheduleType -eq 'Hourly') {
                # Hourly: the next slot aligned to RunTime
                $k     = [math]::Ceiling(($now - $first).TotalMinutes / $Plan.IntervalMinutes)
                $first = $first.AddMinutes($k * $Plan.IntervalMinutes)
                if ($first -le $now) { $first = $first.AddMinutes($Plan.IntervalMinutes) }
            } else {
                $first = $now.Date.Add($runTime)
                if ($first -le $now) { $first = $first.AddDays(1) }
            }
            Write-RunLog INFO ('The start is in the past ({0}); the schedule now starts at {1}' -f $original, $first.ToString('yyyy-MM-dd HH:mm', $script:Inv))
        }
    }
    $Plan.FirstRun     = $first
    $Plan.ScheduleXml  = New-ScheduleXml $Plan
    $Plan.NextRun      = Get-NextRun $Plan
    $Plan.ScheduleText = Format-PlanSchedule $Plan

    # --- Recipients ---
    $Plan.To  = ConvertTo-EmailList (Get-Cell $Row 'To')  'To'  $true
    $Plan.Cc  = ConvertTo-EmailList (Get-Cell $Row 'Cc')  'Cc'  $false
    $Plan.Bcc = ConvertTo-EmailList (Get-Cell $Row 'Bcc') 'Bcc' $false

    # --- Content ---
    $Plan.Subject = ConvertTo-CleanText (Get-Cell $Row 'Subject')
    if (-not $Plan.Subject) {
        $Plan.Subject = '@ReportName was executed at @ExecutionTime'
        Add-PlanWarning $Plan ("Subject is empty; the SSRS default is used: '{0}'" -f $Plan.Subject)
    }
    $fmt = (ConvertTo-CleanText (Get-Cell $Row 'RenderFormat')).ToUpperInvariant()
    if (-not $fmt) { throw ('RenderFormat is empty. Valid values: {0}' -f ($AllowedRenderFormats -join ', ')) }
    if ($AllowedRenderFormats -notcontains $fmt) { throw ("RenderFormat='{0}' is not valid. Valid values: {1}" -f $fmt, ($AllowedRenderFormats -join ', ')) }
    $Plan.RenderFormat = $fmt
    $Plan.Comment      = ConvertTo-CleanText (Get-Cell $Row 'Comment')
    $Plan.IncludeLink  = ConvertTo-Bool (Get-Cell $Row 'IncludeLink') $false 'IncludeLink'

    # --- Report parameters (empty = the report defaults) ---
    $Plan.Parameters     = ConvertTo-ParameterList (Get-Cell $Row 'Parameters')
    $Plan.ParametersText = Format-ParameterList $Plan.Parameters

    # --- Owner ---
    $Plan.ExpectedOwner = ConvertTo-CleanText (Get-Cell $Row 'ExpectedOwner')
    if ($Plan.ExpectedOwner -and -not [string]::Equals($Plan.ExpectedOwner, $script:CurrentUser, $script:IgnoreCase)) {
        $msg = ('Owner mismatch. Excel ExpectedOwner={0}, running as={1}. The owner of a subscription is the account that runs the script.' -f $Plan.ExpectedOwner, $script:CurrentUser)
        if ($AllowOwnerMismatch) { Add-PlanWarning $Plan ($msg + ' (continuing because of -AllowOwnerMismatch)') }
        else { throw ($msg + ' Run the script as that account or use -AllowOwnerMismatch.') }
    }
}

function Write-PlanSummary {
    param($Plan)
    Write-RunLog INFO ('  Report      : {0}' -f $Plan.ReportPath)
    Write-RunLog INFO ('  Subscription: {0}' -f $Plan.SubscriptionName)
    Write-RunLog INFO ('  To          : {0}' -f $Plan.To)
    if ($Plan.Cc)  { Write-RunLog INFO ('  Cc          : {0}' -f $Plan.Cc) }
    if ($Plan.Bcc) { Write-RunLog INFO ('  Bcc         : {0}' -f $Plan.Bcc) }
    Write-RunLog INFO ('  Subject     : {0}' -f $Plan.Subject)
    Write-RunLog INFO ('  Format/Link : {0} / IncludeLink={1}' -f $Plan.RenderFormat, $Plan.IncludeLink)
    Write-RunLog INFO ('  Parameters  : {0}' -f $(if ($Plan.ParametersText) { $Plan.ParametersText } else { '(none - report defaults)' }))
    Write-RunLog INFO ('  Schedule    : {0}' -f $Plan.ScheduleText)
    if ($Plan.OncePast)          { Write-RunLog INFO ('  First run   : (in the past) {0}' -f $Plan.FirstRun.ToString('yyyy-MM-dd HH:mm', $script:Inv)) }
    elseif ($null -ne $Plan.NextRun) { Write-RunLog INFO ('  First run   : {0} ({1}, server time)' -f $Plan.NextRun.ToString('yyyy-MM-dd HH:mm', $script:Inv), $script:DayNames[$Plan.NextRun.DayOfWeek.ToString()]) }
}

# ===========================================================================
# SSRS
# ===========================================================================
function Get-SameNameSubscription {
    # Logs all subscriptions of the report and returns those whose Description (= subscription name) matches
    param($Plan, [switch]$Quiet)
    $all = @($script:Proxy.ListSubscriptions($Plan.ReportPath) | Where-Object { $null -ne $_ })
    if (-not $Quiet) {
        Write-RunLog INFO ('Existing subscriptions on the report: {0}' -f $all.Count)
        foreach ($s in $all) {
            Write-RunLog DEBUG ("  - ID={0} | '{1}' | Owner={2} | Status={3} | LastExecuted={4}" -f $s.SubscriptionID, $s.Description, $s.Owner, $s.Status, $s.LastExecuted)
        }
    }
    return @($all | Where-Object { [string]::Equals((ConvertTo-CleanText $_.Description), $Plan.SubscriptionName, $script:IgnoreCase) })
}

function Get-SubscriptionDetail {
    # Reads all settings of a subscription from SSRS
    param([string]$SubscriptionID)
    $ext = $null; $desc = $null; $active = $null; $status = $null; $evt = $null; $match = $null; $vals = $null
    $owner = $script:Proxy.GetSubscriptionProperties($SubscriptionID, [ref]$ext, [ref]$desc, [ref]$active, [ref]$status, [ref]$evt, [ref]$match, [ref]$vals)
    $pmap = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($pv in @($vals | Where-Object { $null -ne $_ })) {
        $nm = [string]$pv.Name
        if (-not $pmap.ContainsKey($nm)) { $pmap[$nm] = New-Object System.Collections.Generic.List[string] }
        $pmap[$nm].Add([string]$pv.Value)
    }
    return [pscustomobject]@{
        Owner       = [string]$owner
        Extension   = $(if ($null -ne $ext) { [string]$ext.Extension } else { '' })
        Values      = (Get-ExtensionValueMap $ext)
        Description = [string]$desc
        Active      = $active
        Status      = [string]$status
        EventType   = [string]$evt
        MatchData   = [string]$match
        Schedule    = (Get-ScheduleInfo ([string]$match))
        Parameters  = $pmap
    }
}

function New-Check {
    param([string]$Field, $Expected, $Actual, [bool]$Ok, [bool]$InfoOnly = $false)
    return [pscustomobject]@{ Field = $Field; Expected = [string]$Expected; Actual = [string]$Actual; Ok = $Ok; InfoOnly = $InfoOnly }
}

function Compare-PlanToSubscription {
    # Compares the values expected from the Excel row with the subscription on the server.
    #  -ExactStart  : compare the start date and time exactly (for a subscription just created or updated);
    #                 otherwise only the time of day (an existing schedule may have started in the past)
    #  -ForOverwrite: a fixed parameter value on the server that is not in the Excel row counts as a
    #                 difference (an overwrite removes it)
    param($Plan, $Details, [switch]$ExactStart, [switch]$ForOverwrite)
    $checks = New-Object System.Collections.Generic.List[object]
    $v = $Details.Values
    $ext = $script:EmailExtensionName

    $checks.Add((New-Check 'Delivery'    $ext                  $Details.Extension   (Test-TextEqual $ext $Details.Extension)))
    $checks.Add((New-Check 'EventType'   'TimedSubscription'   $Details.EventType   (Test-TextEqual 'TimedSubscription' $Details.EventType)))
    $checks.Add((New-Check 'Description' $Plan.SubscriptionName $Details.Description (Test-TextEqual $Plan.SubscriptionName $Details.Description)))
    foreach ($pair in @(@('TO', $Plan.To), @('CC', $Plan.Cc), @('BCC', $Plan.Bcc))) {
        $act = Get-DictValue $v $pair[0]
        $checks.Add((New-Check $pair[0] $pair[1] $act ([string]::Equals((Get-NormalizedEmailList $pair[1]), (Get-NormalizedEmailList $act), [StringComparison]::Ordinal))))
    }
    $checks.Add((New-Check 'Subject'      $Plan.Subject      (Get-DictValue $v 'Subject')      (Test-TextEqual $Plan.Subject (Get-DictValue $v 'Subject') -CaseSensitive)))
    $checks.Add((New-Check 'Comment'      $Plan.Comment      (Get-DictValue $v 'Comment')      (Test-TextEqual $Plan.Comment (Get-DictValue $v 'Comment') -CaseSensitive)))
    $checks.Add((New-Check 'RenderFormat' $Plan.RenderFormat (Get-DictValue $v 'RenderFormat') (Test-TextEqual $Plan.RenderFormat (Get-DictValue $v 'RenderFormat'))))
    $expLink = if ($Plan.IncludeLink) { 'True' } else { 'False' }
    $checks.Add((New-Check 'IncludeLink'  $expLink           (Get-DictValue $v 'IncludeLink')  (Test-TextEqual $expLink (Get-DictValue $v 'IncludeLink'))))

    # Schedule: type -> pattern -> start
    $s = $Details.Schedule
    $typeOk = [string]::Equals($s.Type, $Plan.ScheduleType, [StringComparison]::Ordinal)
    $patternOk = $false
    if ($typeOk) {
        switch ($Plan.ScheduleType) {
            'Daily'   { $patternOk = ($s.Interval -eq '1') }
            'Weekly'  { $patternOk = (($s.Interval -eq '1') -or ($s.Interval -eq '')) -and ((@($s.WeekDays) -join ',') -eq (@($Plan.WeekDays) -join ',')) }
            'Monthly' { $patternOk = ((@($s.MonthDays) -join ',') -eq (@($Plan.MonthDays) -join ',')) -and ($s.MonthsCount -eq 12) }
            'Hourly'  { $patternOk = ($s.Interval -eq [string]$Plan.IntervalMinutes) }
            'Once'    { $patternOk = $true }
        }
    }
    $startOk = $false
    if ($null -ne $s.Start) {
        if ($ExactStart -or $Plan.StartIsFuture -or $Plan.ScheduleType -eq 'Once') {
            $startOk = ($s.Start -eq $Plan.FirstRun)
        } elseif ($Plan.ScheduleType -eq 'Hourly') {
            # An existing hourly schedule may have started in the past: is it aligned to RunTime?
            $m = [int]$Plan.IntervalMinutes
            $diff = [int][math]::Round(($s.Start.TimeOfDay - $Plan.RunTime).TotalMinutes)
            $startOk = ((($diff % $m) + $m) % $m) -eq 0
        } else {
            $startOk = ($s.Start.TimeOfDay -eq $Plan.RunTime)
        }
    }
    $expText = Format-PlanSchedule $Plan
    $actText = if ($null -ne $s.Start) { Format-ScheduleText $s.Type $s.Start $s.Interval $s.WeekDays $s.MonthDays $s.MonthsCount } else { $Details.MatchData }
    $checks.Add((New-Check 'Schedule' $expText $actText ($typeOk -and $patternOk -and $startOk)))

    # Report parameters
    $pmap = $Details.Parameters
    foreach ($xp in $Plan.Parameters) {
        $act = @()
        if ($pmap.ContainsKey($xp.Name)) { $act = @($pmap[$xp.Name]) }
        $actText = if ($act.Count -gt 0) { $act -join ',' } else { '(default)' }
        $ok = [string]::Equals((Get-NormalizedValueSet $xp.Values), (Get-NormalizedValueSet $act), [StringComparison]::Ordinal)
        $checks.Add((New-Check ('Parameter ' + $xp.Name) ($xp.Values -join ',') $actText $ok))
    }
    foreach ($name in @($pmap.Keys)) {
        $inExcel = @($Plan.Parameters | Where-Object { [string]::Equals($_.Name, $name, $script:IgnoreCase) }).Count -gt 0
        if (-not $inExcel) {
            $checks.Add((New-Check ('Parameter ' + $name) '(default)' (@($pmap[$name]) -join ',') $false (-not $ForOverwrite)))
        }
    }
    # Return an array: it streams item by item. With ',$list' Where-Object would get the whole list as one item.
    return $checks.ToArray()
}

function Resolve-ExistingSubscription {
    # Decides what to do when a subscription with the same name exists.
    # $true  -> the row continues (create, or overwrite)
    # $false -> decided (SKIP_EXISTING / UNCHANGED)
    param($Plan)
    $same = @(Get-SameNameSubscription $Plan)
    if ($same.Count -eq 0) {
        Write-RunLog OK 'No subscription with this name -> it will be created.'
        $Plan.Mode = 'Create'
        return $true
    }

    $ids    = ($same | ForEach-Object { $_.SubscriptionID }) -join ';'
    $owners = ($same | ForEach-Object { $_.Owner }) -join ';'
    $Plan.SubscriptionID = $ids
    $Plan.Owner          = $owners

    if (-not $Overwrite) {
        $Plan.Action  = 'SKIP_EXISTING'
        $Plan.Message = 'A subscription with this name already exists; left alone.'
        Write-RunLog WARN ('A subscription with this name already exists -> SKIP (not deleted, not updated). ID={0} Owner={1}' -f $ids, $owners)
        if ($same.Count -eq 1 -and -not $same[0].IsDataDriven) {
            try {
                $d = Get-SubscriptionDetail ([string]$same[0].SubscriptionID)
                $diffs = @(Compare-PlanToSubscription $Plan $d -ForOverwrite | Where-Object { -not $_.Ok -and -not $_.InfoOnly })
                foreach ($c in $diffs) {
                    Add-PlanWarning $Plan ("Differs from Excel (use -Overwrite to update) {0}: SSRS='{1}' Excel='{2}'" -f $c.Field, $c.Actual, $c.Expected)
                }
                if ($diffs.Count -gt 0) { $Plan.HasDiff = $true; $Plan.Message += (' Fields that differ from Excel: {0}.' -f $diffs.Count) }
                else { Write-RunLog INFO 'The existing subscription matches the Excel row.' }
            } catch {
                Write-RunLog WARN ('Could not read the existing subscription: {0}' -f (Get-ShortError $_.Exception))
            }
        }
        return $false
    }

    # ---- -Overwrite ----
    if ($same.Count -gt 1) {
        throw ("{0} subscriptions have this name (ID={1}). It is unclear which one to update; delete the extra ones in the web portal." -f $same.Count, $ids)
    }
    $ex = $same[0]
    if ($ex.IsDataDriven) { throw ("The existing subscription is data-driven (ID={0}); the script does not change data-driven subscriptions." -f $ex.SubscriptionID) }
    if (-not [string]::Equals([string]$ex.Owner, $script:CurrentUser, $script:IgnoreCase)) {
        $msg = ('The existing subscription belongs to another user (Owner={0}).' -f $ex.Owner)
        if ($AllowOwnerMismatch) { Add-PlanWarning $Plan ($msg + ' It will be updated because of -AllowOwnerMismatch.') }
        else { throw ($msg + ' Use -AllowOwnerMismatch to update it.') }
    }
    $Plan.ExistingId = [string]$ex.SubscriptionID
    $d = Get-SubscriptionDetail $Plan.ExistingId
    Write-RunLog DEBUG ('  Existing MatchData : {0}' -f $d.MatchData)
    Write-RunLog DEBUG ('  Existing parameters: {0}' -f (Format-ParamMap $d.Parameters))
    $diffs = @(Compare-PlanToSubscription $Plan $d -ForOverwrite | Where-Object { -not $_.Ok -and -not $_.InfoOnly })
    if ($diffs.Count -eq 0) {
        $Plan.Action  = 'UNCHANGED'
        $Plan.Message = 'The existing subscription matches the Excel row; no update needed.'
        Write-RunLog OK ('A subscription with this name exists and matches the Excel row -> UNCHANGED (ID={0})' -f $Plan.ExistingId)
        return $false
    }
    Write-RunLog WARN ('A subscription with this name exists, {0} field(s) differ -> WILL BE UPDATED (ID={1})' -f $diffs.Count, $Plan.ExistingId)
    foreach ($c in $diffs) { Write-RunLog WARN ("  [DIFF] {0}: SSRS='{1}' -> Excel='{2}'" -f $c.Field, $c.Actual, $c.Expected) }
    $Plan.Mode     = 'Overwrite'
    $Plan.HasDiff  = $true
    $Plan.DiffText = ($diffs | ForEach-Object { $_.Field }) -join ', '
    return $true
}

function Test-ReportDataSource {
    # Data sources of a subscribed report must use stored credentials ('Store') or none ('None'). Warnings only.
    param($Plan)
    try {
        $dss = @($script:Proxy.GetItemDataSources($Plan.ReportPath) | Where-Object { $null -ne $_ })
    } catch {
        Add-PlanWarning $Plan ('Could not check the data sources: {0}' -f (Get-ShortError $_.Exception))
        return
    }
    Write-RunLog INFO ('Data sources: {0}' -f $dss.Count)
    foreach ($ds in $dss) {
        try {
            $item = $ds.Item
            $kind = $item.GetType().Name
            $def  = $null
            $src  = 'embedded'
            if ($kind -eq 'DataSourceReference') {
                $src = 'shared ' + $item.Reference
                $def = $script:Proxy.GetDataSourceContents($item.Reference)
            } elseif ($kind -eq 'DataSourceDefinition') {
                $def = $item
            } else {
                Add-PlanWarning $Plan ("Data source '{0}' has an invalid reference ({1}). The shared data source may have been deleted or moved; the subscription will not run." -f $ds.Name, $kind)
                continue
            }
            $cr = [string]$def.CredentialRetrieval
            Write-RunLog DEBUG ('  - {0} ({1}) Extension={2} CredentialRetrieval={3} Enabled={4}' -f $ds.Name, $src, $def.Extension, $cr, $def.Enabled)
            if (@('Integrated','Prompt') -contains $cr) {
                Add-PlanWarning $Plan ("Data source '{0}' uses CredentialRetrieval={1}. Subscriptions need stored credentials (Store) or None; SSRS will probably reject the subscription." -f $ds.Name, $cr)
            } else {
                Write-RunLog OK ("  Data source '{0}' credentials are suitable ({1})" -f $ds.Name, $cr)
            }
        } catch {
            Add-PlanWarning $Plan ("Could not read data source '{0}' (permissions?): {1}" -f $ds.Name, (Get-ShortError $_.Exception))
        }
    }
}

function New-ReportParameterValueArray {
    # Excel Parameters -> SSRS ParameterValue[] (one entry per value for multi-value parameters). $null if none.
    param($Plan)
    $n = 0
    foreach ($xp in $Plan.Parameters) { $n += $xp.Values.Count }
    if ($n -eq 0) { return $null }
    $arr = New-Object ('{0}.ParameterValue[]' -f $script:Ns) $n
    $i = 0
    foreach ($xp in $Plan.Parameters) {
        foreach ($val in $xp.Values) {
            $arr[$i] = New-Object ('{0}.ParameterValue' -f $script:Ns) -Property @{ Name = $xp.Name; Value = $val }
            $i++
        }
    }
    return ,$arr
}

function Write-ValidValueList {
    # Lists the values SSRS accepts for a missing or invalid parameter (to copy into the Excel file)
    param($ItemParameter)
    $vv = @($ItemParameter.ValidValues | Where-Object { $null -ne $_ })
    if ($vv.Count -eq 0) {
        Write-RunLog INFO ("    '{0}' has no list of valid values (any value can be entered)." -f $ItemParameter.Name)
        return
    }
    $max = 30
    Write-RunLog INFO ("    Valid values for '{0}' (first {1} of {2}). Put the value on the left into the Excel file:" -f $ItemParameter.Name, [Math]::Min($max, $vv.Count), $vv.Count)
    foreach ($x in ($vv | Select-Object -First $max)) { Write-RunLog INFO ('      {0}   ({1})' -f $x.Value, $x.Label) }
}

function Test-ReportParameter {
    # Checks the report parameters. Throws on definite errors (the row becomes ERROR and blocks -Apply).
    #  1) Static list (ForRendering=false): names, types, defaults
    #  2) Live validation (ForRendering=true + the Excel values): SSRS evaluates defaults and valid values
    #     the way the web portal does and returns a state for each parameter.
    #  If the live validation fails, the static rule applies: required + no default + not in Excel = ERROR.
    param($Plan)
    # HistoryID needs a real null: PowerShell passes $null to a string parameter as "", and SSRS reads ""
    # as a snapshot date and fails with rsParameterTypeMismatch.
    $noHistory = [System.Management.Automation.Language.NullString]::Value
    try {
        $params = @($script:Proxy.GetItemParameters($Plan.ReportPath, $noHistory, $false, $null, $null) | Where-Object { $null -ne $_ })
    } catch {
        if ($Plan.Parameters.Count -gt 0) { throw ('Could not read the report parameters, so Parameters from the Excel row cannot be validated: {0}' -f (Get-ShortError $_.Exception)) }
        Add-PlanWarning $Plan ('Could not read the report parameters (check skipped): {0}' -f (Get-ShortError $_.Exception))
        return
    }
    if ($params.Count -eq 0) {
        if ($Plan.Parameters.Count -gt 0) { throw ('The report has no parameters, but Parameters is filled in: {0}' -f $Plan.ParametersText) }
        Write-RunLog INFO 'The report has no parameters.'
        return
    }
    $names = @($params | ForEach-Object { [string]$_.Name })
    Write-RunLog INFO ('Report parameters ({0}): {1}' -f $params.Count, ($names -join ', '))
    $errs = New-Object System.Collections.Generic.List[string]

    # a) Do the names from Excel exist in the report? Several values for a single-value parameter?
    foreach ($xp in $Plan.Parameters) {
        $rp = @($params | Where-Object { [string]::Equals([string]$_.Name, $xp.Name, $script:IgnoreCase) })[0]
        if ($null -eq $rp) {
            $errs.Add(("Parameter '{0}' from the Excel row is not in the report (report parameters: {1})" -f $xp.Name, ($names -join ', ')))
            continue
        }
        $xp.Name = [string]$rp.Name   # use the spelling from SSRS
        if (-not $rp.MultiValue -and $xp.Values.Count -gt 1) {
            $errs.Add(("Parameter '{0}' takes a single value but {1} values were given: {2}" -f $rp.Name, $xp.Values.Count, ($xp.Values -join ',')))
        }
        if ((Get-ParamTypeName $rp) -eq 'DateTime') {
            Add-PlanWarning $Plan ("Parameter '{0}' gets a fixed date ({1}). The subscription uses the SAME date on every run; leave it out to use the report default." -f $rp.Name, ($xp.Values -join ','))
        }
    }
    if ($errs.Count -gt 0) { throw ($errs -join ' | ') }
    $Plan.ParametersText = Format-ParameterList $Plan.Parameters

    # b) Static information
    foreach ($p in $params) {
        $defaults  = @($p.DefaultValues | Where-Object { $null -ne $_ })
        $defText   = if ($p.DefaultValuesQueryBased) { '(from query)' } elseif ($defaults.Count -gt 0) { $defaults -join ',' } else { '(NONE)' }
        $given     = @($Plan.Parameters | Where-Object { $_.Name -ceq [string]$p.Name })
        $givenText = if ($given.Count -gt 0) { $given[0].Values -join ',' } else { '-' }
        Write-RunLog DEBUG ('  - {0} Type={1} Nullable={2} MultiValue={3} Default={4} Excel={5}' -f $p.Name, (Get-ParamTypeName $p), $p.Nullable, $p.MultiValue, $defText, $givenText)
    }

    # c) Live validation
    $values    = New-ReportParameterValueArray $Plan
    $evaluated = $null
    try {
        $evaluated = @($script:Proxy.GetItemParameters($Plan.ReportPath, $noHistory, $true, $values, $null) | Where-Object { $null -ne $_ })
        Write-RunLog DEBUG 'Parameters validated live on SSRS (ForRendering=true).'
    } catch {
        $code = Get-SoapErrorCode $_.Exception
        if ($Plan.Parameters.Count -gt 0 -and $script:ParamValueErrorCodes -contains $code) {
            throw ('SSRS rejected a parameter value from the Excel row: {0}' -f (Get-ShortError $_.Exception))
        }
        Add-PlanWarning $Plan ('Could not validate the parameters live on SSRS; using the static check instead: {0}' -f (Get-ShortError $_.Exception))
        $evaluated = $null
    }

    foreach ($p in $params) {
        $name  = [string]$p.Name
        $given = @($Plan.Parameters | Where-Object { $_.Name -ceq $name }).Count -gt 0
        $ep    = $null
        if ($null -ne $evaluated) { $ep = @($evaluated | Where-Object { [string]$_.Name -ceq $name })[0] }
        $state = Get-ParamState $ep
        if ($state) {
            $cur   = @($ep.DefaultValues | Where-Object { $null -ne $_ }) -join ','
            if ($ep.PSObject.Properties['ErrorMessage'] -and $ep.ErrorMessage) { Write-RunLog DEBUG ('  {0} SSRS message: {1}' -f $name, $ep.ErrorMessage) }
            if ($state -eq 'HasValidValue') {
                Write-RunLog OK ('  [OK] {0} = {1}  ({2})' -f $name, $(if ($cur) { $cur } else { 'NULL/empty' }), $(if ($given) { 'Excel' } else { 'report default' }))
            } elseif ($state -eq 'MissingValidValue') {
                if ($given) { $errs.Add(("Parameter '{0}': the value from the Excel row is not valid on SSRS." -f $name)) }
                else        { $errs.Add(("Parameter '{0}' is required and has no default; add it to the Parameters column (example: {0}=<value>)." -f $name)) }
                Write-RunLog ERROR ('  [MISSING/INVALID] {0}' -f $name)
                Write-ValidValueList $ep
            } elseif ($state -eq 'HasOutstandingDependencies') {
                Add-PlanWarning $Plan ("Parameter '{0}' depends on a parameter that has no valid value yet: {1}" -f $name, (@($ep.Dependencies) -join ', '))
            } else {
                Add-PlanWarning $Plan ("Parameter '{0}' state: {1} (SSRS could not evaluate its values)." -f $name, $state)
            }
        } else {
            Write-RunLog DEBUG ('  {0}: no live state from SSRS, applying the static rule.' -f $name)
            $defaults   = @($p.DefaultValues | Where-Object { $null -ne $_ })
            $hasDefault = ($defaults.Count -gt 0) -or [bool]$p.DefaultValuesQueryBased
            if (-not $given -and -not $hasDefault) {
                if ($p.Nullable) { Add-PlanWarning $Plan ("Parameter '{0}' has no default (it accepts NULL)." -f $name) }
                else { $errs.Add(("Parameter '{0}' is required and has no default; add it to the Parameters column (example: {0}=<value>)." -f $name)) }
            }
        }
    }
    if ($errs.Count -gt 0) { throw ($errs -join ' | ') }
    $Plan.ParamValues = $values
}

function New-EmailExtensionSetting {
    # Same structure as New-RsSubscription in ReportingServicesTools; empty fields are not sent.
    param($Plan)
    $values = [ordered]@{
        TO            = $Plan.To
        CC            = $Plan.Cc
        BCC           = $Plan.Bcc
        Subject       = $Plan.Subject
        Comment       = $Plan.Comment
        RenderFormat  = $Plan.RenderFormat
        IncludeReport = 'True'
        IncludeLink   = $(if ($Plan.IncludeLink) { 'True' } else { 'False' })
        Priority      = 'NORMAL'
    }
    $pvs = @()
    foreach ($k in @($values.Keys)) {
        $val = [string]$values[$k]
        if ([string]::IsNullOrEmpty($val)) { continue }
        $pvs += New-Object ('{0}.ParameterValue' -f $script:Ns) -Property @{ Name = $k; Value = $val }
        Write-RunLog DEBUG ('  Delivery setting {0} = {1}' -f $k, $val)
    }
    return (New-Object ('{0}.ExtensionSettings' -f $script:Ns) -Property @{ Extension = $script:EmailExtensionName; ParameterValues = $pvs })
}

function Confirm-Subscription {
    # Reads a created or updated subscription back from SSRS and compares it with the Excel row
    param($Plan, [switch]$IsOverwrite)
    $prefix = if ($IsOverwrite) { 'OVERWRITTEN' } else { 'CREATED' }
    $verb   = if ($IsOverwrite) { 'Updated' } else { 'Created' }
    Write-RunLog INFO 'Verification: reading the subscription back from SSRS...'
    try {
        $d = Get-SubscriptionDetail $Plan.SubscriptionID
    } catch {
        $Plan.Action  = $prefix + '_UNVERIFIED'
        $Plan.Message = $verb + ' but could not be read back: ' + (Get-ShortError $_.Exception)
        Write-ErrorDetail $_ ('Verification read failed ({0}; check the subscription in the web portal)' -f $verb)
        return
    }

    $Plan.Owner = $d.Owner
    $diffs = 0
    if ($IsOverwrite) {
        Write-RunLog INFO ('  [INFO] Owner = {0}' -f $d.Owner)
    } elseif ([string]::Equals($d.Owner, $script:CurrentUser, $script:IgnoreCase)) {
        Write-RunLog OK ('  [OK]   Owner = {0}' -f $d.Owner)
    } else {
        $diffs++
        Write-RunLog WARN ("  [DIFF] Owner: expected='{0}' SSRS='{1}'" -f $script:CurrentUser, $d.Owner)
    }

    foreach ($c in (Compare-PlanToSubscription $Plan $d -ExactStart)) {
        if ($c.Ok)           { Write-RunLog OK   ('  [OK]   {0} = {1}' -f $c.Field, $c.Actual) }
        elseif ($c.InfoOnly) { Write-RunLog INFO ("  [INFO] {0} = {1} (not in Excel)" -f $c.Field, $c.Actual) }
        else                 { $diffs++; Write-RunLog WARN ("  [DIFF] {0}: expected='{1}' SSRS='{2}'" -f $c.Field, $c.Expected, $c.Actual) }
    }
    Write-RunLog DEBUG ('  SSRS MatchData : {0}' -f $d.MatchData)
    Write-RunLog DEBUG ('  SSRS parameters: {0}' -f (Format-ParamMap $d.Parameters))
    Write-RunLog INFO  ('  SSRS Status    : {0}' -f $d.Status)

    if ($null -ne $d.Active) {
        foreach ($prop in $d.Active.PSObject.Properties) {
            if ($prop.Name -notlike '*Specified' -and $prop.Value -eq $true) {
                $diffs++
                Write-RunLog WARN ('  [WARN] ActiveState.{0} = True (the subscription may not run)' -f $prop.Name)
            }
        }
    }

    if ($diffs -eq 0) {
        $Plan.Message = $verb + ' and verified.'
        Write-RunLog OK 'Verification PASSED: the subscription on SSRS matches the Excel row.'
    } else {
        $Plan.Action  = $prefix + '_MISMATCH'
        $Plan.Message = ('{0} but {1} field(s) differ; see the log.' -f $verb, $diffs)
        Write-RunLog WARN ('Verification: {0} difference(s). Check the subscription in the web portal.' -f $diffs)
    }
}

# ===========================================================================
# SUMMARY
# ===========================================================================
function Get-ActionCount {
    param([string[]]$Actions)
    return @($script:Plans | Where-Object { $Actions -contains $_.Action }).Count
}

function Write-Summary {
    Write-Step 7 'Summary'
    if ($script:Plans.Count -eq 0) {
        Write-RunLog WARN 'No rows were processed.'
    } else {
        foreach ($g in ($script:Plans | Group-Object Action | Sort-Object Name)) {
            Write-RunLog INFO ('  {0,-24} : {1}' -f $g.Name, $g.Count)
        }
        $table = $script:Plans | Select-Object ExcelRow, Action, ReportPath, SubscriptionName, SubscriptionID, Message |
                 Format-Table -AutoSize -Wrap | Out-String -Width 300
        foreach ($l in ($table -split "`r?`n")) { if ($l.Trim()) { Write-RunLog INFO $l } }

        $script:Plans | Select-Object ExcelRow, Action, ReportPath, SubscriptionName, To, Cc, Bcc, Subject, RenderFormat, IncludeLink,
                     @{ N = 'Parameters'; E = { $_.ParametersText } },
                     @{ N = 'Schedule';   E = { $_.ScheduleText } },
                     @{ N = 'FirstRun';   E = { if ($_.NextRun) { $_.NextRun.ToString('yyyy-MM-dd HH:mm', $script:Inv) } } },
                     SubscriptionID, Owner, Message,
                     @{ N = 'Warnings';   E = { $_.Warnings -join ' | ' } } |
                 Export-Csv -LiteralPath $script:CsvFile -NoTypeInformation -Encoding UTF8 -UseCulture
        Write-RunLog INFO ('Result CSV: {0}' -f $script:CsvFile)
    }
    Write-RunLog INFO ('Log file  : {0}' -f $script:LogFile)

    if ($script:Fatal -or (Get-ActionCount $script:BadActions) -gt 0) { $script:ExitCode = 1 }

    $applyCmd = $script:CommandBase + ' -Apply'
    if ($Overwrite)          { $applyCmd += ' -Overwrite' }
    if ($AllowOwnerMismatch) { $applyCmd += ' -AllowOwnerMismatch' }

    if ($script:Fatal) {
        Write-RunLog ERROR 'The script stopped with a FATAL ERROR. See the ERROR/DEBUG lines above.'
    } elseif (-not $Apply) {
        Write-RunLog STEP 'PREVIEW finished. NOTHING was changed on SSRS.'
        if ((Get-ActionCount 'ERROR') -gt 0) {
            Write-RunLog ERROR 'Some rows have errors. Fix them (or set Enabled=false) and run the preview again.'
        } elseif ((Get-ActionCount @('PREVIEW_CREATE','PREVIEW_OVERWRITE')) -gt 0) {
            Write-RunLog STEP ('To apply, run in the same window:  {0}' -f $applyCmd)
            Write-RunLog INFO ('(In a new window, first: Set-Location "{0}" ; Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force)' -f $PSScriptRoot)
        } else {
            Write-RunLog STEP 'Nothing to change.'
        }
    } else {
        Write-RunLog STEP ('APPLY finished. Created: {0} | Updated: {1} | Verification issues: {2} | Unchanged: {3} | Already existed (SKIP): {4} | Errors: {5} | Blocked: {6}' -f `
            (Get-ActionCount 'CREATED'), (Get-ActionCount 'OVERWRITTEN'),
            (Get-ActionCount @('CREATED_MISMATCH','CREATED_UNVERIFIED','OVERWRITTEN_MISMATCH','OVERWRITTEN_UNVERIFIED')),
            (Get-ActionCount 'UNCHANGED'), (Get-ActionCount 'SKIP_EXISTING'), (Get-ActionCount 'ERROR'), (Get-ActionCount 'BLOCKED'))
    }
    $skipDiff = @($script:Plans | Where-Object { $_.Action -eq 'SKIP_EXISTING' -and $_.HasDiff }).Count
    if ($skipDiff -gt 0 -and -not $Overwrite) {
        Write-RunLog WARN ('{0} existing subscription(s) differ from the Excel file but were skipped. To update them, preview first: {1} -Overwrite' -f $skipDiff, $script:CommandBase)
    }
    Write-RunLog INFO ('Exit code: {0}' -f $script:ExitCode)
}

# ===========================================================================
# MAIN
# ===========================================================================
try {
    # -------------------------------------------------------------------
    Write-Step 1 'Environment'
    # -------------------------------------------------------------------
    if ($Apply) { Write-RunLog WARN 'Mode         : APPLY -> changes WILL be made on SSRS if the pre-checks are clean' }
    else        { Write-RunLog INFO 'Mode         : PREVIEW -> NO changes on SSRS (read-only)' }
    if ($Overwrite) {
        Write-RunLog WARN ('Overwrite    : ON -> subscriptions with the same name that differ from the Excel file are updated{0}' -f $(if ($Apply) { '' } else { ' (the preview only shows them)' }))
    } else {
        Write-RunLog INFO 'Overwrite    : off -> subscriptions with the same name are skipped (SKIP)'
    }
    Write-RunLog INFO ('Run as       : {0} @ {1}' -f $script:CurrentUser, $env:COMPUTERNAME)
    Write-RunLog INFO ('PowerShell   : {0} ({1})' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)
    Write-RunLog INFO ('Culture / TZ : {0} / {1}' -f (Get-Culture).Name, [TimeZoneInfo]::Local.Id)
    Write-RunLog INFO ('Server       : {0}' -f $ReportServerUrl)
    Write-RunLog INFO ('Excel        : {0} (sheet: {1})' -f $ExcelPath, $WorksheetName)
    Write-RunLog INFO ('Log / CSV    : {0}' -f $LogDir)
    if ($AllowOwnerMismatch) { Write-RunLog WARN '-AllowOwnerMismatch is on: an owner mismatch is a warning instead of an error.' }
    Import-RequiredModule 'ImportExcel'

    # -------------------------------------------------------------------
    Write-Step 2 'Reading Excel'
    # -------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $ExcelPath)) { throw ('Excel file not found: {0}' -f $ExcelPath) }
    try {
        $rawRows = @(Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName)
    } catch {
        throw ("Could not read the Excel file (sheet '{0}' missing, or the file is damaged or locked): {1}" -f $WorksheetName, $_.Exception.Message)
    }
    if ($rawRows.Count -eq 0) { throw 'The Excel sheet has no rows below the header.' }
    $columns = @($rawRows[0].PSObject.Properties | ForEach-Object { $_.Name })
    Write-RunLog INFO ('Rows read (without the header, empty rows included): {0}' -f $rawRows.Count)
    Write-RunLog DEBUG ('Columns: {0}' -f ($columns -join ', '))
    $missingRequired = @($RequiredColumns | Where-Object { $columns -notcontains $_ })
    if ($missingRequired.Count -gt 0) { throw ('Required column(s) missing in the Excel sheet: {0}' -f ($missingRequired -join ', ')) }
    $missingOptional = @($OptionalColumns | Where-Object { $columns -notcontains $_ })
    if ($missingOptional.Count -gt 0) { Write-RunLog WARN ('Optional column(s) missing, treated as empty: {0}' -f ($missingOptional -join ', ')) }
    $unknown = @($columns | Where-Object { ($RequiredColumns + $OptionalColumns) -notcontains $_ })
    if ($unknown.Count -gt 0) { Write-RunLog WARN ('Unknown column(s) are ignored: {0}' -f ($unknown -join ', ')) }
    Write-RunLog OK 'Column layout OK.'

    # -------------------------------------------------------------------
    Write-Step 3 'Input check (row by row)'
    # -------------------------------------------------------------------
    $emptyRows = 0
    for ($i = 0; $i -lt $rawRows.Count; $i++) {
        $row      = $rawRows[$i]
        $excelRow = $i + 2   # row 1 is the header
        $hasData  = @($columns | Where-Object { (ConvertTo-CleanText (Get-Cell $row $_)) -ne '' }).Count -gt 0
        if (-not $hasData) { $emptyRows++; continue }

        $script:RowTag = 'Row {0}' -f $excelRow
        $plan = New-Plan $excelRow
        $script:Plans.Add($plan)

        Write-RunLog STEP ('Reading row: {0} | {1}' -f (ConvertTo-CleanText (Get-Cell $row 'ReportPath')), (ConvertTo-CleanText (Get-Cell $row 'SubscriptionName')))
        foreach ($c in $columns) { Write-RunLog DEBUG ('  raw {0,-16} = {1}' -f $c, (Format-Raw (Get-Cell $row $c))) }

        try {
            $enabled = ConvertTo-Bool (Get-Cell $row 'Enabled') $false 'Enabled'
        } catch {
            $plan.Action  = 'ERROR'
            $plan.Message = $_.Exception.Message
            Write-RunLog ERROR $plan.Message
            continue
        }
        if (-not $enabled) {
            $plan.Action           = 'DISABLED'
            $plan.ReportPath       = ConvertTo-CleanText (Get-Cell $row 'ReportPath')
            $plan.SubscriptionName = ConvertTo-CleanText (Get-Cell $row 'SubscriptionName')
            $plan.Message          = 'Enabled=false (or empty), skipped.'
            Write-RunLog INFO $plan.Message
            continue
        }

        try {
            Set-PlanFromRow $plan $row
            Write-PlanSummary $plan
            Write-RunLog DEBUG ('  Schedule XML: {0}' -f $plan.ScheduleXml)
            Write-RunLog OK 'Input valid.'
        } catch {
            $plan.Action  = 'ERROR'
            $plan.Message = $_.Exception.Message
            Write-RunLog ERROR ('Input error: {0}' -f $plan.Message)
        }
    }
    $script:RowTag = ''
    Write-RunLog INFO ('Empty rows (ignored): {0}' -f $emptyRows)

    # Is the same ReportPath + SubscriptionName repeated in the Excel file?
    $seen = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($p in @($script:Plans | Where-Object { $_.Action -eq 'PENDING' })) {
        $key = $p.ReportPath + '|' + $p.SubscriptionName
        if ($seen.ContainsKey($key)) {
            $firstPlan = $seen[$key]
            foreach ($dup in @($firstPlan, $p)) {
                $dup.Action  = 'ERROR'
                $dup.Message = ('Duplicate in the Excel file: same ReportPath + SubscriptionName (rows {0} and {1}).' -f $firstPlan.ExcelRow, $p.ExcelRow)
            }
            Write-RunLog ERROR $p.Message
        } else {
            $seen[$key] = $p
        }
    }

    $pending = @($script:Plans | Where-Object { $_.Action -eq 'PENDING' })
    Write-RunLog INFO ('Valid: {0} | Errors: {1} | Disabled: {2}' -f $pending.Count, (Get-ActionCount 'ERROR'), (Get-ActionCount 'DISABLED'))

    if ($pending.Count -eq 0) {
        Write-RunLog WARN 'No valid rows to send to SSRS; skipping the SSRS steps.'
    } else {
        # -------------------------------------------------------------------
        Write-Step 4 'SSRS connection'
        # -------------------------------------------------------------------
        $baseUrl = $ReportServerUrl.Trim().TrimEnd('/')
        if ($baseUrl.EndsWith('/ReportService2010.asmx', $script:IgnoreCase)) { $baseUrl = $baseUrl.Substring(0, $baseUrl.Length - '/ReportService2010.asmx'.Length) }
        if ($baseUrl -match '/Reports(_[^/]+)?$') { Write-RunLog WARN ('{0} looks like the web portal address. The web service address usually ends with /ReportServer.' -f $baseUrl) }
        $soapUrl = $baseUrl + '/ReportService2010.asmx'
        Write-RunLog INFO ('SOAP address: {0}' -f $soapUrl)
        Write-RunLog INFO ('Sign-in     : current Windows account ({0})' -f $script:CurrentUser)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $script:Proxy = New-WebServiceProxy -Uri $soapUrl -UseDefaultCredential -ErrorAction Stop
        } catch {
            throw ("Could not connect to SSRS ({0}). Try opening the address in a browser; 401 means permissions, a timeout means network/VPN. Details: {1}" -f $soapUrl, $_.Exception.Message)
        }
        $script:Proxy.Timeout = 300000
        $script:Ns = $script:Proxy.GetType().Namespace
        Write-RunLog OK ('Proxy created ({0} ms). Proxy.Url={1}' -f $sw.ElapsedMilliseconds, $script:Proxy.Url)

        $rootType = $script:Proxy.GetItemType('/')
        Write-RunLog OK ("Authentication works (GetItemType('/') = {0})" -f $rootType)
        $si = $script:Proxy.ServerInfoHeaderValue
        if ($si) {
            foreach ($prop in $si.PSObject.Properties) {
                # Only the ReportServer* fields (MustUnderstand etc. come from SoapHeader and are not useful)
                if ($prop.Name -notlike 'ReportServer*' -or $null -eq $prop.Value -or [string]$prop.Value -eq '') { continue }
                if ($prop.Name -eq 'ReportServerTimeZoneInfo' -and $prop.Value.PSObject.Properties['Bias']) {
                    # Bias is in minutes: UTC = local time + Bias (e.g. -180 -> UTC+03:00)
                    $offset = [TimeSpan]::FromMinutes(-[int]$prop.Value.Bias)
                    Write-RunLog INFO ('Server time zone = UTC{0}{1:hh\:mm} (schedules run in this time zone)' -f $(if ($offset -lt [TimeSpan]::Zero) { '-' } else { '+' }), $offset.Duration())
                } else {
                    Write-RunLog INFO ('Server {0} = {1}' -f $prop.Name, $prop.Value)
                }
            }
        }

        $deliveryNames = $null
        try {
            $deliveryNames = @($script:Proxy.ListExtensions('Delivery') | Where-Object { $null -ne $_ } | ForEach-Object { $_.Name })
            Write-RunLog DEBUG ('Delivery extensions: {0}' -f ($deliveryNames -join ', '))
        } catch {
            Write-RunLog WARN ('Could not list the delivery extensions (continuing): {0}' -f (Get-ShortError $_.Exception))
        }
        if ($null -ne $deliveryNames) {
            if ($deliveryNames -notcontains $script:EmailExtensionName) {
                throw ("SSRS has no '{0}' delivery extension. Configure it in Report Server Configuration Manager > E-mail Settings." -f $script:EmailExtensionName)
            }
            Write-RunLog OK ("E-mail delivery extension found: '{0}'" -f $script:EmailExtensionName)
        }

        # -------------------------------------------------------------------
        Write-Step 5 'Pre-check (READ-ONLY)'
        # -------------------------------------------------------------------
        foreach ($plan in $pending) {
            $script:RowTag = 'Row {0}' -f $plan.ExcelRow
            Write-RunLog STEP ('Pre-check: {0} | {1}' -f $plan.ReportPath, $plan.SubscriptionName)
            try {
                $itemType = [string]$script:Proxy.GetItemType($plan.ReportPath)
                if (@('Report','LinkedReport') -notcontains $itemType) {
                    throw ("Report not found, or not a report (ItemType={0}). Check the path and your permissions." -f $itemType)
                }
                Write-RunLog OK ('Report found (ItemType={0})' -f $itemType)

                if (-not (Resolve-ExistingSubscription $plan)) { continue }
                if ($plan.OncePast) {
                    throw ('The Once time is in the past ({0}); a subscription cannot be created or updated with it. Move StartDate/RunTime forward or set Enabled=false.' -f $plan.FirstRun.ToString('yyyy-MM-dd HH:mm', $script:Inv))
                }

                Test-ReportDataSource $plan
                Test-ReportParameter  $plan

                # Build the delivery settings in the preview too, so type and field problems show up before APPLY
                $plan.ExtSettings = New-EmailExtensionSetting $plan

                $plan.Action = if ($plan.Mode -eq 'Overwrite') { 'READY_OVERWRITE' } else { 'READY' }
                Write-RunLog OK ('Pre-check OK ({0}). Warnings: {1}' -f $(if ($plan.Mode -eq 'Overwrite') { 'WILL BE UPDATED' } else { 'WILL BE CREATED' }), $plan.Warnings.Count)
            } catch {
                $plan.Action  = 'ERROR'
                $plan.Message = Get-ShortError $_.Exception
                Write-ErrorDetail $_ 'Pre-check error'
            }
        }
        $script:RowTag = ''

        # -------------------------------------------------------------------
        Write-Step 6 'Apply and verify'
        # -------------------------------------------------------------------
        $ready  = @($script:Plans | Where-Object { $_.Action -eq 'READY' -or $_.Action -eq 'READY_OVERWRITE' })
        $errors = @($script:Plans | Where-Object { $_.Action -eq 'ERROR' })
        Write-RunLog INFO ('To create: {0} | To update: {1} | Unchanged: {2} | Already exist (SKIP): {3} | Errors: {4}' -f `
            (Get-ActionCount 'READY'), (Get-ActionCount 'READY_OVERWRITE'), (Get-ActionCount 'UNCHANGED'), (Get-ActionCount 'SKIP_EXISTING'), $errors.Count)

        if (-not $Apply) {
            foreach ($plan in $ready) {
                if ($plan.Action -eq 'READY_OVERWRITE') {
                    $plan.Action  = 'PREVIEW_OVERWRITE'
                    $plan.Message = ('-Apply -Overwrite will update it. Fields that differ: {0}' -f $plan.DiffText)
                } else {
                    $plan.Action  = 'PREVIEW_CREATE'
                    $plan.Message = ('-Apply will create it. First run {0}' -f $plan.NextRun.ToString('yyyy-MM-dd HH:mm', $script:Inv))
                }
            }
            Write-RunLog INFO 'PREVIEW mode: apply step skipped.'
        } elseif ($errors.Count -gt 0) {
            foreach ($plan in $ready) {
                $plan.Action  = 'BLOCKED'
                $plan.Message = 'Not applied because other rows have errors (safety rule).'
            }
            Write-RunLog ERROR ('SAFETY: {0} row(s) have errors. NOTHING was changed on SSRS. Fix the errors or set those rows to Enabled=false.' -f $errors.Count)
        } elseif ($ready.Count -eq 0) {
            Write-RunLog INFO 'Nothing to change.'
        } else {
            $n = 0
            foreach ($plan in $ready) {
                $n++
                $isOverwrite = $plan.Action -eq 'READY_OVERWRITE'
                $script:RowTag = 'Row {0}' -f $plan.ExcelRow
                Write-RunLog STEP ('[{0}/{1}] {2}: {3} | {4}' -f $n, $ready.Count, $(if ($isOverwrite) { 'UPDATING' } else { 'Creating' }), $plan.ReportPath, $plan.SubscriptionName)
                try {
                    # Last-second check: has anything changed since the pre-check?
                    $same = @(Get-SameNameSubscription $plan -Quiet)
                    $sw = [Diagnostics.Stopwatch]::StartNew()
                    if ($isOverwrite) {
                        if ($same.Count -ne 1 -or -not [string]::Equals([string]$same[0].SubscriptionID, $plan.ExistingId, $script:IgnoreCase)) {
                            throw ('The subscription changed after the pre-check (expected one ID={0}, found: {1}). Not updated for safety; run again.' -f $plan.ExistingId, (($same | ForEach-Object { $_.SubscriptionID }) -join ';'))
                        }
                        $script:Proxy.SetSubscriptionProperties($plan.ExistingId, $plan.ExtSettings, $plan.SubscriptionName, 'TimedSubscription', $plan.ScheduleXml, $plan.ParamValues)
                        $plan.SubscriptionID = $plan.ExistingId
                        $plan.Action         = 'OVERWRITTEN'
                        Write-RunLog OK ('Subscription UPDATED. ID={0} ({1} ms)' -f $plan.SubscriptionID, $sw.ElapsedMilliseconds)
                    } else {
                        if ($same.Count -gt 0) {
                            $plan.Action         = 'SKIP_EXISTING'
                            $plan.SubscriptionID = ($same | ForEach-Object { $_.SubscriptionID }) -join ';'
                            $plan.Message        = 'A subscription with this name appeared after the pre-check; left alone.'
                            Write-RunLog WARN $plan.Message
                            continue
                        }
                        $newId = $script:Proxy.CreateSubscription($plan.ReportPath, $plan.ExtSettings, $plan.SubscriptionName, 'TimedSubscription', $plan.ScheduleXml, $plan.ParamValues)
                        if (-not $newId) { throw 'CreateSubscription returned an empty ID.' }
                        $plan.SubscriptionID = [string]$newId
                        $plan.Action         = 'CREATED'
                        Write-RunLog OK ('Subscription CREATED. ID={0} ({1} ms)' -f $plan.SubscriptionID, $sw.ElapsedMilliseconds)
                    }
                } catch {
                    $plan.Action  = 'ERROR'
                    $plan.Message = Get-ShortError $_.Exception
                    Write-ErrorDetail $_ $(if ($isOverwrite) { 'Update error' } else { 'Create error' })
                    continue
                }
                Confirm-Subscription $plan -IsOverwrite:$isOverwrite
            }
            $script:RowTag = ''
        }
    }
}
catch {
    $script:Fatal = $true
    $script:RowTag = ''
    Write-ErrorDetail $_ 'FATAL ERROR'
}
finally {
    try { Write-Summary }
    catch { Write-Host ('Could not write the summary: {0}' -f $_.Exception.Message) -ForegroundColor Red; $script:ExitCode = 1 }
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}
exit $script:ExitCode
