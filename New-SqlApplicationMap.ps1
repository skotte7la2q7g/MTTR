#Requires -Version 5.1
<#
=========================================================================================
 SQL APPLICATION MAP
-----------------------------------------------------------------------------------------
 WHAT THIS PRODUCES
   One self-contained HTML page. Pick an application and it tells you which SQL servers
   it runs on in every environment (PROD, DR, UAT, QA, DEV ...), which SQL Server
   version each one is on, and which databases sit underneath. Read top to bottom:

     1. UNKNOWN APPLICATION  A large red block. Every database and server the inventory
                             cannot tie to an application, the raw values that caused it,
                             and a lead where the other sheet names an owner. One click
                             exports the list as CSV for the clean-up.
     2. Applications         One row per application: servers and databases in each
                             environment, SQL versions, data size and flags. Click a row.
     3. Application detail   The application at the top of a relationship map, one card
                             per environment, the SQL servers inside it with version and
                             server metadata, and the databases on each server.
     4. SQL Server estate    Version x environment layout with Microsoft support status,
                             platform, location, capacity and database configuration.
                             Click a cell to list the applications on it.
     5. Reconciliation       Rows read, repeats removed, records that did not match
                             across the two sheets - so every total traces back.

 INPUTS
   One workbook with two sheets, or two workbooks / CSV files:
     Server sheet    ServerName, Environment, SQLVersion, BusinessUnit, APMAppID,
                     ApplicationCriticality, ApplicationName, CostCenter, Location,
                     Platform, PhysicalMemoryGB, LogicalCPUCount, Port, PrimaryITOwner,
                     SecondaryITOwner, Size, Subscription, VirtualNetwork, Subnet
     Database sheet  DatabaseType, ServerName, DatabaseName, DatabaseSizeMB,
                     DatabaseStatus, Environment, SQLVersion, BusinessUnit,
                     ApplicationName, PrimaryITOwner, SecondaryITOwner, AGName,
                     CompatibilityLevel, DatabaseClassification, DatabaseCollationName,
                     ListenerName, PageVerifyOption, Platform, RecoveryModel,
                     InventoryDate - Year / Quarter / Month / Day
   Sheets are recognised by their columns, not their names or positions. Pin them with
   -ServerSheet / -DatabaseSheet if the workbook carries other look-alike sheets.

 RULES THE PAGE FOLLOWS
   1. An application name that is blank, Unknown, N/A, TBD, #N/A and the like is UNKNOWN
      and is shown in red. It is never guessed. Where the other sheet ties the same
      server or APM ID to a named application, that is shown beside it as a lead.
   2. A database listed more than once (several inventory dates) is counted once, from
      its latest inventory date. System databases are counted separately.
   3. A value that is not in the workbook prints as a dash, never as zero.
   4. Support status uses Microsoft's fixed-lifecycle dates (table below), measured
      against the day the page is generated. Paid Extended Security Updates are ignored.

 USAGE
   .\New-SqlApplicationMap.ps1 -Simulate                        # made-up estate, nothing read
   .\New-SqlApplicationMap.ps1                                  # pick the workbook(s) in a dialog
   .\New-SqlApplicationMap.ps1 -SourceWorkbook .\SQL_Inventory.xlsx
   .\New-SqlApplicationMap.ps1 -SourceWorkbook .\Servers.xlsx, .\Databases.xlsx
   .\New-SqlApplicationMap.ps1 -SourceWorkbook .\Inv.xlsx -ServerSheet 'Sheet1' -DatabaseSheet 'Sheet2'
=========================================================================================
#>

[CmdletBinding()]
param(
    # One workbook holding both sheets, or one file per sheet. .xlsx, .xlsm, .xls or .csv.
    [Parameter(Mandatory = $false)]
    [string[]]$SourceWorkbook = @(),

    [Parameter(Mandatory = $false)]
    [string]$OutputHtml = '',

    # Only needed when the workbook carries other sheets that look like inventory sheets.
    [Parameter(Mandatory = $false)]
    [string]$ServerSheet = '',

    [Parameter(Mandatory = $false)]
    [string]$DatabaseSheet = '',

    # A release whose extended support ends within this many months is called out.
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 60)]
    [int]$SupportWarningMonths = 18,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Manulife', 'Slate')]
    [string]$Palette = 'Manulife',

    # Build the page from a built-in, made-up estate. Nothing is read.
    [Parameter(Mandatory = $false)]
    [switch]$Simulate,

    [Parameter(Mandatory = $false)]
    [switch]$NoLaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
trap { Write-Host ("FAILED at line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) -ForegroundColor Red; break }

$ReportDate = Get-Date
$Invariant = [Globalization.CultureInfo]::InvariantCulture

# =========================================================================================
# REGION  Reference data
# =========================================================================================

# Microsoft fixed-lifecycle dates, from learn.microsoft.com/lifecycle/products/sql-server-<year>.
# Native is the compatibility level a new database gets on that release.
$SqlReleases = @(
    @{ Label = 'Azure SQL (PaaS)';   Order = 9999;   Native = $null; Mainstream = '';           Extended = '' }
    @{ Label = 'SQL Server 2025';    Order = 2025;   Native = 170;   Mainstream = '2031-01-06'; Extended = '2036-01-06' }
    @{ Label = 'SQL Server 2022';    Order = 2022;   Native = 160;   Mainstream = '2028-01-11'; Extended = '2033-01-11' }
    @{ Label = 'SQL Server 2019';    Order = 2019;   Native = 150;   Mainstream = '2025-02-28'; Extended = '2030-01-08' }
    @{ Label = 'SQL Server 2017';    Order = 2017;   Native = 140;   Mainstream = '2022-10-11'; Extended = '2027-10-12' }
    @{ Label = 'SQL Server 2016';    Order = 2016;   Native = 130;   Mainstream = '2021-07-13'; Extended = '2026-07-14' }
    @{ Label = 'SQL Server 2014';    Order = 2014;   Native = 120;   Mainstream = '2019-07-09'; Extended = '2024-07-09' }
    @{ Label = 'SQL Server 2012';    Order = 2012;   Native = 110;   Mainstream = '2017-07-11'; Extended = '2022-07-12' }
    @{ Label = 'SQL Server 2008 R2'; Order = 2008.5; Native = 100;   Mainstream = '2014-07-08'; Extended = '2019-07-09' }
    @{ Label = 'SQL Server 2008';    Order = 2008;   Native = 100;   Mainstream = '2014-07-08'; Extended = '2019-07-09' }
    @{ Label = 'SQL Server 2005';    Order = 2005;   Native = 90;    Mainstream = '2011-04-12'; Extended = '2016-04-12' }
    @{ Label = 'SQL Server 2000';    Order = 2000;   Native = 80;    Mainstream = '2008-04-08'; Extended = '2013-04-09' }
    @{ Label = 'Unrecognised';       Order = 1;      Native = $null; Mainstream = '';           Extended = '' }
    @{ Label = 'Not recorded';       Order = 0;      Native = $null; Mainstream = '';           Extended = '' }
)

$MajorRelease = @{
    '8' = 'SQL Server 2000'; '9' = 'SQL Server 2005'; '11' = 'SQL Server 2012'; '12' = 'SQL Server 2014'
    '13' = 'SQL Server 2016'; '14' = 'SQL Server 2017'; '15' = 'SQL Server 2019'; '16' = 'SQL Server 2022'; '17' = 'SQL Server 2025'
}
$PaasPattern = 'managed\s*instance|azure\s*sql\s*(db|database|mi|managed)|\bsql\s*mi\b|\bsqlmi\b|\bpaas\b|elastic\s*pool|hyperscale'

# Any value that says "we do not know". Applies to application names, APM IDs, AG names...
$UnknownValuePattern = '^(unknowns?\b.*|unk|n/?a|#.*|none|null|nil|nan|tb[acd]|to be (confirmed|decided|determined)|not (available|applicable|assigned|known|defined|mapped|found|identified)|unassigned|unmapped|unidentified|undefined|no app(lication)?( name)?|\(?blank\)?|empty|[-?._*0]+)$'

$SystemDatabaseNames = @('master', 'model', 'msdb', 'tempdb', 'mssqlsystemresource')

# Canonical field -> accepted header spellings (lower case, letters and digits only).
$ServerFieldMap = [ordered]@{
    ServerName             = @('servername', 'server', 'sqlservername', 'sqlserver', 'instancename', 'sqlinstance', 'instance', 'hostname')
    Environment            = @('environment', 'env')
    SQLVersion             = @('sqlversion', 'sqlserverversion', 'version', 'productversion')
    BusinessUnit           = @('businessunit', 'bu')
    APMAppID               = @('apmappid', 'apmid', 'appid', 'applicationid')
    ApplicationCriticality = @('applicationcriticality', 'criticality', 'appcriticality', 'tier')
    ApplicationName        = @('applicationname', 'application', 'appname')
    CostCenter             = @('costcenter', 'costcentre')
    Location               = @('location', 'datacenter', 'datacentre', 'region', 'site')
    Platform               = @('platform', 'hostingplatform')
    PhysicalMemoryGB       = @('physicalmemorygb', 'memorygb', 'ramgb', 'physicalmemory', 'memory')
    LogicalCPUCount        = @('logicalcpucount', 'cpucount', 'logicalcpus', 'vcpu', 'vcpus', 'cores')
    Port                   = @('port', 'tcpport')
    PrimaryITOwner         = @('primaryitowner', 'itowner', 'primaryowner')
    SecondaryITOwner       = @('secondaryitowner', 'secondaryowner')
    Size                   = @('size', 'vmsize', 'sku')
    Subscription           = @('subscription', 'subscriptionname')
    VirtualNetwork         = @('virtualnetwork', 'vnet')
    Subnet                 = @('subnet')
}
$DatabaseFieldMap = [ordered]@{
    DatabaseType           = @('databasetype', 'dbtype', 'type')
    ServerName             = @('servername', 'server', 'sqlservername', 'sqlserver', 'instancename', 'sqlinstance', 'instance', 'hostname')
    DatabaseName           = @('databasename', 'database', 'dbname')
    DatabaseSizeMB         = @('databasesizemb', 'sizemb', 'dbsizemb', 'databasesize')
    DatabaseStatus         = @('databasestatus', 'dbstatus', 'status', 'state')
    Environment            = @('environment', 'env')
    SQLVersion             = @('sqlversion', 'sqlserverversion', 'version', 'productversion')
    BusinessUnit           = @('businessunit', 'bu')
    ApplicationName        = @('applicationname', 'application', 'appname')
    PrimaryITOwner         = @('primaryitowner', 'itowner', 'primaryowner')
    SecondaryITOwner       = @('secondaryitowner', 'secondaryowner')
    AGName                 = @('agname', 'availabilitygroup', 'availabilitygroupname', 'ag')
    CompatibilityLevel     = @('compatibilitylevel', 'compatlevel', 'compatibility')
    DatabaseClassification = @('databaseclassification', 'classification', 'dataclassification')
    DatabaseCollationName  = @('databasecollationname', 'collationname', 'collation')
    ListenerName           = @('listenername', 'aglistener', 'listener')
    PageVerifyOption       = @('pageverifyoption', 'pageverify')
    Platform               = @('platform', 'hostingplatform')
    RecoveryModel          = @('recoverymodel', 'recoverymodeldesc')
    InventoryYear          = @('inventorydateyear', 'inventoryyear')
    InventoryQuarter       = @('inventorydatequarter', 'inventoryquarter')
    InventoryMonth         = @('inventorydatemonth', 'inventorymonth')
    InventoryDay           = @('inventorydateday', 'inventoryday')
    InventoryDate          = @('inventorydate', 'collectiondate', 'snapshotdate')
}

$KnownHeaders = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($fieldMap in @($ServerFieldMap, $DatabaseFieldMap)) {
    foreach ($aliases in $fieldMap.Values) { foreach ($alias in $aliases) { [void]$KnownHeaders.Add($alias) } }
}

# =========================================================================================
# REGION  Small helpers
# =========================================================================================

function ConvertTo-HeaderKey {
    param($Value)
    if ($null -eq $Value) { return '' }
    return (([string]$Value) -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function ConvertTo-Number {
    param([string]$Text)
    if (-not $Text) { return $null }
    $clean = $Text -replace '[^0-9.\-]', ''
    $number = 0.0
    if ($clean -and [double]::TryParse($clean, [Globalization.NumberStyles]::Float, $Invariant, [ref]$number)) { return $number }
    return $null
}

function Clear-UnknownValue {
    param([string]$Value)
    if (-not $Value -or $Value -match $UnknownValuePattern) { return '' }
    return $Value
}

function New-StringSet {
    return , (New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase))
}

function Add-KnownValue {
    param($Set, [string]$Value)
    if ($Value -and $Value -notmatch $UnknownValuePattern) { [void]$Set.Add($Value) }
}

function Join-Set {
    param($Set)
    return ((@($Set) | Sort-Object) -join ', ')
}

function Get-MostCommon {
    param($Counts)
    $best = ''
    $bestCount = 0
    foreach ($entry in $Counts.GetEnumerator()) {
        if ($entry.Key -and $entry.Key -ne 'UNKNOWN' -and $entry.Value -gt $bestCount) { $best = [string]$entry.Key; $bestCount = $entry.Value }
    }
    return $best
}

function Add-Count {
    param($Counts, [string]$Key)
    if (-not $Key) { return }
    if ($Counts.ContainsKey($Key)) { $Counts[$Key]++ } else { $Counts[$Key] = 1 }
}

# "HOST\INST,1433", "tcp:host.corp.com" and "HOST\MSSQLSERVER" all reduce to one key.
function Get-ServerKey {
    param([string]$Name)
    $key = $Name.Trim().ToUpperInvariant()
    $key = $key -replace '^TCP:', ''
    $key = $key -replace '\s*,\s*\d+$', ''
    $key = $key -replace '\\MSSQLSERVER$', ''
    return $key
}

function Get-HostKey {
    param([string]$Key)
    $hostName = ($Key -split '\\')[0]
    if ($hostName -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { $hostName = ($hostName -split '\.')[0] }
    return $hostName
}

function Get-InstancePart {
    param([string]$Key)
    if ($Key.Contains('\')) { return ($Key -split '\\', 2)[1] }
    return ''
}

$EnvironmentCache = @{}
function Get-CanonicalEnvironment {
    param([string]$Value)
    if ($EnvironmentCache.ContainsKey($Value)) { return $EnvironmentCache[$Value] }
    $text = $Value.ToUpperInvariant()
    $result = $text
    if (-not $text) { $result = 'UNKNOWN' }
    elseif ($text -match '\bNON[\s-]*PRO?D') { $result = 'NONPROD' }
    elseif ($text -match '\bPRE[\s-]*PRO?D|\bPPE\b|\bSTAG(E|ING)\b|\bSTG\b') { $result = 'PREPROD' }
    elseif ($text -match '^DR\d*\b|\bDR\b|DISASTER|\bBCP\b|\bCOB\b') { $result = 'DR' }
    elseif ($text -match '^UAT|\bUAT\b|USER\s*ACCEPT') { $result = 'UAT' }
    elseif ($text -match '^SIT|\bSIT\b|INTEGRATION') { $result = 'SIT' }
    elseif ($text -match '^QA|\bQA\b|QUALITY') { $result = 'QA' }
    elseif ($text -match '\bTE?ST(ING)?\b|^TE?ST') { $result = 'TEST' }
    elseif ($text -match '^DEV|\bDEV(ELOPMENT)?\b') { $result = 'DEV' }
    elseif ($text -match 'SANDBOX|\bSBX\b|\bPOC\b|\bLAB\b') { $result = 'SANDBOX' }
    elseif ($text -match '^PRO?D|\bPRO?D(UCTION)?\b|\bLIVE\b|^P$') { $result = 'PROD' }
    elseif ($text -match $UnknownValuePattern) { $result = 'UNKNOWN' }
    $EnvironmentCache[$Value] = $result
    return $result
}

# A product version like 13.0.6300.2 means nothing to most readers. Name the release.
function Get-SqlRelease {
    param([string]$Version, [string]$Context = '')
    if ($Context -match $PaasPattern -or $Version -match $PaasPattern -or $Version -match '^12\.0\.2000\.') { return 'Azure SQL (PaaS)' }
    if (-not $Version) { return 'Not recorded' }
    if ($Version -match '(?<![\d.])(8|9|10|11|12|13|14|15|16|17)\.(\d{1,2})\.\d+') {
        if ($Matches[1] -eq '10') { if ([int]$Matches[2] -ge 50) { return 'SQL Server 2008 R2' } else { return 'SQL Server 2008' } }
        return $MajorRelease[$Matches[1]]
    }
    if ($Version -match '(?<!\d)(2000|2005|2008|2012|2014|2016|2017|2019|2022|2025)(?!\d)') {
        if ($Matches[1] -eq '2008' -and $Version -match '2008\s*R2') { return 'SQL Server 2008 R2' }
        return 'SQL Server ' + $Matches[1]
    }
    if ($Version -match '^(8|9|10|11|12|13|14|15|16|17)(\.(\d{1,2}))?$') {
        if ($Matches[1] -eq '10') { if ($Version -match '^10\.5') { return 'SQL Server 2008 R2' } else { return 'SQL Server 2008' } }
        return $MajorRelease[$Matches[1]]
    }
    return 'Unrecognised'
}

function Get-ReleaseStatus {
    param($Release)
    switch ($Release.Label) {
        'Azure SQL (PaaS)' { return @('paas', 'Evergreen - patched and upgraded by Microsoft') }
        'Unrecognised'     { return @('unknown', 'Version text not recognised') }
        'Not recorded'     { return @('unknown', 'No version in the inventory') }
    }
    $mainstream = [datetime]::ParseExact($Release.Mainstream, 'yyyy-MM-dd', $Invariant)
    $extended = [datetime]::ParseExact($Release.Extended, 'yyyy-MM-dd', $Invariant)
    if ($ReportDate -gt $extended) { return @('oos', ('Out of support since {0}' -f $extended.ToString('d MMM yyyy', $Invariant))) }
    if ($extended -le $ReportDate.AddMonths($SupportWarningMonths)) { return @('ending', ('Support ends {0}' -f $extended.ToString('d MMM yyyy', $Invariant))) }
    if ($ReportDate -gt $mainstream) { return @('extended', ('Extended support to {0}' -f $extended.ToString('d MMM yyyy', $Invariant))) }
    return @('mainstream', ('Mainstream support to {0}' -f $mainstream.ToString('d MMM yyyy', $Invariant)))
}

$MonthNumber = @{ jan = 1; feb = 2; mar = 3; apr = 4; may = 5; jun = 6; jul = 7; aug = 8; sep = 9; oct = 10; nov = 11; dec = 12 }

# InventoryDate may arrive as one date, or as the Year / Quarter / Month / Day hierarchy a
# Power BI or pivot export produces. Month can be a number or a name.
function Get-InventoryDate {
    param($Record)
    $single = $Record['InventoryDate']
    if ($single) {
        $parsed = [datetime]::MinValue
        $number = 0.0
        if ([double]::TryParse($single, [Globalization.NumberStyles]::Float, $Invariant, [ref]$number) -and $number -gt 20000 -and $number -lt 80000) {
            $parsed = [datetime]::FromOADate($number)
        }
        elseif (-not [datetime]::TryParse($single, $Invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            $parsed = [datetime]::MinValue
        }
        if ($parsed -ne [datetime]::MinValue) { return @{ Sort = [int]$parsed.ToString('yyyyMMdd'); Text = $parsed.ToString('yyyy-MM-dd') } }
    }

    $year = 0
    if (-not [int]::TryParse(($Record['InventoryYear'] -replace '[^0-9]', ''), [ref]$year) -or $year -lt 1990 -or $year -gt 2100) {
        return @{ Sort = 0; Text = '' }
    }
    $month = 0
    $monthText = $Record['InventoryMonth']
    if ($monthText) {
        if (-not [int]::TryParse(($monthText -replace '[^0-9]', ''), [ref]$month)) { $month = 0 }
        if ($month -lt 1 -or $month -gt 12) {
            $month = 0
            $letters = $monthText -replace '[^A-Za-z]', ''
            if ($letters.Length -ge 3) {
                $short = $letters.Substring(0, 3).ToLowerInvariant()
                if ($MonthNumber.ContainsKey($short)) { $month = $MonthNumber[$short] }
            }
        }
    }
    if ($month -eq 0) {
        $quarter = 0
        if ([int]::TryParse(($Record['InventoryQuarter'] -replace '[^0-9]', ''), [ref]$quarter) -and $quarter -ge 1 -and $quarter -le 4) { $month = $quarter * 3 }
    }
    $day = 0
    if (-not [int]::TryParse(($Record['InventoryDay'] -replace '[^0-9]', ''), [ref]$day) -or $day -lt 1 -or $day -gt 31) { $day = 0 }

    $text = '{0:0000}' -f $year
    if ($month) {
        $text += '-{0:00}' -f $month
        if ($day) { $text += '-{0:00}' -f $day }
    }
    return @{ Sort = $year * 10000 + $month * 100 + $day; Text = $text }
}

# =========================================================================================
# REGION  Reading workbooks and CSV files
#
#   Every source becomes the same shape: a 1-based two-dimensional array of cell values,
#   exactly what Excel's UsedRange.Value2 returns. Everything after that is shared.
# =========================================================================================

function Select-SourceFiles {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select the SQL inventory workbook (or both files: servers and databases)'
    $dialog.Filter = 'Excel or CSV (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|All files (*.*)|*.*'
    $dialog.Multiselect = $true
    $dialog.CheckFileExists = $true
    $dialog.RestoreDirectory = $true
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'No workbook was selected. The page was not generated.'
    }
    return $dialog.FileNames
}

function Read-ExcelTables {
    param([Parameter(Mandatory)][string]$Path)
    $tables = [System.Collections.Generic.List[object]]::new()
    $excel = $null
    $workbook = $null
    try {
        try { $excel = New-Object -ComObject Excel.Application }
        catch { throw 'Microsoft Excel is needed to read .xlsx files and is not available. Save each sheet as CSV and pass the CSV files instead.' }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path, 0, $true)
        foreach ($sheet in @($workbook.Worksheets)) {
            $used = $sheet.UsedRange
            $rowCount = [int]$used.Rows.Count
            $columnCount = [int]$used.Columns.Count
            if ($rowCount -ge 2) {
                Write-Host ("  Reading sheet '{0}' ({1:N0} rows x {2} columns)..." -f $sheet.Name, $rowCount, $columnCount) -ForegroundColor DarkCyan
                $tables.Add([pscustomobject]@{
                    File = [IO.Path]::GetFileName($Path); Sheet = [string]$sheet.Name
                    Values = $used.Value2; RowCount = $rowCount; ColumnCount = $columnCount
                })
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        }
    }
    finally {
        if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
        if ($null -ne $excel) { $excel.Quit() }
        if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return , $tables.ToArray()
}

function Read-CsvTable {
    param([Parameter(Mandatory)][string]$Path)
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return , @() }
    $headers = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $values = New-Object 'object[,]' ($rows.Count + 2), ($headers.Count + 1)
    for ($c = 0; $c -lt $headers.Count; $c++) { $values[1, ($c + 1)] = $headers[$c] }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $headers.Count; $c++) { $values[($r + 2), ($c + 1)] = $row.($headers[$c]) }
    }
    $table = [pscustomobject]@{
        File = [IO.Path]::GetFileName($Path); Sheet = [IO.Path]::GetFileNameWithoutExtension($Path)
        Values = $values; RowCount = $rows.Count + 1; ColumnCount = $headers.Count
    }
    return , @($table)
}

# Exports sometimes put a title above the header. The header is the row, among the first
# ten, that carries the most recognised column names.
function Get-HeaderRow {
    param($Table)
    $best = 1
    $bestScore = -1
    for ($r = 1; $r -le [math]::Min(10, $Table.RowCount); $r++) {
        $score = 0
        for ($c = 1; $c -le $Table.ColumnCount; $c++) {
            $key = ConvertTo-HeaderKey $Table.Values[$r, $c]
            if ($key -and $KnownHeaders.Contains($key)) { $score++ }
        }
        if ($score -gt $bestScore) { $best = $r; $bestScore = $score }
    }
    return $best
}

function Get-ColumnMap {
    param($Table, [int]$HeaderRow, $FieldMap)
    $byHeader = @{}
    for ($c = 1; $c -le $Table.ColumnCount; $c++) {
        $key = ConvertTo-HeaderKey $Table.Values[$HeaderRow, $c]
        if ($key -and -not $byHeader.ContainsKey($key)) { $byHeader[$key] = $c }
    }
    $map = [ordered]@{}
    foreach ($field in $FieldMap.Keys) {
        foreach ($alias in $FieldMap[$field]) {
            if ($byHeader.ContainsKey($alias)) { $map[$field] = $byHeader[$alias]; break }
        }
    }
    return $map
}

# One hashtable per row, keyed by canonical field. A field the sheet does not carry is ''.
# Excel error cells (#N/A, #REF! ...) arrive from Value2 as large negative integers.
function Read-Records {
    param($Table, [int]$HeaderRow, $ColumnMap, $FieldMap, [string]$Source)
    $records = [System.Collections.Generic.List[hashtable]]::new()
    $fields = @($FieldMap.Keys)
    $columns = @(foreach ($field in $fields) { if ($ColumnMap.Contains($field)) { [int]$ColumnMap[$field] } else { 0 } })
    $values = $Table.Values
    $whitespace = [regex]'\s+'
    $total = $Table.RowCount
    for ($r = $HeaderRow + 1; $r -le $total; $r++) {
        if (($r % 5000) -eq 0) { Write-Progress -Activity "Reading $Source" -Status ("{0:N0} of {1:N0} rows" -f $r, $total) -PercentComplete (100 * $r / $total) }
        $record = @{}
        $any = $false
        for ($i = 0; $i -lt $fields.Count; $i++) {
            $text = ''
            $column = $columns[$i]
            if ($column -gt 0) {
                $raw = $values[$r, $column]
                if ($null -ne $raw) {
                    if ($raw -is [int] -and $raw -le -2146826246 -and $raw -ge -2146826288) { $text = '#N/A' }
                    else { $text = $whitespace.Replace([string]$raw, ' ').Trim() }
                }
            }
            $record[$fields[$i]] = $text
            if ($text) { $any = $true }
        }
        if ($any) { $record['_Source'] = $Source; $records.Add($record) }
    }
    Write-Progress -Activity "Reading $Source" -Completed
    return , $records
}

# =========================================================================================
# REGION  Simulator
#
#   A made-up estate in exactly the shape the sheet reader produces, so everything after
#   this point runs as it does for real. Built to show every case: unknown applications
#   recorded as blank / Unknown / N/A / TBD / #N/A, a server the server sheet ties to an
#   APM ID but no name, a server missing from the server sheet, a server with no
#   databases, a version that disagrees between sheets, two inventory snapshots, system
#   databases, out-of-support releases, an Azure managed instance, PROD without DR.
# =========================================================================================

function Get-SimulatedInventory {
    $random = New-Object System.Random 20261002
    $serverRows = [System.Collections.Generic.List[hashtable]]::new()
    $databaseRows = [System.Collections.Generic.List[hashtable]]::new()
    $nativeByMajor = @{ '10' = 100; '11' = 110; '12' = 120; '13' = 130; '14' = 140; '15' = 150; '16' = 160; '2019' = 150; '2014' = 120 }

    function Add-SimServer {
        param([string]$Name, [string]$Env, [string]$App, [string]$Apm, [string]$Crit, [string]$Version,
              [string]$Platform = 'Azure IaaS', [string]$Location = 'East US 2', [int]$Cpu = 8, [int]$Memory = 64, [string]$Owner = 'A. Patel')
        $tier = $(if ($Env -match 'Prod|DR|Disaster') { 'prod' } else { 'nonprod' })
        $azure = $Platform -match 'Azure'
        $serverRows.Add(@{
            ServerName = $Name; Environment = $Env; SQLVersion = $Version; BusinessUnit = 'US Retirement'; APMAppID = $Apm
            ApplicationCriticality = $Crit; ApplicationName = $App; CostCenter = 'CC-4410'; Location = $Location; Platform = $Platform
            PhysicalMemoryGB = [string]$Memory; LogicalCPUCount = [string]$Cpu; Port = '1433'; PrimaryITOwner = $Owner; SecondaryITOwner = 'SQL DBA Team'
            Size = $(if ($azure) { 'Standard_E{0}ds_v5' -f $Cpu } else { 'VMware VM' })
            Subscription = $(if ($azure) { "sub-usr-$tier-01" } else { '' })
            VirtualNetwork = $(if ($azure) { "vnet-usr-$tier-eus2" } else { '' })
            Subnet = $(if ($azure) { "snet-sql-$tier" } else { '' })
            _Source = 'SIMULATED > Servers'
        })
    }

    function Add-SimDatabase {
        param([string]$Server, [string]$Database, [string]$Env, [string]$App, [string]$Version, [double]$SizeMb,
              [string]$Status = 'ONLINE', [string]$Recovery = 'FULL', [string]$PageVerify = 'CHECKSUM', $Compat = $null,
              [string]$Ag = '', [string]$Listener = '', [string]$Classification = 'Internal',
              [string]$Collation = 'SQL_Latin1_General_CP1_CI_AS', [string]$Type = 'User', [string]$Platform = 'Azure IaaS',
              [string]$Owner = '', [switch]$OlderSnapshot, [switch]$OnlyOlder)
        $row = @{
            DatabaseType = $Type; ServerName = $Server; DatabaseName = $Database; DatabaseSizeMB = [string][math]::Round($SizeMb, 2)
            DatabaseStatus = $Status; Environment = $Env; SQLVersion = $Version; BusinessUnit = 'US Retirement'; ApplicationName = $App
            PrimaryITOwner = $Owner; SecondaryITOwner = $(if ($Owner) { 'SQL DBA Team' } else { '' }); AGName = $Ag
            CompatibilityLevel = $(if ($null -ne $Compat) { [string]$Compat } else { '' }); DatabaseClassification = $Classification
            DatabaseCollationName = $Collation; ListenerName = $Listener; PageVerifyOption = $PageVerify; Platform = $Platform
            RecoveryModel = $Recovery; InventoryYear = '2026'; InventoryQuarter = 'Qtr 3'; InventoryMonth = 'September'; InventoryDay = '28'
            InventoryDate = ''; _Source = 'SIMULATED > Databases'
        }
        if (-not $OnlyOlder) { $databaseRows.Add($row) }
        if ($OlderSnapshot -or $OnlyOlder) {
            $older = $row.Clone()
            $older.InventoryMonth = 'August'; $older.InventoryDay = '31'
            $older.DatabaseSizeMB = [string][math]::Round($SizeMb * 0.97, 2)
            $databaseRows.Add($older)
        }
    }

    function Add-SimSystemDatabases {
        param([string]$Server, [string]$Env, [string]$Version, [string]$Platform = 'Azure IaaS')
        foreach ($name in @('master', 'model', 'msdb', 'tempdb')) {
            Add-SimDatabase -Server $Server -Database $name -Env $Env -App '' -Version $Version -SizeMb (40 + $random.Next(0, 900)) -Type 'System' -Recovery 'SIMPLE' -Platform $Platform
        }
    }

    function Get-SimNative {
        param([string]$Version)
        $major = ($Version -split '\.')[0]
        if ($Version -match '^10\.50') { return 100 }
        if ($nativeByMajor.ContainsKey($major)) { return $nativeByMajor[$major] }
        return $null
    }

    $environmentText = @{ PRD = @('PROD', 'Production'); DR = @('DR', 'Disaster Recovery'); UAT = @('UAT', 'UAT'); QA = @('QA', 'Quality Assurance'); DEV = @('DEV', 'Development') }
    $catalog = @(
        @{ App = 'Retirement Plan Admin'; Code = 'RPA'; Apm = '10234'; Crit = 'Critical'; Owner = 'A. Patel';  Tiers = 'PRD,DR,UAT,QA,DEV'; Ver = '13.0.6300.2,13.0.6300.2,15.0.4355.3,15.0.4355.3,16.0.4135.4'; Dbs = 'RPA_Core,RPA_Audit,RPA_Stage,RPA_Reporting,RPA_Archive'; Ag = $true; Scale = 9 }
        @{ App = 'Participant Web';       Code = 'PWB'; Apm = '10456'; Crit = 'High';     Owner = 'M. Chen';   Tiers = 'PRD,DR,UAT,QA,DEV'; Ver = '15.0.4355.3,15.0.4355.3,15.0.4355.3,2019,15.0.4355.3'; Dbs = 'PWB_Session,PWB_Content,PWB_Profile'; Ag = $true; Scale = 3 }
        @{ App = 'Recordkeeping Core';    Code = 'RKC'; Apm = '10021'; Crit = 'Critical'; Owner = 'S. Okafor'; Tiers = 'PRD,DR,UAT,QA,DEV'; Ver = '14.0.3465.1,14.0.3465.1,14.0.3465.1,15.0.4355.3,15.0.4355.3'; Dbs = 'RK_Ledger,RK_Participant,RK_Plan,RK_Trade,RK_Batch,RK_History'; Ag = $true; Scale = 20 }
        @{ App = 'Plan Sponsor Portal';   Code = 'PSP'; Apm = '11003'; Crit = 'High';     Owner = 'L. Romero'; Tiers = 'PRD,UAT,DEV';       Ver = '14.0.3465.1,14.0.3465.1,16.0.4135.4'; Dbs = 'PSP_Main,PSP_Docs'; Ag = $false; Scale = 2 }
        @{ App = 'Data Warehouse';        Code = 'DWH'; Apm = '10500'; Crit = 'High';     Owner = 'K. Ivanova'; Tiers = 'PRD,QA,DEV';       Ver = '12.0.6024.0,12.0.6024.0,13.0.6300.2'; Dbs = 'DWH_Stage,DWH_Mart,DWH_History'; Ag = $false; Scale = 60 }
        @{ App = 'Fund Pricing';          Code = 'FPR'; Apm = '10611'; Crit = 'Critical'; Owner = 'D. Byrne';  Tiers = 'PRD,DR,UAT';        Ver = '12.0.2000.8,12.0.2000.8,16.0.4135.4'; Dbs = 'FPR_Prices,FPR_Feeds'; Ag = $false; Scale = 6 }
        @{ App = 'Compliance Reporting';  Code = 'CMP'; Apm = '10812'; Crit = 'Medium';   Owner = 'R. Haddad'; Tiers = 'PRD,QA';            Ver = '11.0.7001.0,11.0.7001.0'; Dbs = 'CMP_Rules,CMP_Evidence'; Ag = $false; Scale = 4 }
        @{ App = 'Call Centre CRM';       Code = 'CRM'; Apm = '10345'; Crit = 'High';     Owner = 'T. Nguyen'; Tiers = 'PRD,DR,UAT,DEV';    Ver = '16.0.4135.4,16.0.4135.4,16.0.4135.4,16.0.4135.4'; Dbs = 'CRM_Main,CRM_Cases,CRM_Integration'; Ag = $true; Scale = 5 }
        @{ App = 'Document Imaging';      Code = 'DIM'; Apm = '10777'; Crit = 'Medium';   Owner = 'P. Silva';  Tiers = 'PRD,QA';            Ver = '13.0.6435.1,13.0.6435.1'; Dbs = 'DIM_Index,DIM_Store'; Ag = $false; Scale = 30 }
    )

    $appIndex = 0
    foreach ($entry in $catalog) {
        $appIndex++
        $tiers = $entry.Tiers -split ','
        $versions = $entry.Ver -split ','
        $dbNames = @($entry.Dbs -split ',')
        for ($t = 0; $t -lt $tiers.Count; $t++) {
            $tier = $tiers[$t]
            $version = $versions[$t]
            $env = $environmentText[$tier][$appIndex % 2]
            $prodLike = $tier -in @('PRD', 'DR')
            $nodes = $(if ($tier -eq 'PRD' -and $entry.Crit -eq 'Critical' -and $entry.Code -ne 'FPR') { 2 } else { 1 })
            $platform = $(if ($version -like '12.0.2000.*') { 'Azure SQL Managed Instance' } elseif ($entry.Code -in @('CMP', 'DIM')) { 'On-Prem VMware' } else { 'Azure IaaS' })
            $location = $(if ($platform -eq 'On-Prem VMware') { 'Boston DC' } elseif ($tier -eq 'DR') { 'Central US' } else { 'East US 2' })
            for ($n = 1; $n -le $nodes; $n++) {
                $server = $(if ($platform -like '*Managed*') { 'sim-{0}-{1}-mi.database.windows.net' -f $entry.Code.ToLowerInvariant(), $tier.ToLowerInvariant() } else { 'SIM{0}{1}{2:00}' -f $tier, $entry.Code, $n })
                $cpu = $(if ($prodLike) { 16 } else { 4 })
                Add-SimServer -Name $server -Env $env -App $entry.App -Apm $entry.Apm -Crit $entry.Crit -Version $(if ($server -eq 'SIMQAPWB01') { '2019' } else { $version }) -Platform $platform -Location $location -Cpu $cpu -Memory ($cpu * 8) -Owner $entry.Owner

                # The QA document-imaging server is in the server sheet but carries no databases.
                if ($entry.Code -eq 'DIM' -and $tier -eq 'QA') { continue }

                $dbVersion = $(if ($server -eq 'SIMQAPWB01') { '14.0.3456.2' } else { $version })
                $native = Get-SimNative $dbVersion
                $names = $(if ($prodLike) { $dbNames } else { @($dbNames | Select-Object -First ([math]::Max(1, $dbNames.Count - 2))) })
                $k = 0
                foreach ($baseName in $names) {
                    $k++
                    $name = $(if ($prodLike) { $baseName } else { '{0}_{1}' -f $baseName, $tier })
                    $size = ($entry.Scale * 1024) * (0.4 + $random.NextDouble()) * $(if ($prodLike) { 1 } else { 0.25 })
                    $compat = $native
                    if ($entry.Code -in @('DWH', 'CMP')) { $compat = 100 }
                    if ($baseName -eq 'RPA_Archive') { $compat = 110 }
                    $recovery = $(if ($prodLike -and $entry.Code -ne 'DWH') { 'FULL' } else { 'SIMPLE' })
                    $pageVerify = $(if ($entry.Code -eq 'CMP') { 'TORN_PAGE_DETECTION' } else { 'CHECKSUM' })
                    $status = $(if ($baseName -eq 'RK_History' -and $tier -eq 'DR') { 'RESTORING' } else { 'ONLINE' })
                    $ag = $(if ($entry.Ag -and $prodLike) { 'AG_{0}_PRD' -f $entry.Code } else { '' })
                    $listener = $(if ($ag) { '{0}-prd-lsn' -f $entry.Code.ToLowerInvariant() } else { '' })
                    $class = $(if ($prodLike -and $k -le 2) { 'Confidential' } else { 'Internal' })
                    $collation = $(if ($entry.Code -eq 'DWH') { 'Latin1_General_CI_AS' } else { 'SQL_Latin1_General_CP1_CI_AS' })
                    Add-SimDatabase -Server $server -Database $name -Env $env -App $entry.App -Version $dbVersion -SizeMb $size -Status $status `
                        -Recovery $recovery -PageVerify $pageVerify -Compat $compat -Ag $ag -Listener $listener -Classification $class `
                        -Collation $collation -Platform $platform -Owner $entry.Owner -OlderSnapshot:($k % 2 -eq 0)
                }
                if ($platform -notlike '*Managed*') { Add-SimSystemDatabases -Server $server -Env $env -Version $dbVersion -Platform $platform }
            }
        }
    }

    # A shared PROD server carrying three named applications, one only the database sheet
    # knows about, and databases nobody has claimed.
    $shared = 'SIMPRDSHR01'
    Add-SimServer -Name $shared -Env 'PROD' -App 'Payroll Interface' -Apm '10988' -Crit 'Medium' -Version '13.0.6300.2' -Cpu 8 -Memory 64 -Owner 'J. Rivera'
    Add-SimServer -Name $shared -Env 'PROD' -App 'Batch Scheduler' -Apm '10099' -Crit 'Low' -Version '13.0.6300.2' -Cpu 8 -Memory 64 -Owner 'J. Rivera'
    Add-SimServer -Name $shared -Env 'PROD' -App 'SSRS Reporting' -Apm '' -Crit 'Low' -Version '13.0.6300.2' -Cpu 8 -Memory 64 -Owner 'J. Rivera'
    foreach ($spec in @(
            @('Payroll_Main', 'Payroll Interface', 8200), @('Payroll_Stage', 'Payroll Interface', 1900), @('JobStore', 'Batch Scheduler', 640),
            @('ReportServer', 'SSRS Reporting', 2300), @('ReportServerTempDB', 'SSRS Reporting', 410), @('FaxQueue', 'Legacy Fax Gateway', 780),
            @('FaxArchive', 'Legacy Fax Gateway', 15400), @('Temp_Extract_2019', 'Unknown', 22600), @('Old_Migration', 'Unknown', 5100), @('ZZ_Backup_Copy', '', 31800))) {
        Add-SimDatabase -Server $shared -Database $spec[0] -Env 'PROD' -App $spec[1] -Version '13.0.6300.2' -SizeMb $spec[2] -Compat 130 -Owner $(if ($spec[1] -match 'Unknown|^$') { '' } else { 'J. Rivera' })
    }
    Add-SimSystemDatabases -Server $shared -Env 'PROD' -Version '13.0.6300.2'

    # Unknown servers. One carries an APM ID with no name, which another row resolves.
    Add-SimServer -Name 'SIMPRDLEG01.corp.example.com' -Env 'Production' -App '' -Apm '10611' -Crit '' -Version '10.50.6000.34' -Platform 'On-Prem VMware' -Location 'Boston DC' -Cpu 4 -Memory 32 -Owner ''
    foreach ($spec in @(@('PricingHist_Old', 48000, 'ONLINE'), @('FPR_Legacy', 9300, 'ONLINE'), @('Archive2014', 61000, 'OFFLINE'))) {
        Add-SimDatabase -Server 'SIMPRDLEG01' -Database $spec[0] -Env 'Production' -App '' -Version '10.50.6000.34' -SizeMb $spec[1] -Status $spec[2] -Compat 100 -PageVerify $(if ($spec[0] -eq 'FPR_Legacy') { 'NONE' } else { 'CHECKSUM' }) -Platform 'On-Prem VMware'
    }
    Add-SimSystemDatabases -Server 'SIMPRDLEG01' -Env 'Production' -Version '10.50.6000.34' -Platform 'On-Prem VMware'

    Add-SimServer -Name 'SIMUATOLD02' -Env 'UAT' -App 'Unknown' -Apm '' -Crit '' -Version '2014' -Cpu 4 -Memory 32 -Owner ''
    Add-SimDatabase -Server 'SIMUATOLD02' -Database 'Test_Copy1' -Env 'UAT' -App 'N/A' -Version '12.0.6024.0' -SizeMb 4100 -Compat 120 -Recovery 'SIMPLE'
    Add-SimDatabase -Server 'SIMUATOLD02' -Database 'Test_Copy2' -Env 'UAT' -App 'N/A' -Version '12.0.6024.0' -SizeMb 3800 -Status 'SUSPECT' -Compat 120 -Recovery 'SIMPLE'

    Add-SimServer -Name 'SIMDEVTMP03' -Env 'Development' -App 'N/A' -Apm '' -Crit '' -Version 'SQL Server 2019 Developer' -Cpu 4 -Memory 16 -Owner ''

    # Only the database sheet knows this server.
    Add-SimDatabase -Server 'SIMPRDORPH09,1433' -Database 'Orphan_App_DB' -Env 'PROD' -App 'TBD' -Version '13.0.5026.0' -SizeMb 2700 -Compat 130
    Add-SimDatabase -Server 'SIMPRDORPH09,1433' -Database 'RPA_Recon_Work' -Env 'PROD' -App 'Retirement Plan Admin' -Version '13.0.5026.0' -SizeMb 1200 -Compat 130 -Owner 'A. Patel'

    # A database claimed by '#N/A' on a named application's server, and one only seen in an older snapshot.
    Add-SimDatabase -Server 'SIMPRDRPA01' -Database 'RPA_Legacy_Extract' -Env 'PROD' -App '#N/A' -Version '13.0.6300.2' -SizeMb 6400 -Compat 130 -Recovery 'SIMPLE'
    Add-SimDatabase -Server 'SIMPRDPWB01' -Database 'PWB_OldCampaign' -Env 'PROD' -App 'Participant Web' -Version '15.0.4355.3' -SizeMb 900 -Compat 150 -OnlyOlder

    [pscustomobject]@{ Servers = $serverRows; Databases = $databaseRows }
}

# =========================================================================================
# REGION  Read the sources
# =========================================================================================

$serverRecords = [System.Collections.Generic.List[hashtable]]::new()
$databaseRecords = [System.Collections.Generic.List[hashtable]]::new()
$sourceNotes = [System.Collections.Generic.List[object]]::new()

if ($Simulate) {
    Write-Host 'SIMULATION: building the page from a made-up estate. No workbook is read.' -ForegroundColor Yellow
    $simulation = Get-SimulatedInventory
    $serverRecords.AddRange($simulation.Servers)
    $databaseRecords.AddRange($simulation.Databases)
    $sourceNotes.Add([ordered]@{ source = 'SIMULATED > Servers'; role = 'Server'; rows = $serverRecords.Count; headerRow = 1; missing = '' })
    $sourceNotes.Add([ordered]@{ source = 'SIMULATED > Databases'; role = 'Database'; rows = $databaseRecords.Count; headerRow = 1; missing = '' })
    $sourceName = 'SIMULATED estate (built in, made up)'
    if ([string]::IsNullOrWhiteSpace($OutputHtml)) { $OutputHtml = Join-Path (Get-Location).Path 'SQL_Application_Map_SIMULATED.html' }
}
else {
    if ($SourceWorkbook.Count -eq 0) { $SourceWorkbook = @(Select-SourceFiles) }
    $paths = @(foreach ($path in $SourceWorkbook) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Source file not found: $path" }
        (Resolve-Path -LiteralPath $path).Path
    })
    if ([string]::IsNullOrWhiteSpace($OutputHtml)) { $OutputHtml = Join-Path (Split-Path -Parent $paths[0]) 'SQL_Application_Map.html' }
    $sourceName = (@($paths | ForEach-Object { [IO.Path]::GetFileName($_) })) -join ' + '

    $tables = [System.Collections.Generic.List[object]]::new()
    foreach ($path in $paths) {
        Write-Host "Reading $([IO.Path]::GetFileName($path))..." -ForegroundColor Cyan
        if ([IO.Path]::GetExtension($path) -eq '.csv') { $found = Read-CsvTable -Path $path } else { $found = Read-ExcelTables -Path $path }
        foreach ($table in $found) { $tables.Add($table) }
    }

    foreach ($table in $tables) {
        $label = '{0} > {1}' -f $table.File, $table.Sheet
        $headerRow = Get-HeaderRow $table
        $databaseColumns = Get-ColumnMap $table $headerRow $DatabaseFieldMap
        $serverColumns = Get-ColumnMap $table $headerRow $ServerFieldMap

        $role = 'Ignored'
        if ($DatabaseSheet -and $table.Sheet -eq $DatabaseSheet) { $role = 'Database' }
        elseif ($ServerSheet -and $table.Sheet -eq $ServerSheet) { $role = 'Server' }
        elseif ($databaseColumns.Contains('DatabaseName') -and $databaseColumns.Contains('ServerName')) {
            if (-not $DatabaseSheet) { $role = 'Database' }
        }
        elseif ($serverColumns.Contains('ServerName') -and ($serverColumns.Contains('ApplicationName') -or $serverColumns.Contains('APMAppID') -or
                $serverColumns.Contains('LogicalCPUCount') -or $serverColumns.Contains('PhysicalMemoryGB'))) {
            if (-not $ServerSheet) { $role = 'Server' }
        }

        $count = 0
        $missing = ''
        if ($role -eq 'Database') {
            $records = Read-Records -Table $table -HeaderRow $headerRow -ColumnMap $databaseColumns -FieldMap $DatabaseFieldMap -Source $label
            $databaseRecords.AddRange($records)
            $count = $records.Count
            $missing = (@($DatabaseFieldMap.Keys | Where-Object { -not $databaseColumns.Contains($_) -and $_ -notlike 'Inventory*' })) -join ', '
        }
        elseif ($role -eq 'Server') {
            $records = Read-Records -Table $table -HeaderRow $headerRow -ColumnMap $serverColumns -FieldMap $ServerFieldMap -Source $label
            $serverRecords.AddRange($records)
            $count = $records.Count
            $missing = (@($ServerFieldMap.Keys | Where-Object { -not $serverColumns.Contains($_) })) -join ', '
        }
        $colour = $(if ($role -eq 'Ignored') { 'DarkGray' } else { 'Green' })
        Write-Host ("  {0,-9} {1}  ({2:N0} rows, header on row {3})" -f $role, $label, $count, $headerRow) -ForegroundColor $colour
        if ($missing) { Write-Host "            columns not found: $missing" -ForegroundColor Yellow }
        $sourceNotes.Add([ordered]@{ source = $label; role = $role; rows = $count; headerRow = $headerRow; missing = $missing })
    }

    if ($serverRecords.Count -eq 0 -and $databaseRecords.Count -eq 0) {
        throw 'No server sheet and no database sheet were recognised. A database sheet needs ServerName and DatabaseName columns; a server sheet needs ServerName plus ApplicationName, APMAppID, LogicalCPUCount or PhysicalMemoryGB.'
    }
    if ($serverRecords.Count -eq 0) { Write-Host '  No server sheet recognised. The page is built from the database sheet alone.' -ForegroundColor Yellow }
    if ($databaseRecords.Count -eq 0) { Write-Host '  No database sheet recognised. The page will show servers but no databases.' -ForegroundColor Yellow }
}

if (-not [IO.Path]::IsPathRooted($OutputHtml)) { $OutputHtml = Join-Path (Get-Location).Path $OutputHtml }

# =========================================================================================
# REGION  Applications and servers from the server sheet
# =========================================================================================

$UnknownKey = '?UNKNOWN'
$apps = [ordered]@{}
$servers = [ordered]@{}
$hostIndex = @{}
$links = [ordered]@{}

function New-Application {
    param([string]$Key, [bool]$Unknown)
    [ordered]@{
        key = $Key; id = $apps.Count; unknown = $Unknown
        variants = [System.Collections.Hashtable]::new(); raw = @{}
        apm = (New-StringSet); criticality = (New-StringSet); businessUnit = (New-StringSet); costCenter = (New-StringSet)
        owner1 = (New-StringSet); owner2 = (New-StringSet)
        inServerSheet = $false; inDatabaseSheet = $false
    }
}

function Resolve-Application {
    param([string]$Name)
    if (-not $Name -or $Name -match $UnknownValuePattern) {
        $app = $apps[$UnknownKey]
        Add-Count $app.raw $(if ($Name) { $Name } else { '(blank)' })
        return $app
    }
    $key = $Name.ToUpperInvariant()
    if (-not $apps.Contains($key)) { $apps[$key] = New-Application $key $false }
    $app = $apps[$key]
    if ($app.variants.ContainsKey($Name)) { $app.variants[$Name]++ } else { $app.variants[$Name] = 1 }
    return $app
}

function New-Server {
    param([string]$Key, [string]$Name)
    [ordered]@{
        key = $Key; id = $servers.Count; name = $Name; host = (Get-HostKey $Key)
        env = ''; envRaw = ''; envNote = ''
        version = ''; dbVersions = @{}; dbEnvs = @{}; context = ''
        platform = ''; location = ''; cpu = $null; memoryGb = $null; port = ''; size = ''
        subscription = ''; vnet = ''; subnet = ''
        inServerSheet = $false; serverRows = 0; dbRows = 0
    }
}

$apps[$UnknownKey] = New-Application $UnknownKey $true

$copyFields = [ordered]@{ version = 'SQLVersion'; platform = 'Platform'; location = 'Location'; port = 'Port'; size = 'Size'; subscription = 'Subscription'; vnet = 'VirtualNetwork'; subnet = 'Subnet' }
$serverRowsSkipped = 0
foreach ($record in $serverRecords) {
    $name = $record['ServerName']
    if (-not $name) { $serverRowsSkipped++; continue }
    $key = Get-ServerKey $name
    if (-not $servers.Contains($key)) {
        $servers[$key] = New-Server $key $name
        $hostKey = $servers[$key].host
        if (-not $hostIndex.ContainsKey($hostKey)) { $hostIndex[$hostKey] = [System.Collections.Generic.List[string]]::new() }
        $hostIndex[$hostKey].Add($key)
    }
    $server = $servers[$key]
    $server.inServerSheet = $true
    $server.serverRows++

    $env = Get-CanonicalEnvironment $record['Environment']
    if (-not $server.env -or $server.env -eq 'UNKNOWN') { $server.env = $env; $server.envRaw = $record['Environment'] }
    elseif ($env -ne 'UNKNOWN' -and $env -ne $server.env) { $server.envNote = "Server sheet also lists it as $env" }
    foreach ($field in $copyFields.Keys) {
        if (-not $server[$field] -and $record[$copyFields[$field]]) { $server[$field] = $record[$copyFields[$field]] }
    }
    if ($null -eq $server.cpu) { $server.cpu = ConvertTo-Number $record['LogicalCPUCount'] }
    if ($null -eq $server.memoryGb) { $server.memoryGb = ConvertTo-Number $record['PhysicalMemoryGB'] }
    if (-not $server.context) { $server.context = $record['Platform'] }

    $app = Resolve-Application $record['ApplicationName']
    $app.inServerSheet = $true
    Add-KnownValue $app.apm $record['APMAppID']
    Add-KnownValue $app.criticality $record['ApplicationCriticality']
    Add-KnownValue $app.businessUnit $record['BusinessUnit']
    Add-KnownValue $app.costCenter $record['CostCenter']
    Add-KnownValue $app.owner1 $record['PrimaryITOwner']
    Add-KnownValue $app.owner2 $record['SecondaryITOwner']

    $linkKey = '{0}|{1}' -f $key, $app.key
    if (-not $links.Contains($linkKey)) { $links[$linkKey] = [ordered]@{ server = $key; app = $app.key; apm = (New-StringSet) } }
    Add-KnownValue $links[$linkKey].apm $record['APMAppID']
}

# A blank application row on a server that also has named applications says nothing new.
# The server is only "unknown" when the server sheet names no application for it at all.
$serversWithNamedApp = @{}
foreach ($link in $links.Values) { if ($link.app -ne $UnknownKey) { $serversWithNamedApp[$link.server] = $true } }
foreach ($linkKey in @($links.Keys)) {
    $link = $links[$linkKey]
    if ($link.app -eq $UnknownKey -and $serversWithNamedApp.ContainsKey($link.server)) { $links.Remove($linkKey) }
}

# =========================================================================================
# REGION  Databases
# =========================================================================================

# The database sheet may name a server as HOST.domain, HOST,port or HOST\INSTANCE where the
# server sheet says HOST. Match on the host when that is unambiguous; otherwise the server
# stands alone and is reported as missing from the server sheet.
$resolveCache = @{}
$hostMatched = [System.Collections.Generic.List[string]]::new()
function Resolve-ServerKey {
    param([string]$Key)
    if ($resolveCache.ContainsKey($Key)) { return $resolveCache[$Key] }
    $result = $Key
    if (-not ($servers.Contains($Key) -and $servers[$Key].inServerSheet)) {
        $hostKey = Get-HostKey $Key
        if ($hostIndex.ContainsKey($hostKey)) {
            $candidates = @($hostIndex[$hostKey])
            $instance = Get-InstancePart $Key
            $same = @($candidates | Where-Object { (Get-InstancePart $_) -eq $instance })
            if ($same.Count -ge 1) { $result = $same[0] }
            elseif ($candidates.Count -eq 1) { $result = $candidates[0] }
            if ($result -ne $Key) { $hostMatched.Add(('{0} -> {1}' -f $Key, $result)) }
        }
    }
    $resolveCache[$Key] = $result
    return $result
}

# One database is one row, however many inventory dates list it. The latest date wins.
$latest = @{}
$dbRowsSkipped = 0
$dbRepeats = 0
$snapshotRows = @{}
foreach ($record in $databaseRecords) {
    if (-not $record['DatabaseName']) { $dbRowsSkipped++; continue }
    $inventory = Get-InventoryDate $record
    Add-Count $snapshotRows $(if ($inventory.Text) { $inventory.Text } else { '(no date)' })
    $serverKey = $(if ($record['ServerName']) { Resolve-ServerKey (Get-ServerKey $record['ServerName']) } else { '(NO SERVER NAME)' })
    $record['_inv'] = $inventory.Text
    $record['_invSort'] = $inventory.Sort
    $record['_serverKey'] = $serverKey
    $dedupeKey = $serverKey + '|' + $record['DatabaseName'].ToUpperInvariant()
    if ($latest.ContainsKey($dedupeKey)) {
        $dbRepeats++
        if ($latest[$dedupeKey]['_invSort'] -gt $inventory.Sort) { continue }
    }
    $latest[$dedupeKey] = $record
}

$dbItems = [System.Collections.Generic.List[object]]::new()
foreach ($record in $latest.Values) {
    $key = $record['_serverKey']
    if (-not $servers.Contains($key)) {
        $servers[$key] = New-Server $key $(if ($record['ServerName']) { $record['ServerName'] } else { '(no server name)' })
    }
    $server = $servers[$key]
    $server.dbRows++
    Add-Count $server.dbVersions $record['SQLVersion']
    Add-Count $server.dbEnvs (Get-CanonicalEnvironment $record['Environment'])
    if (-not $server.platform -and $record['Platform']) { $server.platform = $record['Platform'] }
    if ($record['Platform'] -or $record['DatabaseType']) { $server.context = '{0} {1} {2}' -f $server.context, $record['Platform'], $record['DatabaseType'] }

    $isSystem = ($record['DatabaseType'] -match 'system') -or ($SystemDatabaseNames -contains $record['DatabaseName'].ToLowerInvariant())
    $appId = -1
    $rawApp = ''
    if (-not $isSystem) {
        $app = Resolve-Application $record['ApplicationName']
        $app.inDatabaseSheet = $true
        Add-KnownValue $app.businessUnit $record['BusinessUnit']
        Add-KnownValue $app.owner1 $record['PrimaryITOwner']
        Add-KnownValue $app.owner2 $record['SecondaryITOwner']
        $appId = $app.id
        if ($app.unknown) { $rawApp = $(if ($record['ApplicationName']) { $record['ApplicationName'] } else { '(blank)' }) }
    }
    $dbItems.Add([pscustomobject]@{ Record = $record; Server = $server; AppId = $appId; Raw = $rawApp; System = [bool]$isSystem })
}

# =========================================================================================
# REGION  Settle each server: environment, version, release
# =========================================================================================

$releaseByLabel = @{}
foreach ($release in $SqlReleases) { $releaseByLabel[$release.Label] = $release }
$usedReleases = @{}

foreach ($server in $servers.Values) {
    if (-not $server.env -or $server.env -eq 'UNKNOWN') {
        $fromDatabases = Get-MostCommon $server.dbEnvs
        if ($fromDatabases) { $server.env = $fromDatabases; $server.envRaw = '(from the database sheet)' }
        elseif (-not $server.env) { $server.env = 'UNKNOWN' }
    }
    $dbVersion = Get-MostCommon $server.dbVersions
    $sheetRelease = $(if ($server.version) { Get-SqlRelease $server.version $server.context } else { '' })
    $dbRelease = $(if ($dbVersion) { Get-SqlRelease $dbVersion $server.context } else { '' })
    $vague = @('', 'Unrecognised', 'Not recorded')
    if ($vague -notcontains $sheetRelease) { $release = $sheetRelease }
    elseif ($vague -notcontains $dbRelease) { $release = $dbRelease }
    elseif ($sheetRelease) { $release = $sheetRelease }
    else { $release = Get-SqlRelease '' $server.context }
    $server['release'] = $release
    $server['versionText'] = $(if ($server.version) { $server.version } else { $dbVersion })
    $server['versionNote'] = ''
    if ($vague -notcontains $sheetRelease -and $vague -notcontains $dbRelease -and $sheetRelease -ne $dbRelease) {
        $server['versionNote'] = "Server sheet: $($server.version) ($sheetRelease). Database sheet: $dbVersion ($dbRelease)."
    }
    $usedReleases[$release] = $true
}

# =========================================================================================
# REGION  Build the payload. Tables are columns + rows, which keeps the page small.
# =========================================================================================

function New-Table {
    param([string[]]$Columns)
    return [ordered]@{ cols = $Columns; rows = [System.Collections.Generic.List[object]]::new() }
}

$releaseTable = New-Table @('label', 'order', 'status', 'text', 'native', 'mainstream', 'extended')
foreach ($release in $SqlReleases) {
    $status = Get-ReleaseStatus $release
    $releaseTable.rows.Add([object[]]@($release.Label, [double]$release.Order, $status[0], $status[1], $release.Native, $release.Mainstream, $release.Extended))
}

$serverTable = New-Table @('id', 'name', 'host', 'env', 'envRaw', 'envNote', 'rel', 'ver', 'verNote', 'platform', 'location', 'cpu', 'mem', 'port', 'size', 'sub', 'vnet', 'subnet', 'inSrv', 'dbRows')
foreach ($s in $servers.Values) {
    $serverTable.rows.Add([object[]]@([int]$s.id, $s.name, $s.host, $s.env, $s.envRaw, $s.envNote, $s.release, $s.versionText, $s.versionNote,
            $s.platform, $s.location, $s.cpu, $s.memoryGb, $s.port, $s.size, $s.subscription, $s.vnet, $s.subnet, [bool]$s.inServerSheet, [int]$s.dbRows))
}

$appTable = New-Table @('id', 'name', 'unknown', 'apm', 'crit', 'bu', 'cc', 'own1', 'own2', 'inSrv', 'inDb')
foreach ($app in $apps.Values) {
    $name = 'UNKNOWN APPLICATION'
    if (-not $app.unknown) { $name = [string](($app.variants.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key) }
    $appTable.rows.Add([object[]]@([int]$app.id, $name, [bool]$app.unknown, (Join-Set $app.apm), (Join-Set $app.criticality), (Join-Set $app.businessUnit),
            (Join-Set $app.costCenter), (Join-Set $app.owner1), (Join-Set $app.owner2), [bool]$app.inServerSheet, [bool]$app.inDatabaseSheet))
}

$linkTable = New-Table @('server', 'app', 'apm')
foreach ($link in $links.Values) {
    $linkTable.rows.Add([object[]]@([int]$servers[$link.server].id, [int]$apps[$link.app].id, (Join-Set $link.apm)))
}

$dbTable = New-Table @('server', 'app', 'name', 'env', 'mb', 'status', 'type', 'sys', 'compat', 'recovery', 'cls', 'coll', 'ag', 'lsn', 'pv', 'inv', 'raw', 'own')
$dbUser = 0
$dbSystem = 0
foreach ($item in $dbItems) {
    $record = $item.Record
    $env = Get-CanonicalEnvironment $record['Environment']
    if ($env -eq 'UNKNOWN') { $env = $item.Server.env }
    $owner = $record['PrimaryITOwner']
    if ($record['SecondaryITOwner']) { $owner = ('{0} / {1}' -f $owner, $record['SecondaryITOwner']).Trim(' /') }
    if ($item.System) { $dbSystem++ } else { $dbUser++ }
    $dbTable.rows.Add([object[]]@([int]$item.Server.id, [int]$item.AppId, $record['DatabaseName'], $env, (ConvertTo-Number $record['DatabaseSizeMB']),
            $record['DatabaseStatus'].ToUpperInvariant(), $record['DatabaseType'], $item.System, (ConvertTo-Number $record['CompatibilityLevel']),
            $record['RecoveryModel'].ToUpperInvariant(), $record['DatabaseClassification'], $record['DatabaseCollationName'],
            (Clear-UnknownValue $record['AGName']), (Clear-UnknownValue $record['ListenerName']), $record['PageVerifyOption'].ToUpperInvariant(),
            $record['_inv'], $item.Raw, $owner))
}

$unknownValues = [System.Collections.Generic.List[object]]::new()
foreach ($entry in ($apps[$UnknownKey].raw.GetEnumerator() | Sort-Object Value -Descending)) { $unknownValues.Add([object[]]@([string]$entry.Key, [int]$entry.Value)) }
$snapshots = [System.Collections.Generic.List[object]]::new()
foreach ($entry in ($snapshotRows.GetEnumerator() | Sort-Object Name -Descending)) { $snapshots.Add([object[]]@([string]$entry.Key, [int]$entry.Value)) }
$latestInventory = [string](@($snapshotRows.Keys | Where-Object { $_ -ne '(no date)' } | Sort-Object -Descending) | Select-Object -First 1)

$palettes = @{
    Manulife = @{ Brand = '#34384B'; BrandDark = '#23263A'; BrandSoft = '#EEF0F4'; Accent = '#00A758'; AccentDark = '#00875D'; AccentSoft = '#E6F6EE' }
    Slate    = @{ Brand = '#2F3640'; BrandDark = '#1F242C'; BrandSoft = '#EDEFF2'; Accent = '#4A7C9B'; AccentDark = '#3A6480'; AccentSoft = '#E9F1F6' }
}
$theme = $palettes[$Palette]

$payload = [ordered]@{
    meta     = [ordered]@{
        simulated = [bool]$Simulate
        source    = $sourceName
        generated = $ReportDate.ToString('yyyy-MM-dd HH:mm', $Invariant)
        stamp     = $ReportDate.ToString('yyyyMMdd_HHmm', $Invariant)
        latest    = $latestInventory
        warnMonths = $SupportWarningMonths
        unknownValues = $unknownValues
        recon     = [ordered]@{
            sources = $sourceNotes
            serverRows = $serverRecords.Count; serverRowsSkipped = $serverRowsSkipped
            dbRows = $databaseRecords.Count; dbRowsSkipped = $dbRowsSkipped; dbRepeats = $dbRepeats
            dbKept = $dbItems.Count; dbUser = $dbUser; dbSystem = $dbSystem
            snapshots = $snapshots; hostMatched = @($hostMatched | Select-Object -Unique)
        }
    }
    releases = $releaseTable
    servers  = $serverTable
    apps     = $appTable
    links    = $linkTable
    dbs      = $dbTable
}

# JavaScriptSerializer is part of .NET Framework and far faster than ConvertTo-Json on a
# large estate. If it is unavailable, ConvertTo-Json does the same job more slowly.
try {
    Add-Type -AssemblyName System.Web.Extensions
    $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $serializer.MaxJsonLength = [int]::MaxValue
    $serializer.RecursionLimit = 64
    $json = $serializer.Serialize($payload)
}
catch {
    $json = $payload | ConvertTo-Json -Depth 10 -Compress
}
$json = $json.Replace('</', '<\/')

# =========================================================================================
# REGION  The page. Red is reserved for UNKNOWN and for things that are broken; environment
#         colours are identity only, so PROD is blue rather than red.
# =========================================================================================

$htmlStyle = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SQL Application Map</title>
<style>
:root{
  --brand:__T_BRAND__;--brand-dark:__T_BRAND_DARK__;--brand-soft:__T_BRAND_SOFT__;
  --accent:__T_ACCENT__;--accent-dark:__T_ACCENT_DARK__;--accent-soft:__T_ACCENT_SOFT__;
  --bg:#F4F6F8;--panel:#fff;--ink:#1F2333;--ink2:#4A5265;--ink3:#6F778A;
  --line:#DCE1E8;--line2:#EDF0F4;
  --crit:#C6093B;--crit-dark:#8F0629;--crit-bg:#FDECEF;--crit-line:#EFBCC7;
  --serious:#B4470B;--serious-bg:#FFF1E6;--serious-line:#F3C99C;
  --warn:#8A6A00;--warn-bg:#FFF8E1;--warn-line:#E6D4A0;
  --good:#00875D;--good-bg:#E6F6EE;--good-line:#A8DCC5;
  --info:#0067B9;--info-bg:#E8F1FA;--info-line:#B5D2EC;
  --shadow:0 6px 20px rgba(31,35,51,.08);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:13px;line-height:1.4}
button{font-family:inherit}

/* ---- header and jump bar ---- */
.top{background:linear-gradient(105deg,var(--brand-dark),var(--brand));color:#fff;padding:16px 24px;display:flex;justify-content:space-between;align-items:center;gap:20px;border-bottom:3px solid var(--accent);flex-wrap:wrap}
.top h1{margin:0;font-size:21px;font-weight:700;letter-spacing:.2px}
.top p{margin:5px 0 0;color:#C9CEDA;font-size:11.5px;max-width:760px}
.top-stats{display:flex;gap:9px;flex-wrap:wrap}
.top-stat{min-width:106px;text-align:center;background:#ffffff14;border:1px solid #ffffff2e;border-radius:8px;padding:8px 12px}
.top-stat b{display:block;font-size:21px;line-height:1.1}
.top-stat span{display:block;color:#C9CEDA;font-size:9.5px;text-transform:uppercase;letter-spacing:.5px;margin-top:3px}
.top-stat.bad{background:var(--crit);border-color:#fff;box-shadow:0 0 0 2px var(--crit)}
.top-stat.bad span{color:#fff}
.sim{display:inline-block;background:var(--warn-bg);color:var(--warn);padding:2px 8px;border-radius:4px;font-weight:800;letter-spacing:.4px;margin-right:6px}
.jump{position:sticky;top:0;z-index:30;display:flex;gap:4px;align-items:center;padding:7px 18px;background:#fff;border-bottom:1px solid var(--line);box-shadow:0 2px 8px rgba(31,35,51,.06);flex-wrap:wrap}
.jump a{color:var(--ink2);text-decoration:none;font-weight:650;padding:6px 11px;border-radius:6px;border:1px solid transparent}
.jump a:hover{border-color:var(--accent);color:var(--brand)}
.jump a.j-unk{color:var(--crit)}
.jump .sp{flex:1}
.jump .asof{font-size:11px;color:var(--ink3)}

/* ---- surfaces ---- */
.wrap{max-width:1680px;margin:0 auto;padding:16px 18px 28px;display:grid;gap:16px}
section{scroll-margin-top:60px;min-width:0}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow);padding:18px}
.sec-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:12px}
h2.sec{margin:0;color:var(--brand);font-size:18px}
h3.sub{margin:20px 0 9px;color:var(--brand);font-size:12.5px;text-transform:uppercase;letter-spacing:.5px;padding-bottom:6px;border-bottom:2px solid var(--line)}
.note{font-size:11px;color:var(--ink3);line-height:1.55;margin:4px 0 0}
.nul{color:var(--ink3)}
.mono{font-variant-numeric:tabular-nums}
.small{font-size:10.5px}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
.tools{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.search,select.search{padding:8px 10px;border:1px solid var(--line);border-radius:7px;font:inherit;outline:none;background:#fff;color:var(--ink)}
.search:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
input.search{width:280px}
.btn{border:1px solid var(--line);background:#fff;color:var(--ink2);border-radius:7px;padding:8px 13px;font-weight:650;cursor:pointer}
.btn:hover{border-color:var(--accent);color:var(--brand)}
.chip-filter{display:inline-flex;align-items:center;gap:8px;background:var(--info-bg);border:1px solid var(--info-line);color:var(--info);border-radius:999px;padding:5px 6px 5px 12px;font-weight:650;margin-bottom:10px}
.chip-filter button{border:0;background:var(--info);color:#fff;border-radius:999px;width:20px;height:20px;cursor:pointer;line-height:1}

/* ---- the UNKNOWN block: as loud as the page gets ---- */
.unk{display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:24px;align-items:center;background:var(--crit);color:#fff;border-radius:12px;padding:22px 26px;border:3px solid var(--crit-dark);box-shadow:0 12px 30px rgba(198,9,59,.35)}
.unk-mark{width:104px;height:104px;border-radius:50%;background:#fff;color:var(--crit);display:grid;place-items:center;font-size:68px;font-weight:900;line-height:1;box-shadow:0 0 0 6px rgba(255,255,255,.25)}
.unk-kicker{font-size:11px;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;opacity:.9}
.unk h2{margin:2px 0 4px;font-size:38px;font-weight:900;letter-spacing:1px;line-height:1.05}
.unk p{margin:0;font-size:13px;max-width:900px;line-height:1.5}
.unk-figs{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}
.unk-fig{background:rgba(255,255,255,.13);border:1px solid rgba(255,255,255,.4);border-radius:9px;padding:9px 14px;min-width:132px}
.unk-fig b{display:block;font-size:30px;line-height:1.05}
.unk-fig span{display:block;font-size:10px;text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-top:3px}
.unk-fig small{display:block;font-size:10.5px;opacity:.85;margin-top:2px}
.unk-raw{margin-top:12px;font-size:11.5px;line-height:1.7}
.unk-raw .rv{display:inline-block;background:#fff;color:var(--crit);border-radius:4px;padding:1px 7px;margin:0 4px 3px 0;font-weight:700}
.unk-actions{display:flex;flex-direction:column;gap:9px}
.unk-actions button{border:2px solid #fff;border-radius:8px;padding:11px 16px;font-weight:800;cursor:pointer;font-size:12.5px;white-space:nowrap}
.btn-w{background:#fff;color:var(--crit)}
.btn-o{background:transparent;color:#fff}
.btn-w:hover,.btn-o:hover{box-shadow:0 0 0 3px rgba(255,255,255,.35)}
.allclear{background:var(--good-bg);border:1px solid var(--good-line);color:var(--good);border-radius:10px;padding:14px 18px;font-weight:650;font-size:13px}

/* ---- tags, chips, flags ---- */
.env{display:inline-block;padding:2px 7px;border-radius:4px;color:#fff;background:var(--env-c,#5F6673);font-size:9.5px;font-weight:700;letter-spacing:.4px;white-space:nowrap;vertical-align:middle}
.rel{display:inline-flex;align-items:center;gap:4px;padding:2px 7px 2px 5px;border-radius:999px;font-size:10px;font-weight:650;white-space:nowrap;border:1px solid;vertical-align:middle;margin:1px 3px 1px 0}
.rel i,.st i{font-style:normal;font-weight:900;font-size:9px;width:13px;height:13px;border-radius:50%;display:inline-grid;place-items:center;color:#fff}
.rel b{font-weight:700;margin-left:2px}
.st{display:inline-flex;align-items:center;gap:6px;padding:3px 9px 3px 5px;border-radius:999px;font-size:10.5px;font-weight:650;white-space:nowrap;border:1px solid}
.st-oos{color:var(--crit);background:var(--crit-bg);border-color:var(--crit-line)}.st-oos i{background:var(--crit)}
.st-ending{color:var(--serious);background:var(--serious-bg);border-color:var(--serious-line)}.st-ending i{background:var(--serious)}
.st-ext{color:var(--warn);background:var(--warn-bg);border-color:var(--warn-line)}.st-ext i{background:#B08900}
.st-main{color:var(--good);background:var(--good-bg);border-color:var(--good-line)}.st-main i{background:var(--good)}
.st-paas{color:var(--info);background:var(--info-bg);border-color:var(--info-line)}.st-paas i{background:var(--info)}
.st-unk{color:var(--ink2);background:var(--line2);border-color:var(--line);border-style:dashed}.st-unk i{background:var(--ink3)}
.flag{display:inline-block;padding:2px 7px;border-radius:4px;font-size:9.5px;font-weight:700;white-space:nowrap;border:1px solid;margin:1px 3px 1px 0;vertical-align:middle}
.flag.crit{color:var(--crit);background:var(--crit-bg);border-color:var(--crit-line)}
.flag.warn{color:var(--warn);background:var(--warn-bg);border-color:var(--warn-line)}
.flag.info{color:var(--info);background:var(--info-bg);border-color:var(--info-line)}
.flag.good{color:var(--good);background:var(--good-bg);border-color:var(--good-line)}
.crit-c{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;white-space:nowrap;border:1px solid var(--brand)}
.crit-c.c0{background:var(--brand);color:#fff}
.crit-c.c1{color:var(--brand);background:var(--brand-soft)}
.crit-c.c2,.crit-c.c3,.crit-c.c4{color:var(--ink2);border-color:var(--line);background:#fff}

/* ---- tables ---- */
.tbl-wrap{overflow:auto;border:1px solid var(--line);border-radius:9px;background:#fff}
.table{width:100%;border-collapse:collapse;font-size:11.5px}
.table th{position:sticky;top:0;background:var(--brand);color:#fff;text-align:left;padding:8px 9px;font-weight:650;z-index:2;white-space:nowrap;vertical-align:bottom}
.table td{padding:7px 9px;border-bottom:1px solid var(--line2);vertical-align:middle}
.table td.num,.table th.num{text-align:right;white-space:nowrap}
.table tbody tr:hover td{background:var(--brand-soft)}
.table th.sortable{cursor:pointer;user-select:none}
.table th.sortable:hover{background:var(--brand-dark)}
.table th small{display:block;font-weight:400;font-size:9px;color:#C9CEDA;margin-top:2px}
.apps-wrap{max-height:620px}
.apps td.an{font-weight:650;color:var(--brand);min-width:190px}
.apps tbody tr{cursor:pointer}
.apps tbody tr.sel td{background:var(--accent-soft)}
.apps tbody tr.sel td.an{box-shadow:inset 4px 0 0 var(--accent)}
.apps td.ec{text-align:center;min-width:64px;border-left:1px solid var(--line2)}
.apps td.ec b{display:block;font-size:14px;color:var(--brand);line-height:1.1}
.apps td.ec span{display:block;font-size:9.5px;color:var(--ink3)}
.apps td.ec.empty{color:var(--line)}
.apps th.ec-h{text-align:center;border-left:1px solid var(--brand-dark)}
.apps tr.unk-row td{background:var(--crit)!important;color:#fff;border-bottom:2px solid var(--crit-dark)}
.apps tr.unk-row td.ec b,.apps tr.unk-row td.ec span,.apps tr.unk-row td b{color:#fff}
.apps tr.unk-row td.ec.empty{color:rgba(255,255,255,.4)}
.apps tr.unk-row .rel{background:#fff}
.apps tr.unk-row .flag{background:#fff;color:var(--crit);border-color:#fff}
.unk-word{font-size:13px;font-weight:900;letter-spacing:.8px;color:#fff}
.apps tr.unk-row.sel td{box-shadow:inset 0 0 0 2px #fff}
.rels,.flags{display:flex;flex-wrap:wrap;gap:2px}

/* ---- application detail ---- */
.det-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap}
.det-head h2{margin:0;font-size:24px;color:var(--brand)}
.det-head.red{background:var(--crit);color:#fff;border-radius:10px;padding:16px 18px;margin:-4px -4px 0}
.det-head.red h2{color:#fff;font-size:30px;font-weight:900;letter-spacing:.8px}
.det-head.red .note{color:#fff;opacity:.92;font-size:12px}
.meta{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:9px;margin:14px 0 4px}
.meta-box{border:1px solid var(--line);background:var(--brand-soft);border-radius:7px;padding:8px 11px;min-width:0}
.meta-box b{display:block;font-size:9px;color:var(--ink3);text-transform:uppercase;letter-spacing:.4px}
.meta-box>span{display:block;font-size:12px;margin-top:3px;overflow-wrap:anywhere}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-top:14px}
.et{border:1px solid var(--line);border-top:5px solid var(--env-c,#5F6673);border-radius:9px;padding:11px 13px;background:#fff}
.et.none{background:#FAFBFC;border-top-color:var(--line)}
.et.none .et-n{color:var(--line)}
.et-n{font-size:30px;font-weight:700;color:var(--brand);line-height:1.05;margin-top:7px}
.et-n small{font-size:11px;font-weight:600;color:var(--ink3);margin-left:4px}
.et-s{font-size:11px;color:var(--ink2);margin-top:4px}
.et .rels{margin-top:7px}

/* ---- relationship map ---- */
.topo{padding:18px 4px 6px}
.flow-root{display:flex;justify-content:center}
.root-circle{width:164px;height:164px;border-radius:50%;background:linear-gradient(135deg,var(--brand-dark),var(--brand));color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:16px;border:5px solid var(--brand-soft);box-shadow:0 10px 26px rgba(31,35,51,.22)}
.root-circle.unk{background:var(--crit);border-color:var(--crit-bg);box-shadow:0 10px 26px rgba(198,9,59,.4)}
.root-circle b{font-size:13.5px;font-weight:800;line-height:1.25;overflow-wrap:anywhere}
.root-circle span{display:block;font-size:10px;color:#E1E4EC;margin-top:6px;line-height:1.45}
.pointer-line{width:2px;height:30px;background:var(--line);margin:0 auto;position:relative}
.pointer-line:after{content:"";position:absolute;bottom:-1px;left:-4px;border-left:5px solid transparent;border-right:5px solid transparent;border-top:8px solid var(--line)}
.env-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:13px;margin-top:12px}
.env-card{border:1px solid var(--line);border-top:5px solid var(--env-c,#5F6673);border-radius:10px;background:#fff;padding:14px;box-shadow:0 2px 10px rgba(31,35,51,.05);min-width:0}
.eh{display:flex;align-items:center;gap:13px}
.env-circle{width:62px;height:62px;flex:0 0 62px;border-radius:50%;background:var(--env-c,#5F6673);color:#fff;display:flex;align-items:center;justify-content:center;text-align:center;font-size:10px;font-weight:800;letter-spacing:.3px;padding:5px;word-break:break-word;box-shadow:0 0 0 3px #fff,0 0 0 5px var(--env-c,#5F6673)}
.env-sum b{display:block;font-size:14px;color:var(--brand)}
.env-sum span{display:block;font-size:11px;color:var(--ink3);margin-top:3px}
.inst-stack{margin:13px 0 0 30px;padding-left:22px;border-left:2px solid var(--line)}
.inst-node{position:relative;padding:10px 0}
.inst-node+.inst-node{border-top:1px dashed var(--line2)}
.inst-node:before{content:"";position:absolute;left:-22px;top:18px;width:22px;height:2px;background:var(--line)}
.inst-node:after{content:"";position:absolute;left:-28px;top:13px;width:12px;height:12px;border-radius:50%;background:#fff;border:2.5px solid var(--env-c,#5F6673)}
.inst-top{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}
.srv-name{font-size:13px;font-weight:750;color:var(--brand);overflow-wrap:anywhere}
.inst-meta{font-size:10.5px;color:var(--ink3);margin-top:3px;line-height:1.5;overflow-wrap:anywhere}
.inst-meta b{color:var(--ink2);font-weight:650}
.inst-fig{font-size:10px;color:var(--ink3);white-space:nowrap;text-align:right}
.inst-fig b{color:var(--brand);font-size:15px}
.lead{margin-top:6px;font-size:10.5px;color:var(--crit);background:var(--crit-bg);border:1px solid var(--crit-line);border-radius:6px;padding:5px 8px}
.dbchips{display:flex;flex-wrap:wrap;gap:5px;margin-top:9px}
.dbchip{display:inline-flex;align-items:baseline;gap:6px;border:1px solid var(--line);border-left:4px solid var(--good);border-radius:6px;background:#FBFCFE;padding:4px 8px;max-width:100%}
.dbchip b{font-size:10.5px;font-weight:650;color:var(--ink);overflow-wrap:anywhere}
.dbchip span{font-size:9.5px;color:var(--ink3);white-space:nowrap}
.dbchip em{font-style:normal;font-size:9px;font-weight:800;color:var(--crit)}
.dbchip.bad{border-left-color:var(--crit);background:var(--crit-bg)}
.dbchip.unk{border-color:var(--crit);border-left-color:var(--crit);background:var(--crit-bg)}
.dbchip.unk b{color:var(--crit)}
.more{border:1px dashed var(--line);background:#fff;border-radius:6px;padding:4px 9px;font-size:10px;color:var(--ink2);cursor:pointer}
.unk-inline{margin-top:7px;display:inline-block;background:var(--crit);color:#fff;font-size:10px;font-weight:800;border-radius:5px;padding:3px 8px;letter-spacing:.3px}
.map-legend{display:flex;gap:14px;flex-wrap:wrap;align-items:center;padding:10px 4px 0;font-size:10.5px;color:var(--ink2)}
.map-legend span{display:inline-flex;align-items:center;gap:6px}
.map-legend i{width:11px;height:11px;border-radius:3px;display:inline-block}

/* ---- estate ---- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px}
.kpi{background:#fff;border:1px solid var(--line);border-top:4px solid var(--k-c,var(--accent));border-radius:9px;padding:12px 14px}
.kpi b{display:block;color:var(--brand);font-size:26px;line-height:1.05}
.kpi .lbl{display:block;margin-top:6px;color:var(--ink2);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px}
.kpi .sub{display:block;margin-top:4px;color:var(--ink3);font-size:10.5px}
.matrix td.cell{text-align:center;cursor:pointer;min-width:74px;border:2px solid #fff}
.matrix td.cell b{display:block;font-size:15px;line-height:1.1}
.matrix td.cell span{display:block;font-size:9.5px;opacity:.85}
.matrix td.cell.zero{color:var(--line);cursor:default;background:#fff}
.matrix td.cell:not(.zero):hover{box-shadow:inset 0 0 0 2px var(--brand)}
.matrix td.tot{font-weight:700;color:var(--brand);text-align:center;background:var(--brand-soft)}
.matrix tfoot td{font-weight:700;background:var(--brand-soft);color:var(--brand);text-align:center}
.matrix th.rel-h{min-width:150px}
.ramp{display:inline-flex;vertical-align:middle;margin:0 6px}
.ramp i{width:16px;height:10px;display:inline-block}
.cfg{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px}
.cfg-card{border:1px solid var(--line);border-left:5px solid var(--c-c,var(--ink3));border-radius:8px;padding:11px 13px;background:#fff}
.cfg-card .h{display:flex;justify-content:space-between;gap:8px;align-items:center;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--ink2)}
.cfg-card .n{font-size:24px;font-weight:700;color:var(--brand);margin-top:6px;line-height:1.1}
.cfg-card .n small{font-size:11px;color:var(--ink3);font-weight:600;margin-left:5px}
.cfg-card .d{font-size:10.5px;color:var(--ink3);margin-top:5px;line-height:1.5}
details.rc{border-top:1px solid var(--line2);padding:8px 0}
details.rc summary{cursor:pointer;font-weight:650;color:var(--brand)}
details.rc .list{font-size:11px;color:var(--ink2);margin-top:6px;line-height:1.7;max-height:220px;overflow:auto}
.rc-line{display:flex;justify-content:space-between;gap:12px;padding:7px 0;border-bottom:1px solid var(--line2)}
.rc-line b{color:var(--brand);font-variant-numeric:tabular-nums}
.empty{color:var(--ink3);text-align:center;padding:30px}
.foot{font-size:10.5px;color:var(--ink3);padding:10px 18px 18px;text-align:center}

@media(max-width:1100px){
  .meta{grid-template-columns:repeat(2,minmax(0,1fr))}
  .grid2{grid-template-columns:1fr}
  .unk{grid-template-columns:1fr}.unk-mark{width:70px;height:70px;font-size:44px}
  .unk-actions{flex-direction:row;flex-wrap:wrap}
  .env-grid{grid-template-columns:1fr}
  input.search{width:100%}
}
@media print{
  .jump,.tools,.unk-actions,.more{display:none!important}
  body{background:#fff}.panel{box-shadow:none}
  .apps-wrap,.tbl-wrap{max-height:none!important;overflow:visible}
  .unk,.apps tr.unk-row td,.det-head.red,.root-circle.unk,.unk-inline{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
</style>
</head>
'@

$htmlBody = @'
<body>
<header class="top">
  <div>
    <h1>SQL Application Map</h1>
    <p id="topLine"></p>
  </div>
  <div class="top-stats" id="globalStats"></div>
</header>
<nav class="jump">
  <a href="#unknown" class="j-unk">Unknown</a>
  <a href="#apps">Applications</a>
  <a href="#detail">Application detail</a>
  <a href="#estate">SQL Server estate</a>
  <a href="#recon">Reconciliation</a>
  <span class="sp"></span><span class="asof" id="asOf"></span>
</nav>
<main class="wrap">
  <section id="unknown"></section>

  <section id="apps" class="panel">
    <div class="sec-head">
      <div>
        <h2 class="sec">Applications &rarr; SQL Server</h2>
        <p class="note">Each environment column shows <b>SQL servers</b> (large) and <b>databases</b> (small) for that application. Click a row to open its map below.</p>
      </div>
      <div class="tools">
        <input id="q" class="search" placeholder="Search application, APM ID, server or database" aria-label="Search">
        <select id="critFilter" class="search" aria-label="Criticality"></select>
        <select id="showFilter" class="search" aria-label="Show"></select>
      </div>
    </div>
    <div id="cellFilter"></div>
    <div class="tbl-wrap apps-wrap"><table class="table apps" id="appTable"></table></div>
    <p class="note" id="appCount"></p>
  </section>

  <section id="detail" class="panel"></section>
  <section id="estate" class="panel"></section>
  <section id="recon" class="panel"></section>
</main>
<footer class="foot" id="foot"></footer>
'@

$htmlScriptModel = @'
<script>
"use strict";
const P = __PAYLOAD__;
const META = P.meta;
const rowsOf = t => t.rows.map(r => { const o = {}; t.cols.forEach((c, i) => { o[c] = r[i]; }); return o; });
const RELEASES = rowsOf(P.releases);
const REL = new Map(RELEASES.map(r => [r.label, r]));
const SERVERS = rowsOf(P.servers);
const APPS = rowsOf(P.apps);
const LINKS = rowsOf(P.links);
const DBS = rowsOf(P.dbs);
const UNK = APPS.find(a => a.unknown);

/* ---------------------------------------------------------------------------
   Helpers. A value the workbook does not hold prints as a dash, never as zero.
--------------------------------------------------------------------------- */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const has = v => v !== null && v !== undefined && v !== '';
const nf = n => Number(n).toLocaleString('en-US');
const sum = a => a.reduce((t, x) => t + x, 0);
const byName = (a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
const plural = (n, one, many) => `${nf(n)} ${n === 1 ? one : (many || one + 's')}`;
const DASH = '<span class="nul">&ndash;</span>';
const nz = v => has(v) ? esc(v) : DASH;
const uniq = a => [...new Set(a)];
const nameList = (items, max) => items.slice(0, max).join(', ') + (items.length > max ? ` +${items.length - max} more` : '');
const mbOf = dbs => sum(dbs.filter(d => d.mb != null).map(d => d.mb));
function sizeText(mb) {
  if (mb === null || mb === undefined) return '\u2013';
  if (mb < 1024) return `${Math.round(mb)} MB`;
  const gb = mb / 1024;
  if (gb < 1024) return `${gb < 10 ? gb.toFixed(1) : nf(Math.round(gb))} GB`;
  const tb = gb / 1024;
  return `${tb >= 100 ? nf(Math.round(tb)) : tb.toFixed(2)} TB`;
}

/* Environment colour is identity, not status, so PROD is blue: red belongs to UNKNOWN. */
const ENV_ORDER = ['PROD','DR','PREPROD','UAT','SIT','QA','TEST','DEV','SANDBOX','NONPROD'];
const CORE_ENVS = ['PROD','DR','UAT','QA','DEV'];
const ENV_HEX = {PROD:'#1C5CAB',DR:'#5B3FA8',PREPROD:'#0E6B78',UAT:'#9A4A0C',SIT:'#0F7A5A',QA:'#0F7A5A',TEST:'#0F7A5A',
  DEV:'#4F6420',SANDBOX:'#6B5B45',NONPROD:'#5F6673',UNKNOWN:'#8A909C'};
const envRank = e => { const i = ENV_ORDER.indexOf(e); return i >= 0 ? i : (e === 'UNKNOWN' ? 99 : 50); };
const envSort = (a, b) => envRank(a) - envRank(b) || byName(a, b);
const envHex = e => ENV_HEX[e] || '#5F6673';
const envLabel = e => e === 'UNKNOWN' ? 'NO ENV' : e;
const envTag = e => `<span class="env" style="--env-c:${envHex(e)}">${esc(envLabel(e))}</span>`;

/* SQL Server release and Microsoft support status. Every status carries a mark and words. */
const STATUS = {
  oos:        { cls: 'st-oos',    mark: '!', word: 'Out of support' },
  ending:     { cls: 'st-ending', mark: '!', word: 'Support ending' },
  extended:   { cls: 'st-ext',    mark: '~', word: 'Extended support' },
  mainstream: { cls: 'st-main',   mark: '&#10003;', word: 'Mainstream' },
  paas:       { cls: 'st-paas',   mark: '&#8734;', word: 'Evergreen PaaS' },
  unknown:    { cls: 'st-unk',    mark: '?', word: 'Unknown' }
};
const relInfo = l => REL.get(l) || { label: l, status: 'unknown', text: 'Unknown', order: 0, native: null };
const relOrder = l => relInfo(l).order;
const relStatus = l => relInfo(l).status;
const relShort = l => String(l).replace(/^SQL Server /, '');
const relChip = (label, count) => {
  const r = relInfo(label), s = STATUS[r.status];
  return `<span class="rel ${s.cls}" title="${esc(label + ' \u2013 ' + r.text)}"><i>${s.mark}</i>${esc(relShort(label))}${count != null ? `<b>&times;${nf(count)}</b>` : ''}</span>`;
};
const statusChip = label => { const r = relInfo(label), s = STATUS[r.status]; return `<span class="st ${s.cls}"><i>${s.mark}</i>${esc(r.text)}</span>`; };
const nativeOf = s => relInfo(s.rel).native;

const critRank = c => {
  const t = String(c || '').toLowerCase();
  if (/crit|mission|tier\s*[01]\b|^[01]$|^p[01]\b/.test(t)) return 0;
  if (/high|tier\s*2|^2$|^p2\b/.test(t)) return 1;
  if (/med|moderate|tier\s*3|^3$|^p3\b/.test(t)) return 2;
  if (/low|tier\s*4|^4$|^p4\b/.test(t)) return 3;
  return has(c) ? 4 : 5;
};
const critChip = c => has(c) ? String(c).split(', ').map(x => `<span class="crit-c c${critRank(x)}">${esc(x)}</span>`).join(' ') : DASH;

/* ---------------------------------------------------------------------------
   The model. A server belongs to an application when either sheet says so.
   A server no sheet ties to a named application belongs to UNKNOWN.
--------------------------------------------------------------------------- */
SERVERS.forEach(s => Object.assign(s, { dbs: [], sys: 0, apps: new Set(), sheetApps: new Set(), unkApm: [] }));
APPS.forEach(a => Object.assign(a, { servers: new Set(), dbs: [], sheetServers: new Set() }));
DBS.forEach(d => {
  const s = SERVERS[d.server];
  d.srv = s;
  if (d.sys) { s.sys++; return; }
  s.dbs.push(d);
  s.apps.add(d.app);
  APPS[d.app].dbs.push(d);
  APPS[d.app].servers.add(s.id);
});
LINKS.forEach(l => {
  const s = SERVERS[l.server], a = APPS[l.app];
  s.apps.add(a.id); s.sheetApps.add(a.id); a.servers.add(s.id); a.sheetServers.add(s.id);
  if (a.unknown && has(l.apm)) s.unkApm.push(...String(l.apm).split(', '));
});
SERVERS.forEach(s => {
  s.known = [...s.apps].filter(id => !APPS[id].unknown);
  if (!s.known.length) { s.apps.add(UNK.id); UNK.servers.add(s.id); }
  s.unkDbs = s.dbs.filter(d => APPS[d.app].unknown).length;
  s.mb = mbOf(s.dbs);
});
const USER_DBS = DBS.filter(d => !d.sys);
const ESTATE_ENVS = uniq(SERVERS.map(s => s.env)).sort(envSort);
const APM_APP = new Map();
APPS.forEach(a => { if (!a.unknown && has(a.apm)) String(a.apm).split(', ').forEach(id => { if (!APM_APP.has(id)) APM_APP.set(id, a.name); }); });

/* Per application: servers and databases in each environment, keyed by the server's
   environment so the map always nests database under server under environment. */
const statCache = new Map();
function appStats(a) {
  if (statCache.has(a.id)) return statCache.get(a.id);
  const envs = new Map();
  const bucket = e => { if (!envs.has(e)) envs.set(e, { env: e, servers: [], dbs: [], mb: 0 }); return envs.get(e); };
  [...a.servers].forEach(id => bucket(SERVERS[id].env).servers.push(SERVERS[id]));
  a.dbs.forEach(d => { const b = bucket(d.srv.env); b.dbs.push(d); if (d.mb != null) b.mb += d.mb; });
  const rels = new Map();
  [...a.servers].forEach(id => { const r = SERVERS[id].rel; rels.set(r, (rels.get(r) || 0) + 1); });
  const oos = [...a.servers].filter(id => relStatus(SERVERS[id].rel) === 'oos').length;
  const sized = a.dbs.filter(d => d.mb != null);
  const shared = new Set();
  [...a.servers].forEach(id => SERVERS[id].apps.forEach(x => { if (x !== a.id && !APPS[x].unknown) shared.add(x); }));
  const st = { envs, rels, oos, mb: mbOf(a.dbs), sized: sized.length, shared, flags: [] };
  if (!a.unknown) {
    const prod = envs.get('PROD');
    if (oos) st.flags.push(['crit', `${oos} on out-of-support SQL`, 'SQL Server release past the end of Microsoft extended support.']);
    if (prod && prod.servers.length && !envs.has('DR')) st.flags.push(['warn', 'No DR recorded', 'No server for this application is marked DR in the inventory. DR may exist outside it, for example as an availability-group replica marked PROD.']);
    if (!a.inSrv) st.flags.push(['info', 'Not in server sheet', 'Only the database sheet names this application, so it has no APM ID, criticality or cost center.']);
    else if (!has(a.apm)) st.flags.push(['warn', 'No APM ID', 'The server sheet names this application without an APM ID.']);
    if (!a.dbs.length) st.flags.push(['info', 'No databases recorded', 'The server sheet names this application, but no database in the database sheet does.']);
  }
  statCache.set(a.id, st);
  return st;
}

/* Where the inventory points at an owner for an unattributed server. A lead, not a fact. */
function leadFor(s) {
  const parts = [];
  const sheet = [...s.sheetApps].filter(id => !APPS[id].unknown).map(id => APPS[id].name).sort(byName);
  if (sheet.length) parts.push(`Server sheet ties this server to ${sheet.join(', ')}`);
  else {
    const fromDbs = uniq(s.dbs.filter(d => !APPS[d.app].unknown).map(d => APPS[d.app].name)).sort(byName);
    if (fromDbs.length) parts.push(`Other databases on this server belong to ${fromDbs.join(', ')}`);
  }
  uniq(s.unkApm).forEach(id => parts.push(APM_APP.has(id)
    ? `APM ID ${id} on the server sheet is ${APM_APP.get(id)} elsewhere`
    : `APM ID ${id} on the server sheet, no application name found for it`));
  return parts.join('; ');
}

function downloadCsv(fileName, header, rows) {
  const q = v => { const t = v === null || v === undefined ? '' : String(v); return /[",\r\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t; };
  const text = [header, ...rows].map(r => r.map(q).join(',')).join('\r\n');
  const blob = new Blob(['\ufeff' + text], { type: 'text/csv;charset=utf-8' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => { URL.revokeObjectURL(link.href); link.remove(); }, 500);
}
const kpiHtml = ([label, value, sub, colour]) =>
  `<div class="kpi" style="--k-c:${colour}"><b>${value}</b><span class="lbl">${esc(label)}</span><span class="sub">${sub}</span></div>`;

/* ---------------------------------------------------------------------------
   Header and the UNKNOWN block
--------------------------------------------------------------------------- */
function renderHeader() {
  const known = APPS.filter(a => !a.unknown);
  const stats = [
    ['Applications', nf(known.length), ''],
    ['SQL servers', nf(SERVERS.length), ''],
    ['Databases', nf(USER_DBS.length), ''],
    ['Data', esc(sizeText(mbOf(USER_DBS))), ''],
    ['Unknown-app DBs', nf(UNK.dbs.length), UNK.dbs.length ? 'bad' : '']
  ];
  document.getElementById('globalStats').innerHTML = stats.map(([l, v, c]) => `<div class="top-stat ${c}"><b>${v}</b><span>${esc(l)}</span></div>`).join('');
  document.getElementById('topLine').innerHTML = (META.simulated ? '<span class="sim">SIMULATED</span> Every name and figure here is made up. ' : '')
    + `Application &rarr; SQL Server &rarr; database, from <b>${esc(META.source)}</b>.`
    + ` ${nf(USER_DBS.length)} user databases plus ${nf(DBS.length - USER_DBS.length)} system databases, counted once each.`;
  document.getElementById('asOf').textContent = `Inventory as of ${META.latest || 'an unrecorded date'} \u00b7 generated ${META.generated}`;
  if (META.simulated) document.title = 'SIMULATED - ' + document.title;
}

function unknownServers() {
  const list = [...UNK.servers].map(id => SERVERS[id]);
  return { bare: list.filter(s => !s.known.length), partly: list.filter(s => s.known.length && s.unkDbs > 0) };
}

function renderUnknown() {
  const el = document.getElementById('unknown');
  const dbs = UNK.dbs;
  const { bare, partly } = unknownServers();
  if (!dbs.length && !bare.length) {
    el.innerHTML = '<div class="allclear">&#10003; Every SQL server and every user database in the inventory is tied to a named application.</div>';
    return;
  }
  const pct = USER_DBS.length ? (dbs.length / USER_DBS.length * 100) : 0;
  const prodDbs = dbs.filter(d => d.srv.env === 'PROD').length;
  const prodSrv = bare.filter(s => s.env === 'PROD').length;
  const withLead = dbs.filter(d => leadFor(d.srv)).length + bare.filter(s => leadFor(s)).length;
  const raw = (META.unknownValues || []).map(([v, n]) => `<span class="rv">${esc(v)}</span>${nf(n)}`).join(' &nbsp; ');
  el.innerHTML = `<div class="unk">
    <div class="unk-mark" aria-hidden="true">?</div>
    <div>
      <div class="unk-kicker">Not tied to any application</div>
      <h2>UNKNOWN APPLICATION</h2>
      <p>The inventory does not say which application these belong to. Nobody can be called when they fail, and they cannot be planned into an upgrade or a DR test until an owner is named.</p>
      <div class="unk-figs">
        <div class="unk-fig"><b>${nf(dbs.length)}</b><span>databases</span><small>${pct.toFixed(1)}% of all user databases</small></div>
        <div class="unk-fig"><b>${nf(prodDbs)}</b><span>of them in PROD</span><small>${esc(sizeText(mbOf(dbs)))} of data in total</small></div>
        <div class="unk-fig"><b>${nf(bare.length)}</b><span>servers, no application</span><small>${nf(prodSrv)} in PROD</small></div>
        <div class="unk-fig"><b>${nf(partly.length)}</b><span>servers carrying some</span><small>alongside a named application</small></div>
      </div>
      ${raw ? `<div class="unk-raw">Recorded in the workbook as: ${raw}</div>` : ''}
      ${withLead ? `<div class="unk-raw">${nf(withLead)} of these have a lead: the server or its APM ID is tied to a named application elsewhere in the inventory. Each lead is listed beside its record below.</div>` : ''}
    </div>
    <div class="unk-actions">
      <button class="btn-w" id="unkShow">Show them &darr;</button>
      <button class="btn-o" id="unkCsv">Export CSV</button>
    </div>
  </div>`;
  document.getElementById('unkShow').onclick = () => selectApp(UNK.id, true);
  document.getElementById('unkCsv').onclick = exportUnknown;
}

function exportUnknown() {
  const rows = [];
  UNK.dbs.slice().sort((a, b) => envSort(a.srv.env, b.srv.env) || byName(a.srv.name, b.srv.name) || byName(a.name, b.name)).forEach(d => {
    const s = d.srv;
    rows.push(['Database', d.name, s.name, s.env, d.mb ?? '', d.status, d.raw, s.rel, s.ver, s.platform, s.location, d.own, leadFor(s)]);
  });
  unknownServers().bare.sort((a, b) => envSort(a.env, b.env) || byName(a.name, b.name)).forEach(s => {
    rows.push(['Server', '', s.name, s.env, s.mb || '', '', '', s.rel, s.ver, s.platform, s.location, '', leadFor(s)]);
  });
  downloadCsv(`Unknown_application_${META.stamp}.csv`,
    ['Record', 'Database', 'Server', 'Environment', 'Size MB', 'Status', 'Recorded application', 'SQL release', 'SQL version', 'Platform', 'Location', 'IT owner (database sheet)', 'Lead'], rows);
}
'@

$htmlScriptApps = @'
/* ---------------------------------------------------------------------------
   Applications table. UNKNOWN is pinned to the top whatever the sort.
--------------------------------------------------------------------------- */
const state = { q: '', crit: 'ALL', show: 'ALL', cell: null, sort: { key: 'name', dir: 1 }, selected: null };

const SHOW = {
  ALL:    ['All applications', () => true],
  oos:    ['On out-of-support SQL', (a, st) => st.oos > 0],
  ending: ['On SQL whose support is ending', (a) => [...a.servers].some(id => relStatus(SERVERS[id].rel) === 'ending')],
  nodr:   ['PROD but no DR recorded', (a, st) => st.flags.some(f => f[1] === 'No DR recorded')],
  gaps:   ['Missing from one of the sheets', (a) => !a.unknown && (!a.inSrv || !a.inDb)],
  shared: ['Sharing a server with other applications', (a, st) => st.shared.size > 0]
};

function initFilters() {
  const crits = uniq(APPS.filter(a => !a.unknown && has(a.crit)).flatMap(a => String(a.crit).split(', '))).sort((x, y) => critRank(x) - critRank(y) || byName(x, y));
  document.getElementById('critFilter').innerHTML = '<option value="ALL">All criticality</option>' + crits.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('');
  document.getElementById('showFilter').innerHTML = Object.entries(SHOW).map(([k, v]) => `<option value="${k}">${esc(v[0])}</option>`).join('');
}

function matchesSearch(a, q) {
  if (!q) return true;
  if (a.name.toLowerCase().includes(q) || String(a.apm || '').toLowerCase().includes(q)) return true;
  if ([...a.servers].some(id => SERVERS[id].name.toLowerCase().includes(q))) return true;
  return a.dbs.some(d => d.name.toLowerCase().includes(q));
}

function filteredApps() {
  const q = state.q.trim().toLowerCase();
  return APPS.filter(a => {
    if (a.unknown && !a.dbs.length && !a.servers.size) return false;
    const st = appStats(a);
    if (!matchesSearch(a, q)) return false;
    if (state.crit !== 'ALL' && (a.unknown || !String(a.crit || '').split(', ').includes(state.crit))) return false;
    if (!SHOW[state.show][1](a, st)) return false;
    if (state.cell) {
      const c = state.cell;
      if (![...a.servers].some(id => SERVERS[id].rel === c.rel && (c.env === 'ALL' || SERVERS[id].env === c.env))) return false;
    }
    return true;
  });
}

function sortValue(a, key) {
  const st = appStats(a);
  switch (key) {
    case 'name': return a.name.toLowerCase();
    case 'apm': return has(a.apm) ? String(a.apm) : '~';
    case 'crit': return critRank(String(a.crit || '').split(', ')[0]);
    case 'servers': return a.servers.size;
    case 'dbs': return a.dbs.length;
    case 'mb': return st.mb;
    case 'rel': return Math.min(...[...st.rels.keys()].map(relOrder), 99999);
    case 'flags': return st.flags.length;
  }
  if (key.startsWith('env:')) { const b = st.envs.get(key.slice(4)); return b ? b.servers.length * 1e6 + b.dbs.length : -1; }
  return 0;
}

function renderAppTable() {
  const list = filteredApps();
  const { key, dir } = state.sort;
  list.sort((x, y) => {
    if (x.unknown !== y.unknown) return x.unknown ? -1 : 1;
    const vx = sortValue(x, key), vy = sortValue(y, key);
    const c = typeof vx === 'string' ? vx.localeCompare(vy) : vx - vy;
    return (c * dir) || byName(x.name, y.name);
  });

  const th = (k, label, cls = '', extra = '') => `<th class="sortable ${cls}" data-sort="${k}" title="Sort">${label}${state.sort.key === k ? (dir > 0 ? ' &#9650;' : ' &#9660;') : ''}${extra}</th>`;
  const head = `<thead><tr>
    ${th('name', 'Application')}${th('apm', 'APM ID')}${th('crit', 'Criticality')}
    ${ESTATE_ENVS.map(e => th('env:' + e, envTag(e), 'ec-h', '<small>servers / dbs</small>')).join('')}
    ${th('servers', 'Servers', 'num')}${th('dbs', 'Databases', 'num')}${th('mb', 'Data', 'num')}
    ${th('rel', 'SQL Server versions')}${th('flags', 'Flags')}
  </tr></thead>`;

  const rows = list.map(a => {
    const st = appStats(a);
    const cells = ESTATE_ENVS.map(e => {
      const b = st.envs.get(e);
      if (!b || (!b.servers.length && !b.dbs.length)) return '<td class="ec empty">&middot;</td>';
      return `<td class="ec" title="${esc(`${a.name} in ${envLabel(e)}: ${plural(b.servers.length, 'SQL server')}, ${plural(b.dbs.length, 'database')}, ${sizeText(b.mb)}`)}"><b>${nf(b.servers.length)}</b><span>${nf(b.dbs.length)} db</span></td>`;
    }).join('');
    const rels = [...st.rels.entries()].sort((x, y) => relOrder(y[0]) - relOrder(x[0])).map(([l, n]) => relChip(l, n)).join('');
    const flags = a.unknown
      ? '<span class="flag crit">No owner recorded</span>'
      : st.flags.map(([k, t, tip]) => `<span class="flag ${k}" title="${esc(tip)}">${esc(t)}</span>`).join('')
        + (st.shared.size ? `<span class="flag info" title="${esc('Shares a SQL server with: ' + [...st.shared].map(id => APPS[id].name).sort(byName).join(', '))}">Shares servers with ${st.shared.size}</span>` : '');
    return `<tr class="${a.unknown ? 'unk-row' : ''} ${state.selected === a.id ? 'sel' : ''}" data-app="${a.id}" tabindex="0">
      <td class="an">${a.unknown ? '<span class="unk-word">&#9888; UNKNOWN APPLICATION</span>' : esc(a.name)}</td>
      <td class="mono">${a.unknown ? '' : nz(a.apm)}</td>
      <td>${a.unknown ? '' : critChip(a.crit)}</td>
      ${cells}
      <td class="num"><b>${nf(a.servers.size)}</b></td>
      <td class="num"><b>${nf(a.dbs.length)}</b></td>
      <td class="num mono">${st.sized ? esc(sizeText(st.mb)) : DASH}</td>
      <td><div class="rels">${rels}</div></td>
      <td><div class="flags">${flags}</div></td></tr>`;
  }).join('');

  const table = document.getElementById('appTable');
  table.innerHTML = head + `<tbody>${rows || `<tr><td colspan="${ESTATE_ENVS.length + 8}" class="empty">No application matches these filters.</td></tr>`}</tbody>`;
  table.querySelectorAll('th[data-sort]').forEach(h => h.onclick = () => {
    const k = h.dataset.sort;
    if (state.sort.key === k) state.sort.dir *= -1;
    else state.sort = { key: k, dir: (k === 'name' || k === 'apm' || k === 'crit' || k === 'rel') ? 1 : -1 };
    renderAppTable();
  });
  table.querySelectorAll('tr[data-app]').forEach(r => {
    r.onclick = () => selectApp(Number(r.dataset.app), true);
    r.onkeydown = e => { if (e.key === 'Enter') selectApp(Number(r.dataset.app), true); };
  });

  const known = APPS.filter(a => !a.unknown).length;
  const shownKnown = list.filter(a => !a.unknown).length;
  document.getElementById('appCount').textContent = `Showing ${nf(shownKnown)} of ${nf(known)} named applications`
    + (list.some(a => a.unknown) ? ', plus UNKNOWN.' : '.')
    + ' A server or database shared by several applications is counted under each of them.';

  const chip = document.getElementById('cellFilter');
  if (state.cell) {
    chip.innerHTML = `<span class="chip-filter">Only applications on ${esc(state.cell.rel)}${state.cell.env === 'ALL' ? '' : ' in ' + esc(envLabel(state.cell.env))}<button type="button" title="Clear" id="clearCell">&times;</button></span>`;
    document.getElementById('clearCell').onclick = () => { state.cell = null; renderAppTable(); };
  } else chip.innerHTML = '';
}

function selectApp(id, scroll) {
  state.selected = id;
  renderAppTable();
  renderDetail();
  if (scroll) document.getElementById('detail').scrollIntoView({ behavior: 'smooth', block: 'start' });
}
'@

$htmlScriptDetail = @'
/* ---------------------------------------------------------------------------
   Application detail: header, environment tiles, relationship map, databases
--------------------------------------------------------------------------- */
const CHIP_LIMIT = 36;
const expanded = new Set();

function dbChip(d) {
  const bad = has(d.status) && d.status !== 'ONLINE';
  const unk = APPS[d.app].unknown;
  const tip = `${d.name} on ${d.srv.name}\n${sizeText(d.mb)} \u00b7 ${d.status || 'status not recorded'} \u00b7 compatibility ${d.compat ?? '\u2013'} \u00b7 ${d.recovery || 'recovery model not recorded'}`
    + (has(d.ag) ? `\nAvailability group ${d.ag}` : '') + (unk ? `\nApplication recorded as: ${d.raw || '(blank)'}` : '');
  return `<span class="dbchip ${bad ? 'bad' : ''} ${unk ? 'unk' : ''}" title="${esc(tip)}"><b>${esc(d.name)}</b><span>${esc(sizeText(d.mb))}</span>${bad ? `<em>${esc(d.status)}</em>` : ''}</span>`;
}

function serverNodeHtml(s, a) {
  const mine = s.dbs.filter(d => d.app === a.id).sort((x, y) => byName(x.name, y.name));
  const others = s.known.filter(id => id !== a.id).map(id => APPS[id].name).sort(byName);
  const meta = [
    has(s.platform) ? esc(s.platform) : '',
    has(s.location) ? esc(s.location) : '',
    s.cpu != null ? `<b>${nf(s.cpu)}</b> cores` : '',
    s.mem != null ? `<b>${nf(s.mem)}</b> GB memory` : '',
    has(s.port) ? `port ${esc(s.port)}` : '',
    has(s.size) ? esc(s.size) : ''
  ].filter(Boolean).join(' &middot; ');
  const net = [s.sub, s.vnet, s.subnet].filter(has).map(esc).join(' &rsaquo; ');
  const ags = uniq(mine.map(d => d.ag).filter(has)), lsns = uniq(mine.map(d => d.lsn).filter(has));
  const key = `${a.id}|${s.id}`;
  const shown = expanded.has(key) ? mine : mine.slice(0, CHIP_LIMIT);
  const lead = a.unknown ? leadFor(s) : '';
  return `<div class="inst-node">
    <div class="inst-top">
      <div style="min-width:0">
        <span class="srv-name">${esc(s.name)}</span> ${relChip(s.rel)}
        ${!s.inSrv ? '<span class="flag warn" title="This server appears only in the database sheet">not in server sheet</span>' : ''}
        ${has(s.verNote) ? `<span class="flag warn" title="${esc(s.verNote)}">versions disagree</span>` : ''}
        <div class="inst-meta">${has(s.ver) ? `Version <b>${esc(s.ver)}</b>` : 'Version not recorded'}${meta ? ' &middot; ' + meta : ''}</div>
        ${net ? `<div class="inst-meta">${net}</div>` : ''}
        ${ags.length ? `<div class="inst-meta">Availability group <b>${esc(ags.join(', '))}</b>${lsns.length ? ` &middot; listener <b>${esc(lsns.join(', '))}</b>` : ''}</div>` : ''}
        ${others.length ? `<div class="inst-meta">Also hosts: ${esc(nameList(others, 5))}</div>` : ''}
        ${lead ? `<div class="lead">Lead: ${esc(lead)}</div>` : ''}
      </div>
      <div class="inst-fig"><b>${nf(mine.length)}</b> db<br>${esc(sizeText(mbOf(mine)))}${s.sys ? `<br>+${s.sys} system` : ''}</div>
    </div>
    <div class="dbchips">${mine.length ? shown.map(dbChip).join('') : '<span class="note">No database of this application is recorded on this server.</span>'}
      ${mine.length > CHIP_LIMIT ? `<button type="button" class="more" data-more="${key}">${expanded.has(key) ? 'Show fewer' : `+${mine.length - CHIP_LIMIT} more`}</button>` : ''}</div>
    ${!a.unknown && s.unkDbs ? `<span class="unk-inline">+ ${plural(s.unkDbs, 'database')} on this server with NO application</span>` : ''}
  </div>`;
}

function mapHtml(a, st) {
  if (!st.envs.size) return '<div class="empty">No server or database is recorded for this application.</div>';
  const cards = [...st.envs.values()].sort((x, y) => envSort(x.env, y.env)).map(b => {
    const servers = b.servers.slice().sort((x, y) => byName(x.name, y.name));
    return `<div class="env-card" style="--env-c:${envHex(b.env)}">
      <div class="eh"><div class="env-circle">${esc(envLabel(b.env))}</div>
        <div class="env-sum"><b>${plural(b.servers.length, 'SQL server')}</b><span>${plural(b.dbs.length, 'database')} &middot; ${esc(sizeText(b.mb))}</span></div></div>
      <div class="inst-stack">${servers.map(s => serverNodeHtml(s, a)).join('')}</div></div>`;
  }).join('');
  const envText = [...st.envs.values()].sort((x, y) => envSort(x.env, y.env)).map(b => `${envLabel(b.env)} ${b.servers.length}`).join(' \u00b7 ');
  return `<div class="flow-root"><div class="root-circle ${a.unknown ? 'unk' : ''}">
      <b>${a.unknown ? 'UNKNOWN APPLICATION' : esc(a.name)}</b>
      <span>${plural(a.servers.size, 'SQL server')}<br>${plural(a.dbs.length, 'database')}<br>${esc(envText)}</span></div></div>
    <div class="pointer-line"></div>
    <div class="env-grid">${cards}</div>
    <div class="map-legend">
      <span><i style="background:var(--good)"></i>database ONLINE</span>
      <span><i style="background:var(--crit)"></i>not ONLINE, or no application recorded</span>
      <span style="margin-left:auto;color:var(--ink3)">The chip beside each server is its SQL Server release, coloured by Microsoft support status. Hover any database for its details.</span>
    </div>`;
}

function dbRowsHtml(a, dbs) {
  return dbs.map(d => {
    const s = d.srv, native = nativeOf(s);
    const low = d.compat != null && native != null && d.compat < native;
    return `<tr>
      <td><b>${esc(d.name)}</b></td>
      <td>${esc(s.name)}</td>
      <td>${envTag(s.env)}${d.env !== s.env && d.env !== 'UNKNOWN' ? ` <span class="flag warn" title="The database sheet lists this database as ${esc(d.env)}">DB sheet: ${esc(d.env)}</span>` : ''}</td>
      <td>${relChip(s.rel)}</td>
      <td class="num mono">${esc(sizeText(d.mb))}</td>
      <td>${has(d.status) ? (d.status === 'ONLINE' ? 'ONLINE' : `<span class="flag crit">${esc(d.status)}</span>`) : DASH}</td>
      <td class="mono">${d.compat != null ? Math.round(d.compat) : DASH}${low ? ` <span class="flag warn" title="${esc(s.rel)} creates new databases at ${native}">below ${native}</span>` : ''}</td>
      <td>${nz(d.recovery)}</td>
      <td>${has(d.ag) ? esc(d.ag) + (has(d.lsn) ? `<div class="note">${esc(d.lsn)}</div>` : '') : DASH}</td>
      <td>${nz(d.cls)}</td>
      <td class="small">${nz(d.coll)}</td>
      <td>${has(d.pv) && d.pv !== 'CHECKSUM' ? `<span class="flag warn">${esc(d.pv)}</span>` : nz(d.pv)}</td>
      <td class="mono">${nz(d.inv)}</td>
      ${a.unknown ? `<td><span class="flag crit">${esc(d.raw || '(blank)')}</span></td><td class="small">${nz(leadFor(s))}</td><td class="small">${nz(d.own)}</td>` : ''}
    </tr>`;
  }).join('');
}

function renderDetail() {
  const el = document.getElementById('detail');
  const a = APPS[state.selected];
  if (!a) { el.innerHTML = '<div class="empty">Click an application above.</div>'; return; }
  const st = appStats(a);

  let head;
  if (a.unknown) {
    const raw = (META.unknownValues || []).map(([v, n]) => `${v} (${nf(n)})`).join(', ');
    head = `<div class="det-head red"><div>
        <h2>&#9888; UNKNOWN APPLICATION</h2>
        <p class="note">Every database and server below is in the inventory with no application recorded against it${raw ? ` &mdash; written as ${esc(raw)}` : ''}. Where the other sheet ties the same server or APM ID to a named application, that lead is shown in red beside it. A lead is a place to start asking, not an answer.</p>
      </div><button class="btn" id="detCsv">Export CSV</button></div>`;
  } else {
    const sharedNames = [...st.shared].map(id => APPS[id].name).sort(byName);
    const box = (label, value) => `<div class="meta-box"><b>${esc(label)}</b><span>${value}</span></div>`;
    head = `<div class="det-head"><div><h2>${esc(a.name)}</h2>
        <p class="note">${[...st.envs.keys()].sort(envSort).map(envTag).join(' ')} &nbsp; ${plural(a.servers.size, 'SQL server')} &middot; ${plural(a.dbs.length, 'database')} &middot; ${esc(sizeText(st.mb))}${a.dbs.length > st.sized ? ` (${nf(a.dbs.length - st.sized)} without a size)` : ''}</p>
        <div class="flags" style="margin-top:6px">${st.flags.map(([k, t, tip]) => `<span class="flag ${k}" title="${esc(tip)}">${esc(t)}</span>`).join('')}</div>
      </div><button class="btn" id="detCsv">Export CSV</button></div>
      <div class="meta">
        ${box('APM ID', nz(a.apm))}${box('Criticality', critChip(a.crit))}${box('Business unit', nz(a.bu))}${box('Cost center', nz(a.cc))}
        ${box('Primary IT owner', nz(a.own1))}${box('Secondary IT owner', nz(a.own2))}
        ${box('Found in', `${a.inSrv ? '&#10003;' : '&#10007;'} server sheet &nbsp; ${a.inDb ? '&#10003;' : '&#10007;'} database sheet`)}
        ${box('Shares SQL servers with', sharedNames.length ? esc(nameList(sharedNames, 6)) : 'No other application')}
      </div>`;
  }

  const envList = uniq([...CORE_ENVS, ...st.envs.keys()]).sort(envSort);
  const prod = st.envs.get('PROD');
  const tiles = envList.map(e => {
    const b = st.envs.get(e);
    if (!b) {
      const noDr = e === 'DR' && prod && prod.servers.length && !a.unknown;
      return `<div class="et none"><div>${envTag(e)}</div><div class="et-n">0</div><div class="et-s">${noDr ? '<span class="flag warn">No DR recorded</span>' : 'none recorded'}</div></div>`;
    }
    const rels = new Map();
    b.servers.forEach(s => rels.set(s.rel, (rels.get(s.rel) || 0) + 1));
    return `<div class="et" style="--env-c:${envHex(e)}"><div>${envTag(e)}</div>
      <div class="et-n">${nf(b.servers.length)}<small>${b.servers.length === 1 ? 'server' : 'servers'}</small></div>
      <div class="et-s"><b>${nf(b.dbs.length)}</b> databases &middot; ${esc(sizeText(b.mb))}</div>
      <div class="rels">${[...rels.entries()].sort((x, y) => relOrder(y[0]) - relOrder(x[0])).map(([l, n]) => relChip(l, n)).join('')}</div></div>`;
  }).join('');

  const dbs = a.dbs.slice().sort((x, y) => envSort(x.srv.env, y.srv.env) || byName(x.srv.name, y.srv.name) || byName(x.name, y.name));
  const LIMIT = 1500;
  el.innerHTML = `${head}
    <div class="tiles">${tiles}</div>
    <h3 class="sub">Relationship map &mdash; application, environment, SQL server, databases</h3>
    <div class="topo">${mapHtml(a, st)}</div>
    <h3 class="sub">Every database (${nf(dbs.length)})</h3>
    ${dbs.length ? `<div class="tbl-wrap" style="max-height:520px"><table class="table">
      <thead><tr><th>Database</th><th>SQL server</th><th>Env</th><th>Release</th><th class="num">Size</th><th>Status</th><th>Compat</th><th>Recovery</th><th>AG / listener</th><th>Classification</th><th>Collation</th><th>Page verify</th><th>Inventoried</th>
        ${a.unknown ? '<th>Recorded as</th><th>Lead</th><th>IT owner</th>' : ''}</tr></thead>
      <tbody>${dbRowsHtml(a, dbs.slice(0, LIMIT))}</tbody></table></div>
      ${dbs.length > LIMIT ? `<p class="note">Showing the first ${nf(LIMIT)}. Export CSV for all ${nf(dbs.length)}.</p>` : ''}` : '<div class="empty">No database recorded.</div>'}`;

  el.querySelectorAll('[data-more]').forEach(b => b.onclick = () => {
    const k = b.dataset.more;
    if (expanded.has(k)) expanded.delete(k); else expanded.add(k);
    renderDetail();
  });
  document.getElementById('detCsv').onclick = () => {
    if (a.unknown) { exportUnknown(); return; }
    downloadCsv(`${a.name.replace(/[^A-Za-z0-9]+/g, '_')}_${META.stamp}.csv`,
      ['Application', 'APM ID', 'Database', 'SQL server', 'Environment', 'SQL release', 'SQL version', 'Support status', 'Size MB', 'Status', 'Compatibility', 'Recovery model', 'AG', 'Listener', 'Classification', 'Collation', 'Page verify', 'Inventoried', 'Platform', 'Location'],
      dbs.map(d => [a.name, a.apm, d.name, d.srv.name, d.srv.env, d.srv.rel, d.srv.ver, relInfo(d.srv.rel).text, d.mb ?? '', d.status, d.compat ?? '', d.recovery, d.ag, d.lsn, d.cls, d.coll, d.pv, d.inv, d.srv.platform, d.srv.location]));
  };
}
'@

$htmlScriptEstate = @'
/* ---------------------------------------------------------------------------
   SQL Server estate: version layout, platform, location, configuration
--------------------------------------------------------------------------- */
const RAMP = ['#E3EEFC','#cde2fb','#b7d3f6','#9ec5f4','#86b6ef','#6da7ec','#5598e7','#3987e5','#2a78d6','#256abf','#1c5cab','#184f95'];
const rampStep = (v, max) => Math.min(RAMP.length - 1, Math.round(Math.sqrt(v / max) * (RAMP.length - 1)));

function matrixHtml() {
  const rels = uniq(SERVERS.map(s => s.rel)).sort((x, y) => relOrder(y) - relOrder(x));
  const cell = new Map();
  SERVERS.forEach(s => {
    const k = s.rel + '|' + s.env;
    if (!cell.has(k)) cell.set(k, { servers: 0, dbs: 0, mb: 0, cores: 0 });
    const c = cell.get(k);
    c.servers++; c.dbs += s.dbs.length; c.mb += s.mb; c.cores += s.cpu || 0;
  });
  const max = Math.max(1, ...[...cell.values()].map(c => c.servers));
  const rows = rels.map(rel => {
    const list = SERVERS.filter(s => s.rel === rel);
    const builds = [...list.reduce((m, s) => m.set(s.ver || '(blank)', (m.get(s.ver || '(blank)') || 0) + 1), new Map()).entries()]
      .sort((x, y) => y[1] - x[1]);
    const cells = ESTATE_ENVS.map(e => {
      const c = cell.get(rel + '|' + e);
      if (!c) return '<td class="cell zero">&middot;</td>';
      const step = rampStep(c.servers, max);
      return `<td class="cell" data-rel="${esc(rel)}" data-env="${esc(e)}" style="background:${RAMP[step]};color:${step >= 7 ? '#fff' : 'var(--ink)'}"
        title="${esc(`${rel} in ${envLabel(e)}: ${plural(c.servers, 'server')}, ${plural(c.dbs, 'database')}, ${sizeText(c.mb)}. Click to list the applications.`)}"><b>${nf(c.servers)}</b><span>${nf(c.dbs)} db</span></td>`;
    }).join('');
    return `<tr><th class="rel-h" style="background:#fff;color:var(--ink);position:static;vertical-align:middle">${relChip(rel)}</th><td>${statusChip(rel)}</td>${cells}
      <td class="tot cell" data-rel="${esc(rel)}" data-env="ALL" title="Click to list the applications on ${esc(rel)}">${nf(list.length)}<span style="font-weight:400;color:var(--ink3)"> / ${nf(sum(list.map(s => s.dbs.length)))} db</span></td>
      <td class="small">${builds.slice(0, 4).map(([v, n]) => `${esc(v)} <span class="nul">(${nf(n)})</span>`).join('<br>')}${builds.length > 4 ? `<br><span class="nul">+${builds.length - 4} more builds</span>` : ''}</td></tr>`;
  }).join('');
  const foot = ESTATE_ENVS.map(e => { const l = SERVERS.filter(s => s.env === e); return `<td>${nf(l.length)}<div class="note">${nf(sum(l.map(s => s.dbs.length)))} db</div></td>`; }).join('');
  return `<div class="tbl-wrap"><table class="table matrix">
    <thead><tr><th>SQL Server release</th><th>Microsoft support</th>${ESTATE_ENVS.map(e => `<th class="ec-h" style="text-align:center">${envTag(e)}<small>servers / dbs</small></th>`).join('')}<th style="text-align:center">All environments</th><th>Builds seen (servers)</th></tr></thead>
    <tbody>${rows}</tbody>
    <tfoot><tr><td style="text-align:left">All releases</td><td></td>${foot}<td>${nf(SERVERS.length)}<div class="note">${nf(USER_DBS.length)} db</div></td><td></td></tr></tfoot></table></div>
    <p class="note">Shade shows the number of servers in a cell: <span class="ramp">${RAMP.map(c => `<i style="background:${c}"></i>`).join('')}</span> fewer to more. Click any cell to list the applications on it. Support status is Microsoft's fixed lifecycle measured on ${esc(META.generated.slice(0, 10))}; "Support ending" means extended support ends within ${META.warnMonths} months. Paid Extended Security Updates are not taken into account.</p>`;
}

function crossHtml(title, keyOf) {
  const m = new Map();
  SERVERS.forEach(s => {
    const k = has(keyOf(s)) ? keyOf(s) : 'Not recorded';
    if (!m.has(k)) m.set(k, new Map());
    const row = m.get(k);
    row.set(s.env, (row.get(s.env) || 0) + 1);
  });
  let rows = [...m.entries()].map(([k, row]) => ({ k, row, total: sum([...row.values()]) })).sort((x, y) => y.total - x.total);
  if (rows.length > 12) {
    const rest = rows.slice(11), other = new Map();
    rest.forEach(r => r.row.forEach((n, e) => other.set(e, (other.get(e) || 0) + n)));
    rows = rows.slice(0, 11).concat([{ k: `${rest.length} others`, row: other, total: sum(rest.map(r => r.total)) }]);
  }
  return `<div><h3 class="sub">${esc(title)}</h3><div class="tbl-wrap"><table class="table">
    <thead><tr><th>${esc(title)}</th>${ESTATE_ENVS.map(e => `<th class="num">${envTag(e)}</th>`).join('')}<th class="num">Servers</th></tr></thead>
    <tbody>${rows.map(r => `<tr><td>${r.k === 'Not recorded' ? '<span class="nul">Not recorded</span>' : esc(r.k)}</td>${ESTATE_ENVS.map(e => `<td class="num mono">${r.row.get(e) ? nf(r.row.get(e)) : '<span class="nul">&middot;</span>'}</td>`).join('')}<td class="num mono"><b>${nf(r.total)}</b></td></tr>`).join('')}</tbody></table></div></div>`;
}

function configHtml() {
  const all = USER_DBS, prod = all.filter(d => d.srv.env === 'PROD');
  const count = (list, f) => list.filter(f).length;
  const breakdown = (list, f) => [...list.reduce((m, d) => { const k = f(d) || 'not recorded'; return m.set(k, (m.get(k) || 0) + 1); }, new Map()).entries()]
    .sort((x, y) => y[1] - x[1]).slice(0, 4).map(([k, n]) => `${esc(k)} ${nf(n)}`).join(' &middot; ');
  const notOnline = d => has(d.status) && d.status !== 'ONLINE';
  const simple = d => d.recovery === 'SIMPLE';
  const pvBad = d => has(d.pv) && d.pv !== 'CHECKSUM';
  const compatLow = d => { const n = nativeOf(d.srv); return d.compat != null && n != null && d.compat < n; };
  const inAg = d => has(d.ag);
  const noSize = d => d.mb == null;
  const card = (label, n, prodN, sev, detail, sevWord) => `<div class="cfg-card" style="--c-c:var(--${sev})">
      <div class="h"><span>${esc(label)}</span><span class="flag ${sev === 'ink3' ? 'info' : sev}">${esc(sevWord)}</span></div>
      <div class="n">${nf(n)}<small>${prodN}</small></div><div class="d">${detail}</div></div>`;
  const agProd = count(prod, inAg);
  const cards = [
    card('Databases not ONLINE', count(all, notOnline), `${nf(count(prod, notOnline))} in PROD`, count(all, notOnline) ? 'crit' : 'good',
      count(all, notOnline) ? breakdown(all.filter(notOnline), d => d.status) : 'Every database with a status is ONLINE.', count(all, notOnline) ? 'Check' : 'OK'),
    card('SIMPLE recovery model', count(all, simple), `${nf(count(prod, simple))} in PROD`, count(prod, simple) ? 'warn' : 'good',
      `No point-in-time restore for these. ${breakdown(all, d => d.recovery)}`, count(prod, simple) ? 'PROD affected' : 'OK'),
    card('Page verify not CHECKSUM', count(all, pvBad), `${nf(count(prod, pvBad))} in PROD`, count(all, pvBad) ? 'warn' : 'good',
      count(all, pvBad) ? breakdown(all.filter(pvBad), d => d.pv) : 'Every database with the setting recorded uses CHECKSUM.', count(all, pvBad) ? 'Review' : 'OK'),
    card('Compatibility below the engine', count(all, compatLow), `${nf(count(prod, compatLow))} in PROD`, count(all, compatLow) ? 'warn' : 'good',
      'Running at a lower compatibility level than its SQL Server release gives a new database. Worth knowing before an upgrade.', count(all, compatLow) ? 'Review' : 'OK'),
    card('PROD databases in an availability group', agProd, `of ${nf(prod.length)} in PROD`, 'info',
      prod.length ? `${(agProd / prod.length * 100).toFixed(0)}% of PROD databases carry an AG name. ${nf(uniq(prod.map(d => d.ag).filter(has)).length)} distinct AGs.` : 'No PROD databases.', 'Info'),
    card('Collations in use', uniq(all.map(d => d.coll).filter(has)).length, 'distinct', 'info', breakdown(all, d => d.coll), 'Info'),
    card('Data classification', uniq(all.map(d => d.cls).filter(has)).length, 'distinct values', 'info', breakdown(all, d => d.cls), 'Info'),
    card('No size recorded', count(all, noSize), `of ${nf(all.length)}`, count(all, noSize) ? 'warn' : 'good', 'Left out of every data total on this page.', count(all, noSize) ? 'Gap' : 'OK')
  ];
  return `<div class="cfg">${cards.join('')}</div>`;
}

function renderEstate() {
  const prodServers = SERVERS.filter(s => s.env === 'PROD');
  const oos = SERVERS.filter(s => relStatus(s.rel) === 'oos');
  const ending = SERVERS.filter(s => relStatus(s.rel) === 'ending');
  const unknownVer = SERVERS.filter(s => relStatus(s.rel) === 'unknown');
  const withCpu = SERVERS.filter(s => s.cpu != null), withMem = SERVERS.filter(s => s.mem != null);
  const oosCores = sum(oos.filter(s => s.cpu != null).map(s => s.cpu));
  const memGb = sum(withMem.map(s => s.mem));
  const pct = (n, of) => of ? `${(n / of * 100).toFixed(0)}%` : '0%';
  document.getElementById('estate').innerHTML = `
    <div class="sec-head"><div><h2 class="sec">SQL Server estate</h2>
      <p class="note">Which SQL Server releases run where, how far each is through Microsoft support, and what the servers and databases look like.</p></div></div>
    <div class="kpis">${[
      ['SQL servers', nf(SERVERS.length), `${nf(prodServers.length)} in PROD &middot; ${nf(SERVERS.filter(s => !s.inSrv).length)} only in the database sheet`, 'var(--brand)'],
      ['Out of support', nf(oos.length), `${pct(oos.length, SERVERS.length)} of servers &middot; ${nf(oos.filter(s => s.env === 'PROD').length)} in PROD &middot; ${nf(oosCores)} cores`, 'var(--crit)'],
      ['Support ending soon', nf(ending.length), `extended support ends within ${META.warnMonths} months &middot; ${nf(ending.filter(s => s.env === 'PROD').length)} in PROD`, 'var(--serious)'],
      ['Version unknown', nf(unknownVer.length), 'no version recorded, or not recognised', 'var(--ink3)'],
      ['Cores', nf(sum(withCpu.map(s => s.cpu))), `logical CPUs, ${nf(withCpu.length)} of ${nf(SERVERS.length)} servers recorded`, 'var(--info)'],
      ['Memory', memGb >= 1024 ? `${(memGb / 1024).toFixed(1)} TB` : `${nf(Math.round(memGb))} GB`, `${nf(withMem.length)} of ${nf(SERVERS.length)} servers recorded`, 'var(--info)'],
      ['Data', esc(sizeText(mbOf(USER_DBS))), `${nf(USER_DBS.length)} user databases`, 'var(--accent)']
    ].map(kpiHtml).join('')}</div>
    <h3 class="sub">SQL Server release by environment</h3>
    ${matrixHtml()}
    <div class="grid2">${crossHtml('Platform', s => s.platform)}${crossHtml('Location', s => s.location)}</div>
    <h3 class="sub">Database configuration (user databases)</h3>
    ${configHtml()}`;
  document.querySelectorAll('#estate [data-rel]').forEach(c => c.onclick = () => {
    state.cell = { rel: c.dataset.rel, env: c.dataset.env };
    renderAppTable();
    document.getElementById('apps').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

/* ---------------------------------------------------------------------------
   Reconciliation: every total on the page traces back to rows in the workbook
--------------------------------------------------------------------------- */
function renderRecon() {
  const R = META.recon;
  const notInSheet = SERVERS.filter(s => !s.inSrv);
  const noDbs = SERVERS.filter(s => s.inSrv && !s.dbRows);
  const verDiff = SERVERS.filter(s => has(s.verNote));
  const envDiff = USER_DBS.filter(d => d.env !== 'UNKNOWN' && d.srv.env !== 'UNKNOWN' && d.env !== d.srv.env);
  const onlySrv = APPS.filter(a => !a.unknown && a.inSrv && !a.inDb);
  const onlyDb = APPS.filter(a => !a.unknown && !a.inSrv && a.inDb);
  const line = (label, value) => `<div class="rc-line"><span>${label}</span><b>${value}</b></div>`;
  const listOf = (label, items) => items.length
    ? `<details class="rc"><summary>${label} (${nf(items.length)})</summary><div class="list">${items.slice(0, 300).map(esc).join('<br>')}${items.length > 300 ? `<br>+${nf(items.length - 300)} more` : ''}</div></details>` : '';

  document.getElementById('recon').innerHTML = `
    <div class="sec-head"><div><h2 class="sec">Reconciliation</h2>
      <p class="note">How the rows in the workbook became the numbers on this page, and what did not line up between the two sheets.</p></div></div>
    <div class="grid2">
      <div>
        <h3 class="sub">Sources</h3>
        <div class="tbl-wrap"><table class="table"><thead><tr><th>Sheet</th><th>Read as</th><th class="num">Rows</th><th>Columns not found</th></tr></thead>
          <tbody>${R.sources.map(s => `<tr><td>${esc(s.source)}</td><td>${s.role === 'Ignored' ? '<span class="nul">Ignored</span>' : esc(s.role)}</td><td class="num mono">${nf(s.rows)}</td><td class="small">${nz(s.missing)}</td></tr>`).join('')}</tbody></table></div>
        <h3 class="sub">Server sheet</h3>
        ${line('Rows read', nf(R.serverRows))}
        ${line('Rows without a server name, skipped', nf(R.serverRowsSkipped))}
        ${line('Distinct SQL servers in the server sheet', nf(SERVERS.filter(s => s.inSrv).length))}
        ${line('Servers listed with more than one application', nf(SERVERS.filter(s => s.inSrv && s.sheetApps.size > 1).length))}
        <h3 class="sub">Database sheet</h3>
        ${line('Rows read', nf(R.dbRows))}
        ${line('Rows without a database name, skipped', nf(R.dbRowsSkipped))}
        ${line('Repeats of a database already listed (latest inventory date kept)', nf(R.dbRepeats))}
        ${line('Distinct databases', nf(R.dbKept))}
        ${line('&nbsp;&nbsp;user databases', nf(R.dbUser))}
        ${line('&nbsp;&nbsp;system databases (master, model, msdb, tempdb)', nf(R.dbSystem))}
        ${line('Inventory dates found', esc((R.snapshots || []).map(([d, n]) => `${d} (${nf(n)} rows)`).join(', ') || 'none'))}
      </div>
      <div>
        <h3 class="sub">Between the two sheets</h3>
        ${line('Servers only in the database sheet', nf(notInSheet.length))}
        ${line('Servers in the server sheet with no database listed', nf(noDbs.length))}
        ${line('Servers whose SQL version disagrees between the sheets', nf(verDiff.length))}
        ${line('Databases whose environment differs from their server', nf(envDiff.length))}
        ${line('Applications only in the server sheet', nf(onlySrv.length))}
        ${line('Applications only in the database sheet', nf(onlyDb.length))}
        ${line('Database-sheet server names matched to the server sheet by host name', nf((R.hostMatched || []).length))}
        ${listOf('Servers only in the database sheet', notInSheet.map(s => `${s.name} \u2013 ${envLabel(s.env)}, ${plural(s.dbs.length, 'database')}`))}
        ${listOf('Servers with no database listed', noDbs.map(s => `${s.name} \u2013 ${envLabel(s.env)}`))}
        ${listOf('SQL version disagreements', verDiff.map(s => `${s.name}: ${s.verNote}`))}
        ${listOf('Environment differences', envDiff.map(d => `${d.name} on ${d.srv.name}: database sheet ${d.env}, server ${d.srv.env}`))}
        ${listOf('Applications only in the server sheet', onlySrv.map(a => a.name))}
        ${listOf('Applications only in the database sheet', onlyDb.map(a => a.name))}
        ${listOf('Matched by host name', R.hostMatched || [])}
      </div>
    </div>`;
  document.getElementById('foot').textContent = `Source: ${META.source} \u00b7 Generated ${META.generated} \u00b7 SQL Application Map`;
}

/* ---------------------------------------------------------------------------
   Start
--------------------------------------------------------------------------- */
document.getElementById('q').addEventListener('input', e => { state.q = e.target.value; renderAppTable(); });
document.getElementById('critFilter').addEventListener('change', e => { state.crit = e.target.value; renderAppTable(); });
document.getElementById('showFilter').addEventListener('change', e => { state.show = e.target.value; renderAppTable(); });

initFilters();
renderHeader();
renderUnknown();
const firstNamed = APPS.filter(a => !a.unknown && a.servers.size).sort((x, y) => byName(x.name, y.name))[0];
state.selected = firstNamed ? firstNamed.id : UNK.id;
renderAppTable();
renderDetail();
renderEstate();
renderRecon();
</script>
</body>
</html>
'@

$html = ($htmlStyle + $htmlBody + $htmlScriptModel + $htmlScriptApps + $htmlScriptDetail + $htmlScriptEstate).
    Replace('__T_BRAND__', $theme.Brand).
    Replace('__T_BRAND_DARK__', $theme.BrandDark).
    Replace('__T_BRAND_SOFT__', $theme.BrandSoft).
    Replace('__T_ACCENT__', $theme.Accent).
    Replace('__T_ACCENT_DARK__', $theme.AccentDark).
    Replace('__T_ACCENT_SOFT__', $theme.AccentSoft).
    Replace('__PAYLOAD__', $json)

$outputDirectory = Split-Path -Parent $OutputHtml
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
[IO.File]::WriteAllText($OutputHtml, $html, [Text.UTF8Encoding]::new($false))

# =========================================================================================
# REGION  Console summary
# =========================================================================================

$unknownApp = $apps[$UnknownKey]
$unknownDatabases = @($dbItems | Where-Object { -not $_.System -and $_.AppId -eq $unknownApp.id }).Count
$namedServerKeys = @{}
foreach ($link in $links.Values) { if ($link.app -ne $UnknownKey) { $namedServerKeys[$link.server] = $true } }
foreach ($item in $dbItems) { if (-not $item.System -and $item.AppId -ne $unknownApp.id) { $namedServerKeys[$item.Server.key] = $true } }
$unknownServers = @($servers.Values | Where-Object { -not $namedServerKeys.ContainsKey($_.key) }).Count
$outOfSupport = @($servers.Values | Where-Object { (Get-ReleaseStatus $releaseByLabel[$_.release])[0] -eq 'oos' }).Count

Write-Host ''
Write-Host 'Created: ' -NoNewline -ForegroundColor Green; Write-Host $OutputHtml
Write-Host ('  Applications {0}   SQL servers {1}   Databases {2} user + {3} system' -f ($apps.Count - 1), $servers.Count, $dbUser, $dbSystem) -ForegroundColor Cyan
Write-Host ('  Database rows {0} -> {1} repeats removed (latest inventory date kept), {2} skipped with no database name' -f $databaseRecords.Count, $dbRepeats, $dbRowsSkipped) -ForegroundColor Cyan
$environmentSpread = ($servers.Values | Group-Object { $_.env } | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '   '
Write-Host "  Servers by environment: $environmentSpread" -ForegroundColor Cyan
$colour = $(if ($unknownDatabases -or $unknownServers) { 'Red' } else { 'Green' })
Write-Host ('  UNKNOWN application: {0} databases, {1} servers with no application' -f $unknownDatabases, $unknownServers) -ForegroundColor $colour
if ($outOfSupport) { Write-Host ('  {0} servers on a SQL Server release past Microsoft extended support' -f $outOfSupport) -ForegroundColor Yellow }
if ($Simulate) { Write-Host '  SIMULATED: every name and figure is made up. Do not send it as a finding.' -ForegroundColor Yellow }

if (-not $NoLaunch) { Start-Process $OutputHtml }
