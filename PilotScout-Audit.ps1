<#
.SYNOPSIS
    PilotScout - Microsoft 365 Copilot seat waste audit. Finds Copilot licenses that are
    unassigned, never used, dormant or stuck on disabled accounts, and quantifies how much
    money you can recover.

.DESCRIPTION
    Runs entirely inside YOUR tenant using the Microsoft Graph PowerShell SDK (read-only).
    No data ever leaves your environment - PilotScout has no servers and collects nothing.

    Produces:
      * CSV exports for each finding (for remediation / renewal negotiation)
      * A console summary

    Findings:
      * Unassigned purchased Copilot seats  (paying for seats nobody has)
      * Never-used seats                    (assigned, zero Copilot activity in the period)
      * Dormant seats                       (no Copilot activity in N days)
      * Disabled-but-licensed users         (blocked accounts still holding Copilot)

    Unlike sign-in based audits, PilotScout reads the dedicated Microsoft 365 Copilot usage
    report - so it finds ACTIVE users whose Copilot seat is doing nothing. No Entra ID P1 needed.

.PARAMETER InactiveDays
    A seat with no Copilot activity in this many days is flagged dormant. Default 60.

.PARAMETER UsagePeriod
    Copilot usage report window: D7 | D30 | D90 | D180 | ALL. Default D180.
    Never-used = no activity at all within this window.

.PARAMETER OutputPath
    Folder for the report + CSVs. Default: .\PilotScout-Report-<timestamp>

.PARAMETER PriceTablePath
    CSV mapping skuPartNumber -> friendly name + monthly price. It only prices SKUs: which SKUs are
    Copilot seats is read from their service plans (M365_COPILOT_APPS), never from this list.
    Default: .\config\PriceTable.csv

.PARAMETER Lite
    Free mode: exports the never-used/dormant CSV + a summary only (no $ totals, no HTML report).

.PARAMETER MockDataPath
    Internal/testing: load tenant data from a JSON file instead of calling Microsoft Graph.

.EXAMPLE
    .\PilotScout-Audit.ps1 -InactiveDays 60

.NOTES
    PilotScout  |  pilotscout.dev  |  read-only, runs in your tenant.
    Independent tool - not affiliated with Microsoft. "Microsoft 365 Copilot" is a Microsoft trademark.
#>

[CmdletBinding()]
param(
    [int]    $InactiveDays   = 60,
    [ValidateSet('D7','D30','D90','D180','ALL')]
    [string] $UsagePeriod    = 'D180',
    [string] $OutputPath,
    [string] $PriceTablePath = (Join-Path (Join-Path $PSScriptRoot 'config') 'PriceTable.csv'),
    [switch] $Lite,
    [string] $MockDataPath
)

$ErrorActionPreference = 'Stop'
$script:Version = '1.0.3'
$script:Edition = 'Lite'   # build edition (set per package): Lite | Solo | Pro
$nowUtc = (Get-Date).ToUniversalTime()

# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Write-Step { param($Msg) Write-Host "[PilotScout] $Msg" -ForegroundColor Cyan }
function Write-Warn { param($Msg) Write-Host "[PilotScout] $Msg" -ForegroundColor Yellow }

# A subscribed SKU is a Copilot seat when one of its service plans is M365_COPILOT_APPS (Microsoft 365
# Copilot in Productivity Apps). Microsoft sells the same seat under several part numbers
# (Microsoft_365_Copilot, M365_Copilot, Microsoft_365_Copilot_EDU, Copilot Business) and inside bundles
# (Microsoft 365 E7), so names are never matched. Bing_Chat_Enterprise is Copilot Chat, free with
# E3/E5/Business plans: a SKU that has only that plan is not a paid seat and is not audited.
$script:CopilotAppsPlanId = 'a62f8878-de10-42f3-b68f-6149a25ceb97'
$script:AssumedSeatPrice  = 30   # USD/month for a Copilot SKU with no PriceTable.csv price (list price of the add-on)

function Test-CopilotSku {
    param($Sku)
    foreach ($p in @($Sku.servicePlans)) {
        if ($p -and [string]$p.servicePlanId -eq $script:CopilotAppsPlanId) { return $true }
    }
    return $false
}

function Get-PriceTable {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Price table not found: $Path" }
    $rows = Import-Csv -Path $Path | Where-Object { $_.SkuPartNumber -and -not $_.SkuPartNumber.StartsWith('#') }
    $map = @{}
    foreach ($r in $rows) {
        $price = $null
        if ($r.MonthlyPriceUSD -and ($r.MonthlyPriceUSD -as [double]) -ne $null) { $price = [double]$r.MonthlyPriceUSD }
        $map[$r.SkuPartNumber] = [pscustomobject]@{
            PartNumber   = $r.SkuPartNumber
            FriendlyName = if ($r.FriendlyName) { $r.FriendlyName } else { $r.SkuPartNumber }
            Price        = $price
        }
    }
    return $map
}

# Copilot per-app activity fields in the usage report (v1)
$script:AppFields = [ordered]@{
    'Copilot Chat' = 'copilotChatLastActivityDate'
    'Teams'        = 'microsoftTeamsCopilotLastActivityDate'
    'Word'         = 'wordCopilotLastActivityDate'
    'Excel'        = 'excelCopilotLastActivityDate'
    'PowerPoint'   = 'powerPointCopilotLastActivityDate'
    'Outlook'      = 'outlookCopilotLastActivityDate'
    'OneNote'      = 'oneNoteCopilotLastActivityDate'
    'Loop'         = 'loopCopilotLastActivityDate'
}

function Get-FieldDate {
    param($Row, [string]$Field)
    if ($null -eq $Row) { return $null }
    $p = $Row.PSObject.Properties[$Field]
    if (-not $p -or -not $p.Value) { return $null }
    try { return [datetime]$p.Value } catch { return $null }
}

# ----------------------------------------------------------------------------------------------
# Data acquisition (Graph OR mock)
# ----------------------------------------------------------------------------------------------
function Get-TenantData {
    param([string]$MockPath, [string]$Period)

    if ($MockPath) {
        Write-Step "Loading MOCK data from $MockPath"
        $raw = Get-Content -Path $MockPath -Raw | ConvertFrom-Json
        # Mock dates are absolute, so age/inactivity math runs against the mock's own asOf instead of
        # the clock; otherwise the demo output (and the published sample figures) drift day by day.
        if ($raw.asOf) { $script:nowUtc = ([datetime]$raw.asOf).ToUniversalTime() }
        return $raw
    }

    Write-Step "Connecting to Microsoft Graph (read-only)..."
    $scopes = @('Organization.Read.All','User.Read.All','Directory.Read.All','Reports.Read.All')
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
    } catch {
        throw "Microsoft Graph PowerShell SDK not installed. Run:  Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Connect-MgGraph -Scopes $scopes -NoWelcome

    Write-Step "Reading organization + subscribed SKUs..."
    $org  = Get-MgOrganization
    $skus = Get-MgSubscribedSku -All

    Write-Step "Reading users (license assignments)..."
    $users = Get-MgUser -All -Property 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses,createdDateTime'

    Write-Step "Reading Microsoft 365 Copilot usage report (period=$Period)..."
    $usage = @()
    $usageRetrievable = $true
    $refreshDate = $null
    try {
        $uri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period')"
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
            $vals = @($resp.value)
            foreach ($v in $vals) { $usage += [pscustomobject]$v }
            $uri = $resp.'@odata.nextLink'
        }
        if ($usage.Count -gt 0 -and $usage[0].reportRefreshDate) { $refreshDate = [string]$usage[0].reportRefreshDate }
    } catch {
        Write-Warn "Copilot usage report unavailable (needs Reports.Read.All + a reports-capable admin role)."
        Write-Warn "Continuing with license-position analysis only (unassigned + disabled seats)."
        $usageRetrievable = $false
    }

    # Concealed-names detection: report privacy setting replaces UPNs with 32-char hex hashes
    $concealed = $false
    if ($usageRetrievable -and $usage.Count -gt 0) {
        $sample = @($usage | Select-Object -First 20 | ForEach-Object { [string]$_.userPrincipalName })
        $hashLike = @($sample | Where-Object { $_ -match '^[0-9A-Fa-f]{32}$' })
        if ($hashLike.Count -gt ($sample.Count / 2)) { $concealed = $true }
    }

    return [pscustomobject]@{
        organization      = [pscustomobject]@{ displayName = $org.DisplayName; id = $org.Id }
        usageRetrievable  = $usageRetrievable
        namesConcealed    = $concealed
        usagePeriod       = $Period
        reportRefreshDate = $refreshDate
        subscribedSkus = $skus | ForEach-Object {
            [pscustomobject]@{
                skuId         = $_.SkuId
                skuPartNumber = $_.SkuPartNumber
                prepaidUnits  = [pscustomobject]@{ enabled = $_.PrepaidUnits.Enabled }
                consumedUnits = $_.ConsumedUnits
                servicePlans  = @($_.ServicePlans | ForEach-Object { [pscustomobject]@{ servicePlanId = $_.ServicePlanId; servicePlanName = $_.ServicePlanName } })
            }
        }
        users = $users | ForEach-Object {
            [pscustomobject]@{
                id                = $_.Id
                displayName       = $_.DisplayName
                userPrincipalName = $_.UserPrincipalName
                accountEnabled    = $_.AccountEnabled
                createdDateTime   = $_.CreatedDateTime
                assignedLicenses  = @($_.AssignedLicenses | ForEach-Object { [pscustomobject]@{ skuId = $_.SkuId } })
            }
        }
        copilotUsage = $usage
    }
}

# ----------------------------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------------------------
function Invoke-Analysis {
    param($Data, $PriceMap, [int]$InactiveDays)

    # Copilot seat SKUs: every subscribed SKU carrying the M365_COPILOT_APPS service plan (an add-on, a
    # bundle, or both - each is audited and the totals add up). PriceTable.csv only prices them.
    $skuMeta = @(); $skuById = @{}
    foreach ($s in $Data.subscribedSkus) {
        if (-not (Test-CopilotSku $s)) { continue }
        $meta = $PriceMap[[string]$s.skuPartNumber]
        $friendly = if ($meta) { $meta.FriendlyName } else { $s.skuPartNumber }
        # No row, or a row with a blank price: price at the add-on list rate and flag it. A price of 0 leaves it out of the totals.
        $priceAssumed = (-not $meta) -or ($null -eq $meta.Price)
        $price = if ($priceAssumed) { $script:AssumedSeatPrice } else { $meta.Price }
        $purchased = [int]$s.prepaidUnits.enabled
        $consumed  = [int]$s.consumedUnits
        $unassigned = [math]::Max(0, $purchased - $consumed)
        $info = [pscustomobject]@{
            SkuId=$s.skuId; PartNumber=$s.skuPartNumber; Friendly=$friendly; Price=$price; PriceAssumed=$priceAssumed
            Purchased=$purchased; Consumed=$consumed; Unassigned=$unassigned
            UnassignedWasteMonthly = if ($price) { $unassigned * $price } else { $null }
        }
        $skuById[([string]$s.skuId).ToLower()] = $info
        $skuMeta += $info
    }

    if ($skuMeta.Count -eq 0) {
        $held = (@($Data.subscribedSkus) | ForEach-Object { $_.skuPartNumber } | Sort-Object) -join ', '
        throw "No Microsoft 365 Copilot seat SKU found in this tenant: no subscribed SKU includes the M365_COPILOT_APPS service plan. Copilot Chat on its own (Bing_Chat_Enterprise) is free and has no seats to audit. SKUs seen: $held"
    }

    $cutoff = $nowUtc.AddDays(-$InactiveDays)
    $usageOK   = if ($null -ne $Data.usageRetrievable) { [bool]$Data.usageRetrievable } else { $true }
    $concealed = if ($null -ne $Data.namesConcealed)   { [bool]$Data.namesConcealed }   else { $false }

    # Belt-and-braces concealment detection on the data itself (not just at acquisition time):
    # the reports privacy setting replaces UPNs with 32-char hex hashes, which would otherwise
    # silently fail the directory join and mislabel every seat as never-used.
    if ($usageOK -and -not $concealed) {
        $sample = @(@($Data.copilotUsage) | Select-Object -First 20 | ForEach-Object { [string]$_.userPrincipalName })
        if ($sample.Count -gt 0) {
            $hashLike = @($sample | Where-Object { $_ -match '^[0-9A-Fa-f]{32}$' })
            if ($hashLike.Count -gt ($sample.Count / 2)) { $concealed = $true }
        }
    }

    # usage rows indexed by UPN (lowercase) for the directory join
    $usageByUpn = @{}
    foreach ($r in @($Data.copilotUsage)) {
        $k = ([string]$r.userPrincipalName).ToLower()
        if ($k) { $usageByUpn[$k] = $r }
    }

    # default per-seat price for count-only (concealed) mode = the priced Copilot SKU with most seats
    $primarySku = $skuMeta | Where-Object { $_.Price } | Sort-Object Purchased -Descending | Select-Object -First 1
    $defaultSeatPrice = if ($primarySku) { [double]$primarySku.Price } else { 0 }

    $disabled=@(); $never=@(); $dormant=@()
    $licensedCount = 0; $licensedEnabledCount = 0

    foreach ($u in $Data.users) {
        # matched by the skuId(s) found above, never by name
        $lic = @($u.assignedLicenses | ForEach-Object { $skuById[([string]$_.skuId).ToLower()] } | Where-Object { $_ })
        if ($lic.Count -eq 0) { continue }   # not a Copilot holder
        $licensedCount++

        $userMonthly = ($lic | Where-Object { $_.Price } | Measure-Object -Property Price -Sum).Sum
        if (-not $userMonthly) { $userMonthly = 0 }
        $licNames = ($lic | ForEach-Object { $_.Friendly }) -join '; '

        $row = $null
        if (-not $concealed) { $row = $usageByUpn[([string]$u.userPrincipalName).ToLower()] }

        # last Copilot activity = max of the overall + per-app dates
        $lastAny = Get-FieldDate $row 'lastActivityDate'
        $appsUsed = @()
        foreach ($an in $script:AppFields.Keys) {
            $d = Get-FieldDate $row $script:AppFields[$an]
            if ($d) {
                $appsUsed += $an
                if (-not $lastAny -or $d -gt $lastAny) { $lastAny = $d }
            }
        }

        $created = $null
        if ($u.createdDateTime) { try { $created = [datetime]$u.createdDateTime } catch {} }

        $rec = [pscustomobject]@{
            DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName
            Licenses=$licNames
            LastCopilotUse = if ($lastAny) { $lastAny.ToString('yyyy-MM-dd') } else { 'never' }
            AppsUsed = if ($appsUsed.Count -gt 0) { $appsUsed -join ', ' } else { '-' }
            MonthlyCost=[math]::Round($userMonthly,2)
        }

        if (-not $u.accountEnabled) {
            $disabled += $rec
            continue
        }
        $licensedEnabledCount++

        if ($concealed -or -not $usageOK) { continue }   # per-user activity buckets need joinable usage data

        if (-not $lastAny) {
            # never used - only flag if the account is old enough for that to be meaningful
            if (-not $created -or $created -lt $cutoff) { $never += $rec }
        }
        elseif ($lastAny -lt $cutoff) {
            $dormant += $rec
        }
    }

    # Concealed / no-join mode: derive never+dormant COUNTS from the usage rows alone
    $countOnly = $false
    if ($usageOK -and $concealed) {
        $countOnly = $true
        $rowsWithAny = 0; $dormantRows = 0
        foreach ($r in @($Data.copilotUsage)) {
            $lastAny = Get-FieldDate $r 'lastActivityDate'
            foreach ($an in $script:AppFields.Keys) {
                $d = Get-FieldDate $r $script:AppFields[$an]
                if ($d -and (-not $lastAny -or $d -gt $lastAny)) { $lastAny = $d }
            }
            if ($lastAny) {
                $rowsWithAny++
                if ($lastAny -lt $cutoff) { $dormantRows++ }
            }
        }
        $neverCount   = [math]::Max(0, $licensedEnabledCount - $rowsWithAny)
        $dormantCount = $dormantRows
        # synthesize count-only buckets (no names available)
        $never   = @(1..$neverCount   | ForEach-Object { [pscustomobject]@{ DisplayName='(name concealed)'; UserPrincipalName='-'; Licenses='-'; LastCopilotUse='never'; AppsUsed='-'; MonthlyCost=$defaultSeatPrice } })
        $dormant = @(1..$dormantCount | ForEach-Object { [pscustomobject]@{ DisplayName='(name concealed)'; UserPrincipalName='-'; Licenses='-'; LastCopilotUse='(concealed)'; AppsUsed='-'; MonthlyCost=$defaultSeatPrice } })
        if ($neverCount -eq 0)   { $never = @() }
        if ($dormantCount -eq 0) { $dormant = @() }
    }

    $result = [ordered]@{
        Org=$Data.organization
        Skus=$skuMeta
        LicensedCount=$licensedCount
        LicensedEnabledCount=$licensedEnabledCount
        UnassignedSeats=(($skuMeta | Measure-Object -Property Unassigned -Sum).Sum)
        Disabled=$disabled; Never=$never; Dormant=$dormant
        UsageAvailable=$usageOK
        NamesConcealed=$concealed
        CountOnly=$countOnly
        UsagePeriod=[string]$Data.usagePeriod
        ReportRefreshDate=[string]$Data.reportRefreshDate
        DefaultSeatPrice=$defaultSeatPrice
    }
    return [pscustomobject]$result
}

# ----------------------------------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "  PilotScout v$script:Version - Microsoft 365 Copilot seat waste audit" -ForegroundColor Green
Write-Host "  Read-only. Runs in your tenant. Nothing leaves your environment." -ForegroundColor DarkGray
Write-Host ""

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("PilotScout-Report-" + $nowUtc.ToString('yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$priceMap = Get-PriceTable -Path $PriceTablePath
$data     = Get-TenantData -MockPath $MockDataPath -Period $UsagePeriod
$R        = Invoke-Analysis -Data $data -PriceMap $priceMap -InactiveDays $InactiveDays

# CSV exports (always)
$R.Never    | Export-Csv (Join-Path $OutputPath 'never-used-seats.csv') -NoTypeInformation -Encoding utf8
$R.Dormant  | Export-Csv (Join-Path $OutputPath 'dormant-seats.csv') -NoTypeInformation -Encoding utf8
$R.Disabled | Export-Csv (Join-Path $OutputPath 'disabled-but-licensed.csv') -NoTypeInformation -Encoding utf8

if ($Lite -or $script:Edition -eq 'Lite') {
    Write-Host ""
    Write-Step "LITE mode - free edition"
    Write-Host ("  Copilot seats purchased : {0}" -f (($R.Skus | Measure-Object -Property Purchased -Sum).Sum))
    Write-Host ("  Assigned                : {0}" -f $R.LicensedCount)
    Write-Host ("  Unassigned              : {0}" -f $R.UnassignedSeats)
    Write-Host ("  Never used              : {0}" -f $R.Never.Count)
    Write-Host ("  Dormant > $InactiveDays days       : {0}" -f $R.Dormant.Count)
    Write-Host ("  Disabled but licensed   : {0}" -f $R.Disabled.Count)
    Write-Host ""
    Write-Host "  Solo/Pro add the HTML report — pilotscout.dev" -ForegroundColor Yellow
    Write-Host ""
    Write-Step "CSV saved to: $OutputPath"
    return
}

