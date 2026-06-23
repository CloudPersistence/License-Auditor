<#
.SYNOPSIS
    CloudPersistence License Auditor — Microsoft 365 license waste audit (cloudpersistence.com).

.DESCRIPTION
    Connects to Microsoft Graph using DELEGATED interactive authentication
    (no app registration is created, no secret is stored) and analyses the
    tenant for three high-confidence, low-noise sources of license waste:

        1. Shelfware        : paid-but-unassigned licenses (from subscribedSkus)
        2. Disabled users   : licenses still attached to disabled accounts
        3. Inactive users   : licensed users with no sign-in for N+ days

    Prices are read from (or built into) an external JSON file. On first run
    the script asks only for the SKUs actually present in the tenant, then
    saves the answers so later runs are prompt-free. No tenant data ever
    leaves the operator's computer.

.NOTES
    - Runs READ-ONLY. The recommended role is "Global Reader".
    - Prices you enter are YOUR figures. The hints shown are rough USD list
      prices for convenience only and may be outdated.
    - "Inactive user" detection requires Entra ID P1/P2 in the target tenant
      (for signInActivity). Without P1 the report flags this section as N/A.
    - Required modules: Microsoft.Graph.Authentication,
      Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement.

.EXAMPLE
    .\Invoke-CPLicenseAuditor.ps1
    .\Invoke-CPLicenseAuditor.ps1 -InactiveDays 60 -OutputFolder "C:\Reports"
    .\Invoke-CPLicenseAuditor.ps1 -UseListPrices       # no prompts, use hints
    .\Invoke-CPLicenseAuditor.ps1 -NonInteractive      # no prompts, file only
#>

[CmdletBinding()]
param(
    # Number of days without a sign-in before a licensed user is flagged inactive.
    [int]$InactiveDays = 90,

    # Folder where the HTML report is written. Defaults to the current directory.
    [string]$OutputFolder = (Get-Location).Path,

    # JSON file that stores currency + per-SKU prices between runs.
    [string]$PriceFile = (Join-Path (Get-Location).Path "CPLicenseAuditor-prices.json"),

    # Optional label describing where the prices came from (shown in the report).
    # e.g. "Invoice 2026-05", "CSP price sheet DE (EUR)", "Manual estimate".
    [string]$PriceSource,

    # Do not prompt; fill any missing prices from the built-in USD hints.
    [switch]$UseListPrices,

    # Do not prompt; use only what is already in the price file (rest = n/a).
    [switch]$NonInteractive,

    # Skip auto-opening the report in the default browser.
    [switch]$NoOpen
)

# ---------------------------------------------------------------------------
# Rough USD list-price HINTS, keyed by SkuPartNumber (as returned by
# Get-MgSubscribedSku). Shown only as a reference during the first-run prompt.
# These are NOT authoritative and may be outdated — the value you type wins.
# ---------------------------------------------------------------------------
$ListPriceHintsUSD = @{
    "ENTERPRISEPACK"           = 23.00   # Office 365 E3
    "ENTERPRISEPREMIUM"        = 38.00   # Office 365 E5
    "SPE_E3"                   = 36.00   # Microsoft 365 E3
    "SPE_E5"                   = 57.00   # Microsoft 365 E5
    "SPB"                      = 22.00   # Microsoft 365 Business Premium
    "O365_BUSINESS_ESSENTIALS" = 6.00    # Microsoft 365 Business Basic
    "O365_BUSINESS_PREMIUM"    = 12.50   # Microsoft 365 Business Standard
    "EXCHANGESTANDARD"         = 4.00    # Exchange Online (Plan 1)
    "EXCHANGEENTERPRISE"       = 8.00    # Exchange Online (Plan 2)
    "POWER_BI_PRO"             = 10.00   # Power BI Pro
    "MCOEV"                    = 8.00    # Teams Phone (Phone System)
    "Microsoft_365_Copilot"    = 30.00   # Microsoft 365 Copilot
}

# Friendly display names keyed by SkuPartNumber. Unknown SKUs fall back to the
# raw SkuPartNumber so nothing is mislabeled.
$FriendlyNameMap = @{
    "ENTERPRISEPACK"           = "Office 365 E3"
    "ENTERPRISEPREMIUM"        = "Office 365 E5"
    "SPE_E3"                   = "Microsoft 365 E3"
    "SPE_E5"                   = "Microsoft 365 E5"
    "SPB"                      = "Microsoft 365 Business Premium"
    "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
    "EXCHANGESTANDARD"         = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE"       = "Exchange Online (Plan 2)"
    "POWER_BI_PRO"             = "Power BI Pro"
    "MCOEV"                    = "Teams Phone (Phone System)"
    "Microsoft_365_Copilot"    = "Microsoft 365 Copilot"
}

$RequiredScopes = @(
    "Organization.Read.All",
    "User.Read.All",
    "Directory.Read.All",
    "AuditLog.Read.All"
)

# Runtime price state (populated from file and/or prompts).
$script:Currency    = $null
$script:PriceData   = @{}
$script:PriceSource = $null

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "OK")]
        [string]$Level = "INFO"
    )
    $stamp = (Get-Date).ToString("HH:mm:ss")
    $color = switch ($Level) {
        "INFO"  { "Gray" }
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        "OK"    { "Green" }
    }
    Write-Host "[$stamp][$Level] $Message" -ForegroundColor $color
}

function Get-SkuFriendlyName {
    param([string]$PartNumber)
    if ($FriendlyNameMap.ContainsKey($PartNumber)) { return $FriendlyNameMap[$PartNumber] }
    return $PartNumber
}

function Get-SkuPrice {
    # Returns a [nullable[double]]; $null means "price unknown".
    param([string]$PartNumber)
    if ($script:PriceData.ContainsKey($PartNumber)) { return [double]$script:PriceData[$PartNumber] }
    return $null
}

function Format-Money {
    param([nullable[double]]$Value)
    if ($null -eq $Value) { return "n/a" }
    $cur = if ($script:Currency) { $script:Currency } else { "USD" }
    return "{0} {1:N2}" -f $cur, $Value
}

function Import-PriceFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $prices = @{}
        if ($raw.prices) {
            foreach ($prop in $raw.prices.PSObject.Properties) { $prices[$prop.Name] = [double]$prop.Value }
        }
        return [pscustomobject]@{ currency = $raw.currency; source = $raw.source; prices = $prices }
    }
    catch {
        Write-Log "Could not read price file '$Path': $($_.Exception.Message)" "WARN"
        return $null
    }
}

function Save-PriceFile {
    param([string]$Path)
    try {
        $obj = [ordered]@{ currency = $script:Currency; source = $script:PriceSource; prices = [ordered]@{} }
        foreach ($k in ($script:PriceData.Keys | Sort-Object)) { $obj.prices[$k] = $script:PriceData[$k] }
        $obj | ConvertTo-Json -Depth 4 | Out-File -FilePath $Path -Encoding utf8
        Write-Log "Prices saved to $Path" "OK"
    }
    catch {
        Write-Log "Could not save price file '$Path': $($_.Exception.Message)" "WARN"
    }
}

function ConvertTo-Price {
    # Culture-tolerant parse: accepts comma or dot decimal. Returns $null on fail.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $norm = $Text.Trim() -replace ',', '.'
    $parsed = 0.0
    $ok = [double]::TryParse(
        $norm,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$parsed)
    if ($ok -and $parsed -ge 0) { return $parsed }
    return $null
}

function Resolve-Prices {
    # Loads the price file, then fills gaps for the SKUs present in the tenant.
    param([string[]]$PartNumbers)

    $existing = Import-PriceFile -Path $PriceFile
    if ($existing) {
        if ($existing.currency) { $script:Currency = $existing.currency }
        if ($existing.source)   { $script:PriceSource = $existing.source }
        foreach ($k in $existing.prices.Keys) { $script:PriceData[$k] = [double]$existing.prices[$k] }
        Write-Log "Loaded $($existing.prices.Count) price(s) from $PriceFile" "INFO"
    }

    # Explicit -PriceSource always wins.
    if ($PriceSource) { $script:PriceSource = $PriceSource }

    # Non-prompting modes ------------------------------------------------------
    if ($UseListPrices) {
        foreach ($p in $PartNumbers) {
            if (-not $script:PriceData.ContainsKey($p) -and $ListPriceHintsUSD.ContainsKey($p)) {
                $script:PriceData[$p] = [double]$ListPriceHintsUSD[$p]
            }
        }
        if (-not $script:Currency)    { $script:Currency = "USD" }
        if (-not $script:PriceSource) { $script:PriceSource = "Built-in USD list hints (not contracted prices)" }
        Save-PriceFile -Path $PriceFile
        return
    }
    if ($NonInteractive) {
        if (-not $script:Currency)    { $script:Currency = "USD" }
        if (-not $script:PriceSource) { $script:PriceSource = "Price file" }
        return
    }

    # Interactive --------------------------------------------------------------
    if (-not $script:Currency) {
        $c = Read-Host "Currency code for this tenant [USD]"
        $script:Currency = if ([string]::IsNullOrWhiteSpace($c)) { "USD" } else { $c.Trim().ToUpper() }
    }

    if (-not $script:PriceSource) {
        $s = Read-Host "Price source label (e.g. 'Invoice 2026-05', 'CSP price sheet') [Manual estimate]"
        $script:PriceSource = if ([string]::IsNullOrWhiteSpace($s)) { "Manual estimate" } else { $s.Trim() }
    }

    $missing = $PartNumbers | Where-Object { -not $script:PriceData.ContainsKey($_) } | Sort-Object -Unique
    if ($missing.Count -gt 0) {
        Write-Host ""
        Write-Host "$($missing.Count) licensed SKU(s) need a monthly price in $($script:Currency)." -ForegroundColor Cyan
        Write-Host "Press Enter to skip a SKU (it will show as n/a in the report)." -ForegroundColor Cyan
        foreach ($p in $missing) {
            $name = Get-SkuFriendlyName $p
            $hint = if ($ListPriceHintsUSD.ContainsKey($p)) { " [hint ~USD {0:N2}]" -f $ListPriceHintsUSD[$p] } else { "" }
            $ans  = Read-Host ("  {0} ({1}){2}" -f $name, $p, $hint)
            $val  = ConvertTo-Price $ans
            if ($null -ne $val) {
                $script:PriceData[$p] = $val
            }
            elseif (-not [string]::IsNullOrWhiteSpace($ans)) {
                Write-Log "  '$ans' is not a valid number; leaving $p as unknown." "WARN"
            }
        }
    }

    Save-PriceFile -Path $PriceFile
}

# ---------------------------------------------------------------------------
# 1. Module check
# ---------------------------------------------------------------------------
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

Write-Log "Checking required Microsoft Graph modules..."
$missingModules = @()
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) { $missingModules += $m }
}
if ($missingModules.Count -gt 0) {
    Write-Log "Missing modules: $($missingModules -join ', ')" "ERROR"
    Write-Log "Install them with:" "INFO"
    Write-Host "    Install-Module $($missingModules -join ', ') -Scope CurrentUser" -ForegroundColor Cyan
    return
}

# ---------------------------------------------------------------------------
# 2. Connect (delegated, interactive). Nothing is stored.
# ---------------------------------------------------------------------------
try {
    Write-Log "Connecting to Microsoft Graph (interactive sign-in)..."
    Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
    $ctx = Get-MgContext
    if (-not $ctx) { throw "No Graph context after connect." }
    Write-Log "Connected as $($ctx.Account) on tenant $($ctx.TenantId)" "OK"
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
    return
}

try {
    # -----------------------------------------------------------------------
    # 3. Subscribed SKUs -> resolve prices, then shelfware analysis
    # -----------------------------------------------------------------------
    Write-Log "Reading subscribed SKUs..."
    $skus = Get-MgSubscribedSku -All -ErrorAction Stop

    # Resolve prices only for the SKUs this tenant actually has.
    $tenantParts = $skus | ForEach-Object { $_.SkuPartNumber } | Sort-Object -Unique
    Resolve-Prices -PartNumbers $tenantParts

    # SkuId (GUID) -> SkuPartNumber lookup, derived from live tenant data
    # so we never hardcode or guess GUIDs.
    $skuIdToPart = @{}
    foreach ($s in $skus) { $skuIdToPart[$s.SkuId] = $s.SkuPartNumber }

    $skuRows = @()
    $shelfwareRows = @()
    $shelfwareSavings = 0.0
    $estimatedMonthlySpend = 0.0

    foreach ($s in $skus) {
        $part      = $s.SkuPartNumber
        $name      = Get-SkuFriendlyName $part
        $enabled   = [int]$s.PrepaidUnits.Enabled
        $consumed  = [int]$s.ConsumedUnits
        $unassigned= [math]::Max(0, $enabled - $consumed)
        $price     = Get-SkuPrice $part
        $usagePct  = if ($enabled -gt 0) { [math]::Round(($consumed / $enabled) * 100, 0) } else { 0 }

        if ($null -ne $price) {
            $estimatedMonthlySpend += ($consumed * $price)
        }

        $skuRows += [pscustomobject]@{
            Name        = $name
            PartNumber  = $part
            Enabled     = $enabled
            Consumed    = $consumed
            Unassigned  = $unassigned
            UsagePct    = $usagePct
            Price       = $price
        }

        if ($unassigned -gt 0) {
            $rowSaving = if ($null -ne $price) { $unassigned * $price } else { $null }
            if ($null -ne $rowSaving) { $shelfwareSavings += $rowSaving }
            $shelfwareRows += [pscustomobject]@{
                Name       = $name
                Unassigned = $unassigned
                Price      = $price
                Saving     = $rowSaving
            }
        }
    }
    Write-Log "Found $($skus.Count) SKUs; estimated monthly spend $(Format-Money $estimatedMonthlySpend)" "OK"

    # -----------------------------------------------------------------------
    # 4. Users -> disabled-with-license and inactive
    # -----------------------------------------------------------------------
    Write-Log "Reading users (this can take a while on large tenants)..."

    # signInActivity requires Entra ID P1/P2. On tenants without it, Graph
    # rejects the WHOLE query (Authentication_RequestFromNonPremiumTenant),
    # so we try with it first and transparently fall back without it.
    $signInDataAvailable = $false
    $users = $null
    try {
        $users = Get-MgUser -All -Property "id","displayName","userPrincipalName","accountEnabled","assignedLicenses","signInActivity" -ErrorAction Stop
        $signInDataAvailable = $true
    }
    catch {
        if ($_.Exception.Message -match "NonPremium|premium license") {
            Write-Log "Tenant has no Entra ID P1/P2; retrying without sign-in activity." "WARN"
            $users = Get-MgUser -All -Property "id","displayName","userPrincipalName","accountEnabled","assignedLicenses" -ErrorAction Stop
            $signInDataAvailable = $false
        }
        else {
            throw
        }
    }

    $disabledRows = @()
    $disabledSavings = 0.0
    $inactiveRows = @()
    $inactivePotential = 0.0
    $cutoff = (Get-Date).AddDays(-$InactiveDays)

    foreach ($u in $users) {
        $licenses = @($u.AssignedLicenses)
        if ($licenses.Count -eq 0) { continue }   # only care about licensed users

        # Resolve this user's license names + total price.
        $licNames = @()
        $userCost = 0.0
        $userCostKnown = $true
        foreach ($l in $licenses) {
            $part = $skuIdToPart[$l.SkuId]
            if ($part) {
                $licNames += (Get-SkuFriendlyName $part)
                $p = Get-SkuPrice $part
                if ($null -ne $p) { $userCost += $p } else { $userCostKnown = $false }
            }
            else {
                $licNames += $l.SkuId   # SKU not in tenant list (rare) -> show GUID
                $userCostKnown = $false
            }
        }
        $userCostValue = if ($userCostKnown) { $userCost } else { $null }

        # 4a. Disabled accounts that still hold licenses.
        if ($u.AccountEnabled -eq $false) {
            if ($null -ne $userCostValue) { $disabledSavings += $userCostValue }
            $disabledRows += [pscustomobject]@{
                DisplayName = $u.DisplayName
                Upn         = $u.UserPrincipalName
                Licenses    = ($licNames -join ", ")
                Cost        = $userCostValue
            }
            continue
        }

        # 4b. Active but inactive (no recent sign-in). Only when P1 data exists.
        if ($signInDataAvailable -and $u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn = $u.SignInActivity.LastSignInDateTime
            if ($lastSignIn -lt $cutoff) {
                if ($null -ne $userCostValue) { $inactivePotential += $userCostValue }
                $inactiveRows += [pscustomobject]@{
                    DisplayName = $u.DisplayName
                    Upn         = $u.UserPrincipalName
                    Licenses    = ($licNames -join ", ")
                    LastSignIn  = $lastSignIn
                    Cost        = $userCostValue
                }
            }
        }
    }

    Write-Log "Disabled-with-license: $($disabledRows.Count) | Inactive ($InactiveDays d): $($inactiveRows.Count)" "OK"
    if (-not $signInDataAvailable) {
        Write-Log "No signInActivity data returned. Tenant likely lacks Entra ID P1, or AuditLog.Read.All was not consented." "WARN"
    }

    $totalRecoverable = $shelfwareSavings + $disabledSavings

    # -----------------------------------------------------------------------
    # 5. Build HTML report
    # -----------------------------------------------------------------------
    Write-Log "Generating HTML report..."

    $skuHtml = ""
    foreach ($r in ($skuRows | Sort-Object UsagePct)) {
        $barColor = if ($r.UsagePct -ge 85) { "#639922" } elseif ($r.UsagePct -ge 60) { "#BA7517" } else { "#E24B4A" }
        $skuHtml += "<tr><td>$($r.Name)</td><td class='num'>$($r.Enabled)</td><td class='num'>$($r.Consumed)</td><td class='num'>$($r.Unassigned)</td><td><div class='bar'><div class='fill' style='width:$($r.UsagePct)%;background:$barColor'></div></div><span class='pct'>$($r.UsagePct)%</span></td><td class='num'>$(Format-Money $r.Price)</td></tr>"
    }

    $shelfHtml = ""
    if ($shelfwareRows.Count -eq 0) {
        $shelfHtml = "<tr><td colspan='4' class='empty'>No unassigned paid licenses found.</td></tr>"
    } else {
        foreach ($r in ($shelfwareRows | Sort-Object { $_.Saving } -Descending)) {
            $shelfHtml += "<tr><td>$($r.Name)</td><td class='num'>$($r.Unassigned)</td><td class='num'>$(Format-Money $r.Price)</td><td class='num save'>$(Format-Money $r.Saving)</td></tr>"
        }
    }

    $disabledHtml = ""
    if ($disabledRows.Count -eq 0) {
        $disabledHtml = "<tr><td colspan='4' class='empty'>No disabled accounts hold licenses.</td></tr>"
    } else {
        foreach ($r in ($disabledRows | Sort-Object { $_.Cost } -Descending)) {
            $disabledHtml += "<tr><td>$($r.DisplayName)</td><td>$($r.Upn)</td><td>$($r.Licenses)</td><td class='num save'>$(Format-Money $r.Cost)</td></tr>"
        }
    }

    $inactiveHtml = ""
    if (-not $signInDataAvailable) {
        $inactiveHtml = "<tr><td colspan='5' class='empty'>Sign-in activity unavailable. Requires Entra ID P1/P2 and AuditLog.Read.All consent.</td></tr>"
    } elseif ($inactiveRows.Count -eq 0) {
        $inactiveHtml = "<tr><td colspan='5' class='empty'>No inactive licensed users beyond $InactiveDays days.</td></tr>"
    } else {
        foreach ($r in ($inactiveRows | Sort-Object LastSignIn)) {
            $lastStr = $r.LastSignIn.ToString("yyyy-MM-dd")
            $inactiveHtml += "<tr><td>$($r.DisplayName)</td><td>$($r.Upn)</td><td>$($r.Licenses)</td><td>$lastStr</td><td class='num'>$(Format-Money $r.Cost)</td></tr>"
        }
    }

    $generated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    $tenant = $ctx.TenantId
    $account = $ctx.Account
    $priceSourceLabel = if ($script:PriceSource) { $script:PriceSource } else { "unspecified" }

    $html = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CloudPersistence License Auditor</title>
<style>
  body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1a1a1a;background:#f7f6f2;margin:0;padding:32px;}
  .wrap{max-width:1080px;margin:0 auto;}
  h1{font-size:22px;font-weight:600;margin:0 0 4px;}
  .sub{color:#666;font-size:13px;margin-bottom:24px;}
  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:28px;}
  .card{background:#fff;border:1px solid #e6e3da;border-radius:12px;padding:16px;}
  .card .label{font-size:12px;color:#666;margin-bottom:6px;}
  .card .val{font-size:22px;font-weight:600;}
  .recover{color:#3B6D11;}
  .waste{color:#A32D2D;}
  .section{background:#fff;border:1px solid #e6e3da;border-radius:12px;padding:20px;margin-bottom:16px;}
  .section h2{font-size:15px;font-weight:600;margin:0 0 14px;}
  table{width:100%;border-collapse:collapse;font-size:13px;}
  th{text-align:left;font-size:12px;color:#666;font-weight:600;padding:0 10px 8px 0;border-bottom:1px solid #e6e3da;}
  td{padding:9px 10px 9px 0;border-bottom:1px solid #f0eee7;}
  tr:last-child td{border-bottom:none;}
  .num{text-align:right;}
  .save{color:#3B6D11;font-weight:600;}
  .empty{color:#999;font-style:italic;text-align:center;padding:14px;}
  .bar{display:inline-block;width:70px;height:6px;background:#eee;border-radius:3px;overflow:hidden;vertical-align:middle;}
  .fill{height:100%;}
  .pct{font-size:11px;color:#666;margin-left:6px;}
  .footer{font-size:11px;color:#999;margin-top:8px;line-height:1.6;}
  .brand{color:#185FA5;font-weight:600;}
</style></head><body><div class="wrap">
<h1>CloudPersistence License Auditor — M365 license waste report</h1>
<div class="sub">Tenant $tenant &middot; run by $account &middot; generated $generated</div>

<div class="cards">
  <div class="card"><div class="label">Estimated monthly spend</div><div class="val">$(Format-Money $estimatedMonthlySpend)</div></div>
  <div class="card"><div class="label">Shelfware (unassigned)</div><div class="val waste">$(Format-Money $shelfwareSavings)</div></div>
  <div class="card"><div class="label">Licensed disabled users</div><div class="val waste">$(Format-Money $disabledSavings)</div></div>
  <div class="card"><div class="label">Recoverable / month</div><div class="val recover">$(Format-Money $totalRecoverable)</div></div>
</div>

<div class="section"><h2>License utilization by plan</h2>
<table><thead><tr><th>Plan</th><th class="num">Enabled</th><th class="num">Assigned</th><th class="num">Unassigned</th><th>Usage</th><th class="num">Unit price</th></tr></thead>
<tbody>$skuHtml</tbody></table></div>

<div class="section"><h2>Shelfware — paid but unassigned</h2>
<table><thead><tr><th>Plan</th><th class="num">Unassigned</th><th class="num">Unit price</th><th class="num">Monthly waste</th></tr></thead>
<tbody>$shelfHtml</tbody></table></div>

<div class="section"><h2>Licenses on disabled accounts</h2>
<table><thead><tr><th>User</th><th>UPN</th><th>Licenses</th><th class="num">Monthly cost</th></tr></thead>
<tbody>$disabledHtml</tbody></table></div>

<div class="section"><h2>Inactive licensed users (no sign-in $InactiveDays+ days)</h2>
<table><thead><tr><th>User</th><th>UPN</th><th>Licenses</th><th>Last sign-in</th><th class="num">Monthly cost</th></tr></thead>
<tbody>$inactiveHtml</tbody></table></div>

<div class="footer">
Price source: <b>$priceSourceLabel</b> (currency $($script:Currency)). Verify these match the client's contracted rates before acting.
Inactive users are flagged for review only; some may be legitimate service or shared accounts.
This report was generated locally; no tenant data was transmitted anywhere.
Powered by <span class="brand">cloudpersistence.com</span>
</div>
</div></body></html>
"@

    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $fileName = "CPLicenseAuditor-Report-$((Get-Date).ToString('yyyyMMdd-HHmmss')).html"
    $outPath = Join-Path $OutputFolder $fileName
    $html | Out-File -FilePath $outPath -Encoding utf8
    Write-Log "Report written to: $outPath" "OK"

    if (-not $NoOpen) { Start-Process $outPath }
}
catch {
    Write-Log "Audit failed: $($_.Exception.Message)" "ERROR"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Disconnected from Microsoft Graph."
}
