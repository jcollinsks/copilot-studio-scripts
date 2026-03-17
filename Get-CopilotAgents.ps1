<#
.SYNOPSIS
    Lists all Copilot Studio agents (bots) in a Power Platform Dataverse environment.

.DESCRIPTION
    Authenticates via Connect-AzAccount and queries the Dataverse Web API to retrieve
    all Copilot Studio agents. Also queries the Power Platform Admin API to include
    system-generated and demo bots that Microsoft provisions. Handles pagination
    automatically.

.PARAMETER EnvironmentUrl
    The Dataverse environment URL (e.g., https://yourorg.crm.dynamics.com).

.EXAMPLE
    .\Get-CopilotAgents.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com"

.EXAMPLE
    .\Get-CopilotAgents.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com" | Export-Csv agents.csv

.NOTES
    Requires the Az.Accounts module: Install-Module Az.Accounts
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EnvironmentUrl
)

$ErrorActionPreference = 'Stop'

# Helper: extract plain-text token from Get-AzAccessToken (handles both old string and new SecureString formats)
function Get-PlainToken {
    param([object]$TokenResult)
    if ($TokenResult.Token -is [System.Security.SecureString]) {
        return $TokenResult.Token | ConvertFrom-SecureString -AsPlainText
    }
    return $TokenResult.Token
}

# Helper: safely extract error details from an ErrorRecord (immune to StrictMode leaking from modules)
function Get-ErrorDetail {
    param([object]$ErrorRecord)
    $code = $null
    $msg = $null
    try { $code = $ErrorRecord.Exception.Response.StatusCode.value__ } catch {}
    try { $msg = $ErrorRecord.ErrorDetails.Message } catch {}
    if (-not $msg) {
        try { $msg = $ErrorRecord.Exception.Message } catch { $msg = "$ErrorRecord" }
    }
    return @{ StatusCode = $code; Detail = $msg }
}

# Normalize the environment URL (remove trailing slash)
$EnvironmentUrl = $EnvironmentUrl.TrimEnd('/')

# --- Authentication ---
try {
    $context = Get-AzContext
    if (-not $context) {
        Write-Host "No active Azure session found. Launching interactive login..."
        Connect-AzAccount | Out-Null
        $context = Get-AzContext
    }
    Write-Host "Authenticated as: $($context.Account.Id)" -ForegroundColor Green
}
catch {
    Write-Error "Authentication failed. Ensure the Az.Accounts module is installed (Install-Module Az.Accounts). Error: $_"
    exit 1
}

# --- Get Access Token for Dataverse ---
try {
    $tokenResult = Get-AzAccessToken -ResourceUrl $EnvironmentUrl
    $accessToken = Get-PlainToken $tokenResult
}
catch {
    Write-Error "Failed to acquire access token for '$EnvironmentUrl'. Ensure you have permissions to this environment. Error: $_"
    exit 1
}

$headers = @{
    Authorization  = "Bearer $accessToken"
    Accept         = "application/json"
    "OData-Version" = "4.0"
}

# --- Query Bots with Pagination ---
$selectFields = "botid,name,schemaname,statecode,statuscode,language,createdon,modifiedon"
$requestUrl = "$EnvironmentUrl/api/data/v9.2/bots?`$select=$selectFields&`$orderby=name"

$allBots = [System.Collections.Generic.List[PSObject]]::new()

Write-Host "Querying Copilot Studio agents from $EnvironmentUrl ..." -ForegroundColor Cyan

while ($requestUrl) {
    try {
        $response = Invoke-RestMethod -Uri $requestUrl -Headers $headers -Method Get
    }
    catch {
        $err = Get-ErrorDetail $_
        Write-Error "Dataverse API request failed (HTTP $($err.StatusCode)). $($err.Detail)"
        exit 1
    }

    if ($response.value) {
        foreach ($bot in $response.value) {
            $allBots.Add($bot)
        }
    }

    # Follow pagination link if present
    $requestUrl = $response.'@odata.nextLink'
}

# Map of botId -> applicationId (populated from Admin API)
$botAppIdMap = @{}

# --- Query Power Platform Admin API for system/demo bots ---
# Resolve environment ID from Dataverse
$environmentId = $null
try {
    $orgUrl = "$EnvironmentUrl/api/data/v9.2/organizations?`$select=organizationid,environmentid"
    $orgResponse = Invoke-RestMethod -Uri $orgUrl -Headers $headers -Method Get
    if ($orgResponse.value -and $orgResponse.value.Count -gt 0) {
        $environmentId = $orgResponse.value[0].environmentid
        if (-not $environmentId) {
            $environmentId = $orgResponse.value[0].organizationid
        }
    }
}
catch {
    Write-Warning "Could not resolve environment ID. System/demo bots may not be included."
}

if ($environmentId) {
    Write-Host "Querying Power Platform Admin API for system/demo bots..." -ForegroundColor Cyan
    try {
        $ppToken = Get-PlainToken (Get-AzAccessToken -ResourceUrl "https://api.powerplatform.com")
        $ppHeaders = @{
            Authorization = "Bearer $ppToken"
            Accept        = "application/json"
        }

        $adminBotsUrl = "https://api.powerplatform.com/copilotstudio/environments/$environmentId/bots?api-version=2022-03-01-preview"
        $adminResponse = Invoke-RestMethod -Uri $adminBotsUrl -Headers $ppHeaders -Method Get

        # Collect bot IDs already found via Dataverse
        $existingBotIds = @{}
        foreach ($b in $allBots) {
            if ($b.botid) { $existingBotIds[$b.botid] = $true }
        }

        # Map bot IDs to application IDs from Admin API (used to enrich Dataverse results too)
        $adminBots = if ($adminResponse.value) { $adminResponse.value } else { @() }
        foreach ($adminBot in $adminBots) {
            $abId = $adminBot.botId
            if (-not $abId) { $abId = $adminBot.id }
            $abAppId = $adminBot.applicationId
            if ($abId -and $abAppId) {
                $botAppIdMap[$abId] = $abAppId
            }
        }

        $adminBotCount = 0
        foreach ($adminBot in $adminBots) {
            $adminBotId = $adminBot.botId
            if (-not $adminBotId) { $adminBotId = $adminBot.id }
            if ($adminBotId -and -not $existingBotIds.ContainsKey($adminBotId)) {
                # Create a normalized object that matches the Dataverse shape
                $syntheticBot = [PSCustomObject]@{
                    botid      = $adminBotId
                    name       = if ($adminBot.displayName) { $adminBot.displayName } else { $adminBot.name }
                    schemaname = $adminBot.schemaName
                    statecode  = $adminBot.state
                    statuscode = $adminBot.statusCode
                    language   = $adminBot.language
                    createdon  = $adminBot.createdOn
                    modifiedon = $adminBot.modifiedOn
                    Source     = 'AdminAPI'
                }
                $allBots.Add($syntheticBot)
                $adminBotCount++
            }
        }

        if ($adminBotCount -gt 0) {
            Write-Host "Found $adminBotCount additional bot(s) from Admin API (system/demo)." -ForegroundColor Green
        }
    }
    catch {
        $err = Get-ErrorDetail $_
        Write-Warning "Power Platform Admin API query failed (HTTP $($err.StatusCode)). System/demo bots may not be included. $($err.Detail)"
    }
}

# --- Look up Azure AD App Registrations and Service Principals via Microsoft Graph ---
$graphToken = $null
try {
    $graphToken = Get-PlainToken (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com")
}
catch {
    Write-Warning "Could not acquire Microsoft Graph token. App/Service Principal columns will be empty."
}

$graphHeaders = @{}
if ($graphToken) {
    $graphHeaders = @{
        Authorization = "Bearer $graphToken"
        Accept        = "application/json"
    }
}

# Build a cache of app registrations and service principals keyed by appId
$appCache = @{}
$spCache = @{}

if ($graphToken) {
    # Collect unique application IDs from Admin API map
    $appIds = $botAppIdMap.Values | Sort-Object -Unique

    foreach ($appId in $appIds) {
        # Look up app registration
        try {
            $appResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$appId'&`$select=id,displayName,appId" -Headers $graphHeaders -Method Get
            if ($appResult.value -and $appResult.value.Count -gt 0) {
                $appCache[$appId] = $appResult.value[0]
            }
        }
        catch {
            Write-Warning "Could not look up app registration for appId $appId"
        }

        # Look up service principal
        try {
            $spResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appId'&`$select=id,displayName,appId" -Headers $graphHeaders -Method Get
            if ($spResult.value -and $spResult.value.Count -gt 0) {
                $spCache[$appId] = $spResult.value[0]
            }
        }
        catch {
            Write-Warning "Could not look up service principal for appId $appId"
        }
    }
}

# --- Build Output Objects ---
$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($bot in $allBots) {
    $appId = if ($botAppIdMap.ContainsKey($bot.botid)) { $botAppIdMap[$bot.botid] } else { $null }
    $app = if ($appId -and $appCache.ContainsKey($appId)) { $appCache[$appId] } else { $null }
    $sp  = if ($appId -and $spCache.ContainsKey($appId))  { $spCache[$appId] }  else { $null }

    $source = if ($bot.Source -eq 'AdminAPI') { 'AdminAPI' } else { 'Dataverse' }
    $obj = [PSCustomObject]@{
        Name               = $bot.name
        BotId              = $bot.botid
        SchemaName         = $bot.schemaname
        StateCode          = $bot.statecode
        StatusCode         = $bot.statuscode
        Language           = $bot.language
        CreatedOn          = $bot.createdon
        ModifiedOn         = $bot.modifiedon
        ApplicationId      = $appId
        AppDisplayName     = if ($app) { $app.displayName } else { $null }
        AppObjectId        = if ($app) { $app.id } else { $null }
        ServicePrincipalId = if ($sp) { $sp.id } else { $null }
        SPDisplayName      = if ($sp) { $sp.displayName } else { $null }
        Source             = $source
    }
    $results.Add($obj)
}

# --- Output Results ---
if ($results.Count -eq 0) {
    Write-Host "No Copilot Studio agents found in this environment." -ForegroundColor Yellow
}
else {
    Write-Host "Found $($results.Count) agent(s):" -ForegroundColor Green
    $results | Format-Table -Property Name, BotId, ApplicationId, AppDisplayName, ServicePrincipalId, StateCode, Language, Source -AutoSize | Out-Host
}

# Write objects to pipeline for downstream use (Export-Csv, Where-Object, etc.)
$results
