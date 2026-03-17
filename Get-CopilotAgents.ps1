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

# --- Query Bots with Pagination (no $select — retrieve all fields) ---
$requestUrl = "$EnvironmentUrl/api/data/v9.2/bots?`$orderby=name"

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

# Dump all property names from the first bot for diagnostics
if ($allBots.Count -gt 0) {
    $propNames = ($allBots[0] | Get-Member -MemberType NoteProperty).Name | Where-Object { $_ -notlike '@odata*' }
    Write-Host "`nDataverse bot entity fields:" -ForegroundColor DarkGray
    Write-Host ($propNames -join ', ') -ForegroundColor DarkGray
    Write-Host ""
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

        # Collect bot IDs already found via Dataverse (normalize to lowercase)
        $existingBotIds = @{}
        foreach ($b in $allBots) {
            if ($b.botid) { $existingBotIds[$b.botid.ToString().ToLower()] = $true }
        }

        # Map bot IDs to application IDs from Admin API (used to enrich Dataverse results too)
        # Try multiple possible field names since the API schema varies
        $adminBots = if ($adminResponse.value) { $adminResponse.value } else { @() }

        if ($adminBots.Count -gt 0) {
            # Log available properties from the first bot for diagnostics
            $sampleProps = ($adminBots[0] | Get-Member -MemberType NoteProperty).Name -join ', '
            Write-Verbose "Admin API bot properties: $sampleProps"
        }

        foreach ($adminBot in $adminBots) {
            $abId = $adminBot.botId
            if (-not $abId) { $abId = $adminBot.id }
            if ($abId) { $abId = $abId.ToString().ToLower() }

            # Try multiple possible field names for the application ID
            $abAppId = $null
            foreach ($fieldName in @('applicationId', 'appId', 'aadApplicationId', 'clientId', 'azureApplicationId')) {
                $val = $adminBot.$fieldName
                if ($val) { $abAppId = $val; break }
            }
            # Also check nested properties object
            if (-not $abAppId -and $adminBot.properties) {
                foreach ($fieldName in @('applicationId', 'appId', 'aadApplicationId', 'clientId')) {
                    $val = $adminBot.properties.$fieldName
                    if ($val) { $abAppId = $val; break }
                }
            }
            if ($abId -and $abAppId) {
                $botAppIdMap[$abId] = $abAppId
            }
        }

        $adminBotCount = 0
        foreach ($adminBot in $adminBots) {
            $adminBotId = $adminBot.botId
            if (-not $adminBotId) { $adminBotId = $adminBot.id }
            if ($adminBotId) { $adminBotId = $adminBotId.ToString().ToLower() }
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

# --- Extract App IDs from botcomponent entity (definitive source) ---
Write-Host "Querying bot components for Azure AD app registrations..." -ForegroundColor Cyan

$bcUrl = "$EnvironmentUrl/api/data/v9.2/botcomponents?`$select=_bot_botid_value,componenttype,data,schemaname,name&`$orderby=_bot_botid_value"
while ($bcUrl) {
    try {
        $bcResponse = Invoke-RestMethod -Uri $bcUrl -Headers $headers -Method Get
    }
    catch {
        $err = Get-ErrorDetail $_
        Write-Warning "Could not query botcomponents (HTTP $($err.StatusCode)). $($err.Detail)"
        break
    }

    if ($bcResponse.value) {
        foreach ($bc in $bcResponse.value) {
            $parentBotId = $bc._bot_botid_value
            if (-not $parentBotId) { continue }
            $parentBotId = $parentBotId.ToString().ToLower()
            # Skip if we already found an app ID for this bot
            if ($botAppIdMap.ContainsKey($parentBotId)) { continue }

            # Try to parse the 'data' field as JSON and look for app/client ID fields
            $dataStr = $bc.data
            if (-not $dataStr) { continue }
            try {
                $dataObj = $dataStr | ConvertFrom-Json
            }
            catch { continue }

            # Search common field names in the JSON
            $foundAppId = $null
            foreach ($field in @('applicationId', 'appId', 'aadApplicationId', 'clientId', 'ApplicationId', 'AppId', 'ClientId', 'msAppId')) {
                $val = $dataObj.$field
                if ($val -and $val -match '^[0-9a-fA-F\-]{36}$') {
                    $foundAppId = $val
                    break
                }
            }
            # Also check nested 'settings', 'configuration', 'properties' objects
            if (-not $foundAppId) {
                foreach ($nested in @('settings', 'configuration', 'properties', 'authenticationConfiguration')) {
                    $nestedObj = $dataObj.$nested
                    if (-not $nestedObj) { continue }
                    foreach ($field in @('applicationId', 'appId', 'clientId', 'ApplicationId', 'AppId', 'ClientId', 'msAppId')) {
                        $val = $nestedObj.$field
                        if ($val -and $val -match '^[0-9a-fA-F\-]{36}$') {
                            $foundAppId = $val
                            break
                        }
                    }
                    if ($foundAppId) { break }
                }
            }

            if ($foundAppId) {
                $botAppIdMap[$parentBotId] = $foundAppId
            }
        }
    }

    $bcUrl = $bcResponse.'@odata.nextLink'
}

if ($botAppIdMap.Count -gt 0) {
    Write-Host "  Found app IDs for $($botAppIdMap.Count) bot(s) from bot components." -ForegroundColor Green
}
else {
    Write-Host "  No app IDs found in bot components." -ForegroundColor Yellow
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

# Build caches: appId-based (from botcomponent/Admin API) and name-based (Graph fallback)
$appCache = @{}
$spCache = @{}
$appByNameCache = @{}
$spByNameCache = @{}

if ($graphToken) {
    Write-Host "Looking up Azure AD app registrations and service principals..." -ForegroundColor Cyan

    # 1. Look up by appId for bots where we have a definitive mapping
    if ($botAppIdMap.Count -gt 0) {
        $appIds = $botAppIdMap.Values | Sort-Object -Unique
        foreach ($appId in $appIds) {
            try {
                $appResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$appId'&`$select=id,displayName,appId" -Headers $graphHeaders -Method Get
                if ($appResult.value -and $appResult.value.Count -gt 0) {
                    $appCache[$appId] = $appResult.value[0]
                }
            }
            catch {}

            try {
                $spResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$appId'&`$select=id,displayName,appId" -Headers $graphHeaders -Method Get
                if ($spResult.value -and $spResult.value.Count -gt 0) {
                    $spCache[$appId] = $spResult.value[0]
                }
            }
            catch {}
        }
    }

    # 2. For bots still without an app ID, search Graph using the Copilot Studio naming pattern
    foreach ($bot in $allBots) {
        $botId = $bot.botid.ToString().ToLower()
        if ($botAppIdMap.ContainsKey($botId)) { continue }
        $botName = $bot.name
        if (-not $botName) { continue }
        if ($appByNameCache.ContainsKey($botName)) { continue }

        # Copilot Studio names app registrations as "BOTNAME (Microsoft Copilot Studio)"
        $csName = "$botName (Microsoft Copilot Studio)" -replace "'", "''"
        $escapedName = $botName -replace "'", "''"

        # Try the Copilot Studio naming pattern first, then exact name
        $found = $false
        foreach ($searchName in @($csName, $escapedName)) {
            try {
                $appResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$searchName'&`$select=id,displayName,appId&`$top=1" -Headers $graphHeaders -Method Get
                if ($appResult.value -and $appResult.value.Count -gt 0) {
                    $appByNameCache[$botName] = $appResult.value[0]
                    Write-Host "  Matched app for '$botName' -> '$($appResult.value[0].displayName)' (appId: $($appResult.value[0].appId))" -ForegroundColor Gray
                    $found = $true
                    break
                }
            }
            catch {}
        }
        if (-not $found) {
            Write-Host "  No app registration found for '$botName'" -ForegroundColor DarkGray
        }

        # Service principal lookup (try same patterns)
        foreach ($searchName in @($csName, $escapedName)) {
            try {
                $spResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=displayName eq '$searchName'&`$select=id,displayName,appId&`$top=1" -Headers $graphHeaders -Method Get
                if ($spResult.value -and $spResult.value.Count -gt 0) {
                    $spByNameCache[$botName] = $spResult.value[0]
                    break
                }
            }
            catch {}
        }
    }
}

# --- Build Output Objects ---
$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($bot in $allBots) {
    $botId = $bot.botid.ToString().ToLower()
    $botName = $bot.name

    # Try app ID from botcomponent/Admin API map first
    $appId = if ($botAppIdMap.ContainsKey($botId)) { $botAppIdMap[$botId] } else { $null }
    $app = if ($appId -and $appCache.ContainsKey($appId)) { $appCache[$appId] } else { $null }
    $sp  = if ($appId -and $spCache.ContainsKey($appId))  { $spCache[$appId] }  else { $null }

    # Fall back to name-based Graph lookup
    if (-not $app -and $botName -and $appByNameCache.ContainsKey($botName)) {
        $app = $appByNameCache[$botName]
        if (-not $appId) { $appId = $app.appId }
    }
    if (-not $sp -and $botName -and $spByNameCache.ContainsKey($botName)) {
        $sp = $spByNameCache[$botName]
        if (-not $appId) { $appId = $sp.appId }
    }

    $source = if ($bot.Source -eq 'AdminAPI') { 'AdminAPI' } else { 'Dataverse' }
    $obj = [PSCustomObject]@{
        Name               = $botName
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
