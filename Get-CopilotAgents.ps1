<#
.SYNOPSIS
    Lists all Copilot Studio agents (bots) in a Power Platform Dataverse environment.

.DESCRIPTION
    Authenticates via Connect-AzAccount and queries the Dataverse Web API to retrieve
    all Copilot Studio agents. Handles pagination automatically.

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
$selectFields = "botid,name,schemaname,statecode,statuscode,language,createdon,modifiedon,applicationid"
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
    # Collect unique application IDs
    $appIds = $allBots | Where-Object { $_.applicationid } | ForEach-Object { $_.applicationid } | Sort-Object -Unique

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
    $appId = $bot.applicationid
    $app = if ($appId -and $appCache.ContainsKey($appId)) { $appCache[$appId] } else { $null }
    $sp  = if ($appId -and $spCache.ContainsKey($appId))  { $spCache[$appId] }  else { $null }

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
    }
    $results.Add($obj)
}

# --- Output Results ---
if ($results.Count -eq 0) {
    Write-Host "No Copilot Studio agents found in this environment." -ForegroundColor Yellow
}
else {
    Write-Host "Found $($results.Count) agent(s):" -ForegroundColor Green
    $results | Format-Table -Property Name, BotId, ApplicationId, AppDisplayName, ServicePrincipalId, StateCode, Language -AutoSize | Out-Host
}

# Write objects to pipeline for downstream use (Export-Csv, Where-Object, etc.)
$results
