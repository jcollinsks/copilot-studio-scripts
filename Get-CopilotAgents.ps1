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
            $obj = [PSCustomObject]@{
                Name       = $bot.name
                BotId      = $bot.botid
                SchemaName = $bot.schemaname
                StateCode  = $bot.statecode
                StatusCode = $bot.statuscode
                Language   = $bot.language
                CreatedOn  = $bot.createdon
                ModifiedOn = $bot.modifiedon
            }
            $allBots.Add($obj)
        }
    }

    # Follow pagination link if present
    $requestUrl = $response.'@odata.nextLink'
}

# --- Output Results ---
if ($allBots.Count -eq 0) {
    Write-Host "No Copilot Studio agents found in this environment." -ForegroundColor Yellow
}
else {
    Write-Host "Found $($allBots.Count) agent(s):" -ForegroundColor Green
    $allBots | Format-Table -Property Name, BotId, SchemaName, StateCode, StatusCode, Language, CreatedOn, ModifiedOn -AutoSize | Out-Host
}

# Write objects to pipeline for downstream use (Export-Csv, Where-Object, etc.)
$allBots
