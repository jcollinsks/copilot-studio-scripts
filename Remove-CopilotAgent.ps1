<#
.SYNOPSIS
    Deletes a Copilot Studio agent (bot) from a Power Platform Dataverse environment.

.DESCRIPTION
    Authenticates via Connect-AzAccount and deletes the specified Copilot Studio agent.
    The agent can be identified by display name (-BotName) or GUID (-BotId).
    Tries the Power Platform Admin API first, then falls back to the Dataverse PvaDeleteBot action.

.PARAMETER BotName
    The display name of the bot to delete. Mutually exclusive with -BotId.

.PARAMETER BotId
    The GUID of the bot to delete. Mutually exclusive with -BotName.

.PARAMETER EnvironmentUrl
    The Dataverse environment URL (e.g., https://yourorg.crm.dynamics.com).

.PARAMETER EnvironmentId
    The Power Platform environment GUID. If not provided, the script extracts it from the environment.

.PARAMETER Force
    Skip the confirmation prompt before deleting.

.EXAMPLE
    .\Remove-CopilotAgent.ps1 -BotName "Test Bot" -EnvironmentUrl "https://yourorg.crm.dynamics.com"

.EXAMPLE
    .\Remove-CopilotAgent.ps1 -BotId "a1b2c3d4-e5f6-7890-abcd-ef1234567890" -EnvironmentUrl "https://yourorg.crm.dynamics.com" -Force

.NOTES
    Requires the Az.Accounts module: Install-Module Az.Accounts
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ByName')]
    [ValidateNotNullOrEmpty()]
    [string]$BotName,

    [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
    [ValidateNotNullOrEmpty()]
    [string]$BotId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EnvironmentUrl,

    [Parameter(Mandatory = $false)]
    [string]$EnvironmentId,

    [Parameter(Mandatory = $false)]
    [switch]$Force
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
    $dataverseToken = Get-PlainToken $tokenResult
}
catch {
    Write-Error "Failed to acquire access token for '$EnvironmentUrl'. Ensure you have permissions to this environment. Error: $_"
    exit 1
}

$dataverseHeaders = @{
    Authorization   = "Bearer $dataverseToken"
    Accept          = "application/json"
    "OData-Version" = "4.0"
}

# --- Resolve Bot ID from Name ---
$botDetails = $null

if ($PSCmdlet.ParameterSetName -eq 'ByName') {
    Write-Host "Looking up bot by name: '$BotName' ..." -ForegroundColor Cyan

    $filter = "name eq '$($BotName -replace "'", "''")'"
    $lookupUrl = "$EnvironmentUrl/api/data/v9.2/bots?`$filter=$filter&`$select=botid,name,schemaname,statecode,statuscode,language"

    try {
        $lookupResponse = Invoke-RestMethod -Uri $lookupUrl -Headers $dataverseHeaders -Method Get
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "Failed to query bots (HTTP $statusCode). $detail"
        exit 1
    }

    $matches = $lookupResponse.value
    if (-not $matches -or $matches.Count -eq 0) {
        Write-Error "No bot found with name '$BotName' in environment '$EnvironmentUrl'."
        exit 1
    }
    if ($matches.Count -gt 1) {
        Write-Error "Multiple bots ($($matches.Count)) found with name '$BotName'. Use -BotId to specify the exact bot:"
        foreach ($m in $matches) {
            Write-Error "  - BotId: $($m.botid)  Name: $($m.name)"
        }
        exit 1
    }

    $botDetails = $matches[0]
    $BotId = $botDetails.botid
    Write-Host "Resolved bot '$BotName' to BotId: $BotId" -ForegroundColor Green
}
else {
    # Fetch bot details by ID for confirmation display
    Write-Host "Looking up bot by ID: $BotId ..." -ForegroundColor Cyan
    $detailUrl = "$EnvironmentUrl/api/data/v9.2/bots($BotId)?`$select=botid,name,schemaname,statecode,statuscode,language"

    try {
        $botDetails = Invoke-RestMethod -Uri $detailUrl -Headers $dataverseHeaders -Method Get
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 404) {
            Write-Error "No bot found with ID '$BotId' in environment '$EnvironmentUrl'."
            exit 1
        }
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "Failed to query bot (HTTP $statusCode). $detail"
        exit 1
    }
}

# --- Display Bot Details and Confirm ---
Write-Host ""
Write-Host "Bot to delete:" -ForegroundColor Yellow
Write-Host "  Name:       $($botDetails.name)"
Write-Host "  BotId:      $($botDetails.botid)"
Write-Host "  SchemaName: $($botDetails.schemaname)"
Write-Host "  StateCode:  $($botDetails.statecode)"
Write-Host "  StatusCode: $($botDetails.statuscode)"
Write-Host "  Language:   $($botDetails.language)"
Write-Host ""

if (-not $Force) {
    $confirmation = Read-Host "Are you sure you want to delete this bot? (yes/no)"
    if ($confirmation -notin @('yes', 'y')) {
        Write-Host "Deletion cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# --- Resolve Environment ID (if not provided) ---
if (-not $EnvironmentId) {
    Write-Host "Resolving Power Platform environment ID..." -ForegroundColor Cyan
    try {
        $orgUrl = "$EnvironmentUrl/api/data/v9.2/organizations?`$select=organizationid,environmentid"
        $orgResponse = Invoke-RestMethod -Uri $orgUrl -Headers $dataverseHeaders -Method Get
        if ($orgResponse.value -and $orgResponse.value.Count -gt 0) {
            # Try environmentid field first; fall back to organizationid
            $EnvironmentId = $orgResponse.value[0].environmentid
            if (-not $EnvironmentId) {
                $EnvironmentId = $orgResponse.value[0].organizationid
            }
        }
    }
    catch {
        Write-Warning "Could not resolve environment ID from Dataverse. Admin API deletion may fail."
    }
}

$deleted = $false

# --- Attempt 1: Power Platform Admin API ---
if ($EnvironmentId) {
    Write-Host "Attempting deletion via Power Platform Admin API..." -ForegroundColor Cyan

    try {
        $ppToken = Get-PlainToken (Get-AzAccessToken -ResourceUrl "https://api.powerplatform.com")
        $ppHeaders = @{
            Authorization = "Bearer $ppToken"
            Accept        = "application/json"
        }

        $adminUrl = "https://api.powerplatform.com/copilotstudio/environments/$EnvironmentId/bots/$BotId/api/botAdminOperations?api-version=2022-03-01-preview"

        Invoke-RestMethod -Uri $adminUrl -Headers $ppHeaders -Method Delete
        $deleted = $true
        Write-Host "Bot '$($botDetails.name)' ($BotId) deleted successfully via Admin API." -ForegroundColor Green
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = $_.Exception.Response.StatusCode.value__
        }
        Write-Warning "Admin API deletion failed (HTTP $statusCode). Falling back to Dataverse PvaDeleteBot action..."
    }
}
else {
    Write-Warning "No environment ID available. Skipping Admin API, trying Dataverse PvaDeleteBot..."
}

# --- Attempt 2: Dataverse PvaDeleteBot Action ---
if (-not $deleted) {
    Write-Host "Attempting deletion via Dataverse PvaDeleteBot action..." -ForegroundColor Cyan

    $pvaUrl = "$EnvironmentUrl/api/data/v9.2/bots($BotId)/Microsoft.Dynamics.CRM.PvaDeleteBot"

    try {
        Invoke-RestMethod -Uri $pvaUrl -Headers $dataverseHeaders -Method Post -ContentType "application/json" -Body "{}"
        $deleted = $true
        Write-Host "Bot '$($botDetails.name)' ($BotId) deleted successfully via PvaDeleteBot." -ForegroundColor Green
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        Write-Error "PvaDeleteBot failed (HTTP $statusCode). $detail"
        exit 1
    }
}

if ($deleted) {
    Write-Host "Done. Bot '$($botDetails.name)' has been removed." -ForegroundColor Green
}
