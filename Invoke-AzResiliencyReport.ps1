<#
.SYNOPSIS
    Generates an Azure Resiliency Hub report by creating/updating a Service Group,
    assigning Zonal Resiliency goals, fetching posture and recommendations, and
    exporting everything to an Excel workbook with portal deep links.

.DESCRIPTION
    This script automates the manual process that field engineers and CSAs perform
    today for Azure Resiliency assessments. It:

      1. Creates or updates a Resiliency Service Group (application) from resources
         discovered in a subscription or provided via an explicit resource list.
      2. Assigns Zonal Resiliency goals using a default template.
      3. Fetches the zonal resiliency summary, per-resource posture, and
         recommendations for the service group.
      4. Produces an Excel workbook (Overview, Resources, Recommendations, ApiTrace)
         with strategic portal deep links to encourage self-serve adoption.

    All REST calls use Invoke-AzRestMethod against the Azure Resiliency platform
    APIs (preview). Endpoints, API versions, and request bodies are fully
    overridable via -ApiConfigPath without editing code.

.PARAMETER SubscriptionId
    (Mandatory) The Azure subscription ID to target.

.PARAMETER ResourceListPath
    (Optional) Path to a JSON or CSV file containing explicit ARM resource IDs.
    JSON format: an array of strings or objects with a "resourceId" property.
    CSV format: must have a column named "ResourceId".
    If omitted, all supported resources in the subscription are discovered.

.PARAMETER ResourceGroupName
    (Optional) Name of a resource group to scope resource discovery.
    If specified, only resources in this resource group are included.
    If omitted, all resources in the subscription are discovered.

.PARAMETER ServiceGroupName
    (Mandatory) Name for the Service Group. This is the groupId used in the
    Azure Service Group REST API and cannot be changed after creation.

.PARAMETER ServiceGroupLocation
    (Optional) Azure region for the Service Group resource. Default: "eastus".

.PARAMETER OutputPath
    (Optional) Path for the output Excel file. Defaults to
    .\ResiliencyReport_<SubscriptionId>_<yyyyMMdd-HHmmss>.xlsx

.PARAMETER DryRun
    (Optional) Alias: -WhatIf. Skips create/update/write calls but still
    attempts reads where possible to preview what would happen.

.PARAMETER ApiConfigPath
    (Optional) Path to a JSON file that overrides API catalog entries
    (endpoints, api-versions, request body templates). See end of script
    for the expected schema.

.EXAMPLE
    # Full subscription scan (interactive login)
    .\Invoke-AzResiliencyReport.ps1 -SubscriptionId "00000000-0000-0000-0000-000000000000" -ServiceGroupName "Contoso"

.EXAMPLE
    # Run with an explicit resource list and custom output path
    .\Invoke-AzResiliencyReport.ps1 `
        -SubscriptionId "00000000-0000-0000-0000-000000000000" `
        -ServiceGroupName "Contoso" `
        -ResourceListPath ".\my-resources.json" `
        -OutputPath ".\MyReport.xlsx"

.EXAMPLE
    # Run with custom API config overrides (preview endpoint changes, etc.)
    .\Invoke-AzResiliencyReport.ps1 `
        -SubscriptionId "00000000-0000-0000-0000-000000000000" `
        -ServiceGroupName "Contoso" `
        -ApiConfigPath ".\apiConfig.json" `
        -DryRun

.EXAMPLE
    # Sample apiConfig.json schema:
    # {
    #   "ApiCatalog": {
    #     "CreateServiceGroup": {
    #       "Method": "PUT",
    #       "PathTemplate": "/providers/Microsoft.Management/serviceGroups/{serviceGroupName}",
    #       "ApiVersion": "2024-02-01-preview",
    #       "Description": "Create or update a Service Group"
    #     }
    #   },
    # }

.NOTES
    Prerequisites:
      - Az.Accounts module (for Connect-AzAccount / Invoke-AzRestMethod)
      - ImportExcel module (Install-Module ImportExcel) for XLSX output.
        If ImportExcel is not available, falls back to CSV files.
    Author:  Azure Resiliency Automation
    Version: 1.0.0
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Azure subscription ID to target.")]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $false, HelpMessage = "Path to JSON/CSV file with explicit ARM resource IDs.")]
    [ValidateScript({ if ($_) { Test-Path $_ } else { $true } })]
    [string]$ResourceListPath,

    [Parameter(Mandatory = $false, HelpMessage = "Resource group name to scope resource discovery.")]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true, HelpMessage = "Name for the Service Group (groupId in the REST API).")]
    [string]$ServiceGroupName,

    [Parameter(Mandatory = $false, HelpMessage = "Azure region for the Service Group.")]
    [string]$ServiceGroupLocation = "eastus",

    [Parameter(Mandatory = $false, HelpMessage = "Output path for the Excel report.")]
    [string]$OutputPath,

    [Parameter(Mandatory = $false, HelpMessage = "Skip write operations; read-only preview mode.")]
    [switch]$DryRun,

    [Parameter(Mandatory = $false, HelpMessage = "Path to JSON config that overrides API catalog entries.")]
    [ValidateScript({ if ($_) { Test-Path $_ } else { $true } })]
    [string]$ApiConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ───────────────────────── Constants & Defaults ──────────────────────────

$script:LogEntries     = [System.Collections.Generic.List[object]]::new()
$script:ApiTraceList   = [System.Collections.Generic.List[object]]::new()
$script:StartTimestamp = Get-Date

if (-not $OutputPath) {
    $ts = $script:StartTimestamp.ToString('yyyyMMdd-HHmmss')
    $OutputPath = Join-Path $PWD "ResiliencyReport_${SubscriptionId}_${ts}.xlsx"
}

#endregion

#region ───────────────────────── API Catalog ───────────────────────────────────

# Central catalog mapping operation -> { Method, PathTemplate, ApiVersion, Description }.
# PathTemplate placeholders: {subscriptionId}, {serviceGroupName}, {resourceId}, {relationshipId}, {tenantId}
# Refs: https://learn.microsoft.com/azure/governance/service-groups/create-service-group-rest-api
#       https://learn.microsoft.com/azure/governance/service-groups/create-service-group-member-rest-api
# Any entry can be overridden by -ApiConfigPath without editing code.

$script:ApiCatalog = @{
    # ── Service Group management (Microsoft.Management) ──────────────────────
    CreateServiceGroup = @{
        Method       = 'PUT'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}'
        ApiVersion   = '2024-02-01-preview'
        Description  = 'Create or update a Service Group (async)'
    }
    GetServiceGroup = @{
        Method       = 'GET'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}'
        ApiVersion   = '2024-02-01-preview'
        Description  = 'Retrieve a Service Group'
    }
    DeleteServiceGroup = @{
        Method       = 'DELETE'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}'
        ApiVersion   = '2024-02-01-preview'
        Description  = 'Delete a Service Group'
    }
    # ── Membership via relationship resources (Microsoft.Relationships) ──────
    AddServiceGroupMember = @{
        Method       = 'PUT'
        PathTemplate = '{resourceId}/providers/Microsoft.Relationships/serviceGroupMember/{relationshipId}'
        ApiVersion   = '2023-09-01-preview'
        Description  = 'Add a resource to a Service Group via member relationship'
    }
    RemoveServiceGroupMember = @{
        Method       = 'DELETE'
        PathTemplate = '{resourceId}/providers/Microsoft.Relationships/serviceGroupMember/{relationshipId}'
        ApiVersion   = '2023-09-01-preview'
        Description  = 'Remove a resource member relationship from a Service Group'
    }
    # ── Goal Templates & Assignments (Microsoft.AzureResilienceManagement) ───
    CreateGoalTemplate = @{
        Method       = 'PUT'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.AzureResilienceManagement/goalTemplates/{goalTemplateName}'
        ApiVersion   = '2025-02-01-preview'
        Description  = 'Create or update a Goal Template on a Service Group'
    }
    GetGoalTemplate = @{
        Method       = 'GET'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.AzureResilienceManagement/goalTemplates/{goalTemplateName}'
        ApiVersion   = '2025-02-01-preview'
        Description  = 'Retrieve a Goal Template'
    }
    CreateGoalAssignment = @{
        Method       = 'PUT'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.AzureResilienceManagement/goalAssignments/{goalAssignmentName}'
        ApiVersion   = '2025-02-01-preview'
        Description  = 'Create or update a Goal Assignment on a Service Group'
    }
    GetGoalAssignment = @{
        Method       = 'GET'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.AzureResilienceManagement/goalAssignments/{goalAssignmentName}'
        ApiVersion   = '2025-02-01-preview'
        Description  = 'Retrieve a Goal Assignment'
    }
    # ── Resiliency APIs (preview, layered on Service Group) ──────────────────
    IncludeResource = @{
        Method       = 'POST'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.Resiliency/includeResource'
        ApiVersion   = '2024-10-01-preview'
        Description  = 'Include a resource in resiliency assessment'
    }
    ExcludeResource = @{
        Method       = 'POST'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.Resiliency/excludeResource'
        ApiVersion   = '2024-10-01-preview'
        Description  = 'Exclude a resource from resiliency assessment'
    }
    AttestResource = @{
        Method       = 'POST'
        PathTemplate = '/providers/Microsoft.Management/serviceGroups/{serviceGroupName}/providers/Microsoft.Resiliency/attestResource'
        ApiVersion   = '2024-10-01-preview'
        Description  = 'Attest a resource in resiliency assessment'
    }
    # ── Subscription / Resource Group resource queries ────────────────────────
    ListSubscriptionResources = @{
        Method       = 'GET'
        PathTemplate = '/subscriptions/{subscriptionId}/resources'
        ApiVersion   = '2024-03-01'
        Description  = 'List all resources in a subscription'
    }
    ListResourceGroupResources = @{
        Method       = 'GET'
        PathTemplate = '/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/resources'
        ApiVersion   = '2024-03-01'
        Description  = 'List all resources in a resource group'
    }
}

# Remediation guidance blurbs per resource type (keyed by ARM impactedField, case-insensitive match)
$script:RemediationBlurbs = @{
    'microsoft.containerregistry/registries' =
        'Zone redundancy is enabled by default in supported regions at no extra cost. No action needed for most registries. Premium SKU supports geo-replication for additional cross-region resilience.'
    'microsoft.containerservice/managedclusters' =
        'Existing non-zonal AKS clusters require redeployment with --zones specified. Spread node pools across 3 AZs, use Standard LB (default), configure pod topology spread constraints, and choose ZRS disks for stateful workloads. Cost impact: additional VMs across zones and ZRS disk pricing.'
    'microsoft.compute/virtualmachines' =
        'Regional VMs must be redeployed to a zonal configuration. Options: (1) Attach to VMSS Flex with zone spreading (lowest RTO, requires multi-zone instances + ZRS disks), or (2) Enable Azure Site Recovery zonal replication (higher RTO/RPO, limits regional DR). Cost impact: multi-zone deployment, ZRS disks, and/or ASR replication fees.'
    'microsoft.compute/virtualmachinescalesets' =
        'Multi-zone VMSS: enable app-health extension and automatic instance repair. Single-zone VMSS: redeploy across multiple zones. Use ZRS disks for shared storage. ASR support for VMSS Flex is limited to PowerShell only.'
    'microsoft.documentdb/databaseaccounts' =
        'Cannot enable AZ on an existing region in-place. Must add a temporary region, failover to it, remove the original region, then re-add it with AZ enabled. Brief write unavailability (seconds) during region changes. No performance or cost impact when using autoscale. Serverless accounts can only set AZ at creation.'
    'microsoft.network/azurefirewalls' =
        'Best configured at initial deployment. Existing firewalls (VNet-based only, not secured hubs) can be deallocated and reallocated with zones. Caveat: may receive a new private IP—update all UDRs referencing the old IP. All attached public IPs must share the same zone config.'
    'microsoft.dbforpostgresql/flexibleservers' =
        'Enable zone-redundant HA via portal or CLI to create a standby in a different AZ. Cost impact: standby server doubles compute cost. Cannot switch directly between SameZone and ZoneRedundant modes—must disable HA first, then re-enable with the desired mode.'
    'microsoft.network/publicipaddresses' =
        'Standard SKU IPs can be zone-redundant but the zone setting is immutable after creation—requires a new IP and re-association. Many regions are auto-migrating standard non-zonal IPs to zone-redundant. Basic SKU does not support zones; upgrade to Standard first.'
    'microsoft.sql/servers/databases' =
        'Premium/Business Critical/General Purpose tiers: online operation with a brief disconnect—ensure retry logic is in place. Hyperscale tier: zone redundancy can only be set at creation; existing DBs require redeployment via database copy, point-in-time restore, or geo-replica. Update connection strings after redeployment.'
    'microsoft.sql/servers/elasticpools' =
        'Enabling zone redundancy on an elastic pool makes all databases within it zone-redundant. Premium/BC/GP tiers: online toggle with brief disconnect. Ensure application retry logic is configured.'
    'microsoft.sql/managedinstances' =
        'Zone redundancy can be toggled on existing instances via portal/CLI/PowerShell. Prerequisite: backup storage redundancy must be set to zone- or geo-zone-redundant first. Apply backup change and wait for completion before enabling zone redundancy.'
    'microsoft.servicebus/namespaces' =
        'AZ support is automatically enabled for new namespaces in supported regions at no extra cost and cannot be disabled. Existing namespaces are being auto-migrated where possible. No action typically needed; verify the zoneRedundant property if in doubt.'
    'microsoft.storage/storageaccounts' =
        'Convert LRS to ZRS via customer-initiated conversion (no downtime, starts within 72 hrs). Cost increases with ZRS. Limitations: not supported for NFSv3/NFSv4.1 public endpoints, archive-tier blobs, or block blob premium accounts. For more control or unsupported scenarios, use manual migration (causes downtime).'
}

#endregion

#region ───────────────────────── Write-Log ─────────────────────────────────────

function Write-Log {
    <#
    .SYNOPSIS Structured logging to console and in-memory log list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $entry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString('o')
        Level     = $Level
        Message   = $Message
    }
    $script:LogEntries.Add($entry)
    $color = switch ($Level) {
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'DEBUG' { 'DarkGray' }
        default { 'Cyan' }
    }
    Write-Host "[$($entry.Timestamp)] [$Level] $Message" -ForegroundColor $color
}

#endregion

#region ───────────────────────── REST Helpers ──────────────────────────────────

function Wait-AsyncOperation {
    <#
    .SYNOPSIS Polls an Azure async operation URL until terminal state.
    .DESCRIPTION Service Group creation is asynchronous. The initial PUT returns
                 HTTP 202 with an Azure-AsyncOperation header URL. This function
                 polls that URL until the operation reaches a terminal state
                 (Succeeded, Failed, Canceled).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AsyncUrl,
        [string]$OperationName = 'AsyncPoll',
        [int]$MaxPollSeconds   = 300,
        [int]$PollIntervalSec  = 5
    )

    # Extract the ARM path from the full URL for Invoke-AzRestMethod
    $uri = [System.Uri]$AsyncUrl
    $pollPath = $uri.PathAndQuery

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $MaxPollSeconds) {
        Start-Sleep -Seconds $PollIntervalSec
        try {
            $resp = Invoke-AzRestMethod -Method GET -Path $pollPath
            if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300 -and $resp.Content) {
                $body = $resp.Content | ConvertFrom-Json -Depth 10
                $status = $body.status ?? $body.properties.provisioningState ?? ''
                Write-Log "  Async poll [$OperationName]: status=$status (elapsed $([int]$sw.Elapsed.TotalSeconds)s)" -Level DEBUG

                switch ($status) {
                    'Succeeded'  { Write-Log "Async operation $OperationName succeeded." -Level INFO;  return $body }
                    'Failed'     { Write-Log "Async operation $OperationName FAILED: $($body | ConvertTo-Json -Depth 5 -Compress)" -Level ERROR; return $body }
                    'Canceled'   { Write-Log "Async operation $OperationName was canceled." -Level WARN;  return $body }
                }
            } else {
                Write-Log "  Async poll [$OperationName]: HTTP $($resp.StatusCode)" -Level WARN
            }
        } catch {
            Write-Log "  Async poll [$OperationName] error: $($_.Exception.Message)" -Level WARN
        }
    }

    Write-Log "Async operation $OperationName did not complete within ${MaxPollSeconds}s." -Level WARN
    return $null
}

function Invoke-ResiliencyApi {
    <#
    .SYNOPSIS Central REST wrapper around Invoke-AzRestMethod with retries,
              correlation tracking, and error parsing. Returns parsed JSON.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OperationName,
        [hashtable]$PathParams  = @{},
        [object]$Body           = $null,
        [string]$QueryExtra     = '',  # additional query-string params
        [int]$MaxRetries        = 3,
        [switch]$SkipOnDryRun
    )

    if ($SkipOnDryRun -and $script:DryRunActive) {
        Write-Log "[DryRun] Skipping write operation: $OperationName" -Level WARN
        return $null
    }

    $catalogEntry = $script:ApiCatalog[$OperationName]
    if (-not $catalogEntry) {
        throw "Unknown API operation: '$OperationName'. Check ApiCatalog."
    }

    $method     = $catalogEntry.Method
    $apiVersion = $catalogEntry.ApiVersion

    # Resolve path template
    $path = $catalogEntry.PathTemplate
    foreach ($key in $PathParams.Keys) {
        $path = $path -replace "\{$key\}", $PathParams[$key]
    }
    $path = "${path}?api-version=${apiVersion}"
    if ($QueryExtra) {
        $path = "${path}&${QueryExtra}"
    }

    $payloadJson = $null
    if ($null -ne $Body) {
        $payloadJson = $Body | ConvertTo-Json -Depth 20 -Compress
    }

    $correlationId = [guid]::NewGuid().ToString()
    $traceEntry = [PSCustomObject]@{
        OperationName = $OperationName
        Method        = $method
        Path          = ($path -split '\?')[0]
        ApiVersion    = $apiVersion
        RequestBody   = if ($payloadJson) { ($payloadJson.Length -gt 2000) ? ($payloadJson.Substring(0, 2000) + '…[truncated]') : $payloadJson } else { '' }
        StatusCode    = $null
        Error         = ''
        CorrelationId = $correlationId
        Timestamp     = (Get-Date).ToString('o')
    }

    Write-Log "$method $($catalogEntry.Description) [$OperationName] -> $($path.Substring(0, [Math]::Min($path.Length, 120)))..." -Level DEBUG

    $attempt = 0
    $result  = $null
    while ($attempt -lt $MaxRetries) {
        $attempt++
        try {
            $invokeParams = @{
                Method = $method
                Path   = $path
            }
            if ($payloadJson) {
                $invokeParams['Payload'] = $payloadJson
            }

            $response = Invoke-AzRestMethod @invokeParams

            $traceEntry.StatusCode = $response.StatusCode

            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
                Write-Log "$OperationName completed (HTTP $($response.StatusCode))." -Level INFO

                # Handle async 202 Accepted (e.g. Service Group creation)
                if ($response.StatusCode -eq 202) {
                    $asyncUrl = $null
                    try {
                        if ($response.Headers -and $response.Headers.Contains('Azure-AsyncOperation')) {
                            $asyncUrl = ($response.Headers.GetValues('Azure-AsyncOperation'))[0]
                        }
                    } catch {
                        Write-Log "Could not parse Azure-AsyncOperation header: $($_.Exception.Message)" -Level DEBUG
                    }
                    if ($asyncUrl) {
                        Write-Log "Async operation detected for $OperationName. Polling..." -Level INFO
                        $asyncResult = Wait-AsyncOperation -AsyncUrl $asyncUrl -OperationName $OperationName
                        if ($asyncResult) { $result = $asyncResult }
                        break
                    }
                }

                if ($response.Content) {
                    $result = $response.Content | ConvertFrom-Json -Depth 20
                }
                break
            }

            if ($response.StatusCode -eq 429 -or $response.StatusCode -ge 500) {
                $retryAfter = 5 * $attempt
                # Try to parse Retry-After header if available
                if ($response.Headers -and $response.Headers['Retry-After']) {
                    $parsed = 0
                    if ([int]::TryParse($response.Headers['Retry-After'], [ref]$parsed)) {
                        $retryAfter = $parsed
                    }
                }
                Write-Log "$OperationName returned HTTP $($response.StatusCode). Retry $attempt/$MaxRetries in ${retryAfter}s..." -Level WARN
                Start-Sleep -Seconds $retryAfter
                continue
            }

            # Non-retryable error — parse and throw for caller to handle
            $errBody = $response.Content
            $errMsg  = "HTTP $($response.StatusCode)"
            if ($errBody) {
                try {
                    $errObj = $errBody | ConvertFrom-Json -Depth 10
                    if ($errObj.error.message) { $errMsg += ": $($errObj.error.code) - $($errObj.error.message)" }
                    elseif ($errObj.message)    { $errMsg += ": $($errObj.message)" }
                    else                        { $errMsg += ": $errBody" }
                } catch {
                    $errMsg += ": $errBody"
                }
            }
            $traceEntry.Error = $errMsg
            $script:ApiTraceList.Add($traceEntry)
            throw "$OperationName FAILED: $errMsg"

        } catch {
            # Re-throw non-retryable API errors immediately (thrown from the error-parsing block above)
            if ($_.Exception.Message -match '^.+ FAILED: HTTP \d+') {
                throw
            }
            $traceEntry.Error = $_.Exception.Message
            Write-Log "$OperationName exception (attempt $attempt/$MaxRetries): $($_.Exception.Message)" -Level ERROR
            if ($attempt -ge $MaxRetries) { break }
            Start-Sleep -Seconds (5 * $attempt)
        }
    }

    $script:ApiTraceList.Add($traceEntry)
    return $result
}

#endregion

#region ───────────────────────── Initialize-Context ───────────────────────────

function Initialize-Context {
    <#
    .SYNOPSIS Validate auth, set subscription context, check prerequisites.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Initializing context for subscription $SubscriptionId..."

    # Validate Az module
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        throw "Az.Accounts module is required. Install via: Install-Module Az.Accounts -Scope CurrentUser"
    }

    # Ensure logged in
    $ctx = Get-AzContext
    if (-not $ctx) {
        Write-Log "No Azure context found. Running Connect-AzAccount..."
        Connect-AzAccount -ErrorAction Stop | Out-Null
    }

    # Set subscription
    Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
    $ctx = Get-AzContext
    $script:TenantId = $ctx.Tenant.Id
    Write-Log "Authenticated as $($ctx.Account.Id) on tenant $($script:TenantId), subscription $($ctx.Subscription.Name) ($($ctx.Subscription.Id))."

    # Check and auto-install ImportExcel module
    $script:HasImportExcel = $null -ne (Get-Module -ListAvailable -Name ImportExcel)
    if (-not $script:HasImportExcel) {
        Write-Log "ImportExcel module not found. Attempting to install..." -Level WARN
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Import-Module ImportExcel -ErrorAction Stop
            $script:HasImportExcel = $true
            Write-Log "ImportExcel module installed successfully."
        } catch {
            Write-Log "Failed to install ImportExcel: $($_.Exception.Message). Will fall back to CSV output." -Level WARN
        }
    }

    # Load API config overrides if provided
    if ($ApiConfigPath) {
        Write-Log "Loading API config overrides from $ApiConfigPath..."
        $overrides = Get-Content -Path $ApiConfigPath -Raw | ConvertFrom-Json -Depth 10

        if ($overrides.ApiCatalog) {
            foreach ($prop in ($overrides.ApiCatalog | Get-Member -MemberType NoteProperty)) {
                $opName = $prop.Name
                $opOverride = $overrides.ApiCatalog.$opName
                if ($script:ApiCatalog.ContainsKey($opName)) {
                    if ($opOverride.Method)       { $script:ApiCatalog[$opName].Method       = $opOverride.Method }
                    if ($opOverride.PathTemplate)  { $script:ApiCatalog[$opName].PathTemplate = $opOverride.PathTemplate }
                    if ($opOverride.ApiVersion)    { $script:ApiCatalog[$opName].ApiVersion   = $opOverride.ApiVersion }
                    if ($opOverride.Description)   { $script:ApiCatalog[$opName].Description  = $opOverride.Description }
                    Write-Log "Overrode API catalog entry: $opName" -Level DEBUG
                } else {
                    # Add new catalog entry
                    $script:ApiCatalog[$opName] = @{
                        Method       = $opOverride.Method
                        PathTemplate = $opOverride.PathTemplate
                        ApiVersion   = $opOverride.ApiVersion
                        Description  = $opOverride.Description ?? $opName
                    }
                    Write-Log "Added new API catalog entry: $opName" -Level DEBUG
                }
            }
        }

    }

    $script:DryRunActive = $DryRun.IsPresent
    if ($script:DryRunActive) {
        Write-Log "*** DRY RUN MODE *** Write operations will be skipped." -Level WARN
    }
}

#endregion

#region ───────────────────────── Get-TargetResources ──────────────────────────

function Get-TargetResources {
    <#
    .SYNOPSIS Returns a list of ARM resource IDs from file or subscription query.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Discovering target resources..."

    $resources = @()

    if ($ResourceListPath) {
        Write-Log "Loading resources from file: $ResourceListPath"
        $ext = [System.IO.Path]::GetExtension($ResourceListPath).ToLower()

        if ($ext -eq '.json') {
            $raw = Get-Content -Path $ResourceListPath -Raw | ConvertFrom-Json -Depth 10
            if ($raw -is [System.Array]) {
                foreach ($item in $raw) {
                    if ($item -is [string]) {
                        $resources += [PSCustomObject]@{ ResourceId = $item }
                    } elseif ($item.resourceId) {
                        $resources += [PSCustomObject]@{ ResourceId = $item.resourceId }
                    } elseif ($item.ResourceId) {
                        $resources += [PSCustomObject]@{ ResourceId = $item.ResourceId }
                    }
                }
            }
        } elseif ($ext -eq '.csv') {
            $csv = Import-Csv -Path $ResourceListPath
            foreach ($row in $csv) {
                if ($row.ResourceId) {
                    $resources += [PSCustomObject]@{ ResourceId = $row.ResourceId }
                } elseif ($row.resourceId) {
                    $resources += [PSCustomObject]@{ ResourceId = $row.resourceId }
                }
            }
        } else {
            throw "Unsupported resource list file format: $ext. Use .json or .csv."
        }

        Write-Log "Loaded $($resources.Count) resources from file."
    } else {
        $operationName = 'ListSubscriptionResources'
        $pathParams = @{ subscriptionId = $SubscriptionId }

        if ($ResourceGroupName) {
            $operationName = 'ListResourceGroupResources'
            $pathParams['resourceGroupName'] = $ResourceGroupName
            Write-Log "Querying resource group '$ResourceGroupName' for resources..."
        } else {
            Write-Log "Querying subscription for all resources..."
        }

        $allResources = @()
        $nextLink = $null
        $firstCall = $true

        do {
            if ($firstCall) {
                $page = Invoke-ResiliencyApi -OperationName $operationName -PathParams $pathParams
                $firstCall = $false
            } else {
                # nextLink is a full URL; extract path + query
                $uri = [System.Uri]$nextLink
                $pathAndQuery = $uri.PathAndQuery
                $page = Invoke-ResiliencyApi -OperationName 'ListSubscriptionResources' -PathParams $pathParams -QueryExtra ''
                # For nextLink pagination, call Invoke-AzRestMethod directly
                $resp = Invoke-AzRestMethod -Method GET -Path $pathAndQuery
                if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300 -and $resp.Content) {
                    $page = $resp.Content | ConvertFrom-Json -Depth 10
                } else {
                    break
                }
            }

            if ($page -and (Get-Member -InputObject $page -Name 'value' -MemberType Properties)) {
                $allResources += $page.value
            }
            $nextLink = if ($page -and (Get-Member -InputObject $page -Name 'nextLink' -MemberType Properties)) { $page.nextLink } else { $null }
        } while ($nextLink)

        foreach ($r in $allResources) {
            $resources += [PSCustomObject]@{
                ResourceId   = $r.id
                ResourceName = $r.name
                ResourceType = $r.type
                Location     = $r.location
            }
        }
        Write-Log "Discovered $($resources.Count) resources in subscription."
    }

    if ($resources.Count -eq 0) {
        Write-Log "No resources found. Report will be empty." -Level WARN
    }

    return $resources
}

#endregion

#region ───────────────────────── Ensure-ServiceGroup ──────────────────────────

function Ensure-ServiceGroup {
    <#
    .SYNOPSIS Creates the Service Group if it does not exist; then adds resource
              members via Microsoft.Relationships/serviceGroupMember.
    .DESCRIPTION Uses the documented Service Group APIs:
      - PUT /providers/Microsoft.Management/serviceGroups/{groupId}  (async)
      - PUT {resourceId}/providers/Microsoft.Relationships/serviceGroupMember/{id}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Resources
    )

    $sgPathParams = @{ serviceGroupName = $ServiceGroupName }

    # ── Step 1: Check if Service Group already exists ────────────────────────
    Write-Log "Checking for existing Service Group '$ServiceGroupName'..."
    $existing = $null
    try {
        $existing = Invoke-ResiliencyApi -OperationName 'GetServiceGroup' -PathParams $sgPathParams
    } catch {
        Write-Log "Service Group '$ServiceGroupName' not found or not accessible: $($_.Exception.Message)" -Level DEBUG
    }

    $sgJustCreated = $false
    if ($existing -and $existing.id) {
        Write-Log "Service Group '$ServiceGroupName' already exists (id: $($existing.id))."
    } else {
        # ── Step 2: Create the Service Group ─────────────────────────────────
        Write-Log "Creating Service Group '$ServiceGroupName'..."
        $sgBody = @{
            properties = @{
                displayName = $ServiceGroupName
                parent      = @{
                    resourceId = "/providers/Microsoft.Management/serviceGroups/$($script:TenantId)"
                }
            }
        }

        $sgResult = Invoke-ResiliencyApi -OperationName 'CreateServiceGroup' `
            -PathParams $sgPathParams -Body $sgBody -SkipOnDryRun
        $sgJustCreated = $true
    }

    # ── Step 2b: Poll until Service Group is fully provisioned ───────────────
    if (-not $script:DryRunActive) {
        $maxWaitSec = 300
        $pollInterval = 15
        $elapsed = 0
        $provisioned = $false

        Write-Log "Waiting for Service Group '$ServiceGroupName' to reach provisioningState 'Succeeded'..."
        $authRetryDone = $false
        while ($elapsed -lt $maxWaitSec) {
            $existing = $null
            try {
                $existing = Invoke-ResiliencyApi -OperationName 'GetServiceGroup' -PathParams $sgPathParams
            } catch {
                $pollErr = $_.Exception.Message
                # Detect 403 authorization errors (RBAC propagation delay on new SG scope)
                if ($pollErr -match 'HTTP 403|does not have authorization|AuthorizationFailed') {
                    Write-Log "  GetServiceGroup returned 403 — RBAC for the new Service Group scope is still propagating (elapsed ${elapsed}s)..." -Level WARN
                    # Attempt a one-time token refresh to pick up the new scope permissions
                    if (-not $authRetryDone -and $elapsed -ge 30) {
                        $authRetryDone = $true
                        Write-Log "  Refreshing Azure credentials to pick up new scope permissions..." -Level INFO
                        try {
                            $ctx = Get-AzContext
                            Disconnect-AzAccount -AzureContext $ctx -ErrorAction SilentlyContinue | Out-Null
                            Connect-AzAccount -TenantId $script:TenantId -ErrorAction Stop | Out-Null
                            Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
                            Write-Log "  Credentials refreshed successfully." -Level INFO
                        } catch {
                            Write-Log "  Credential refresh failed: $($_.Exception.Message). Will continue polling." -Level WARN
                        }
                    }
                } else {
                    Write-Log "  Polling GetServiceGroup failed: $pollErr" -Level DEBUG
                }
            }
            if ($existing -and $existing.id) {
                $state = $null
                if ($existing.properties -and (Get-Member -InputObject $existing.properties -Name 'provisioningState' -MemberType Properties)) {
                    $state = $existing.properties.provisioningState
                }
                Write-Log "  Service Group provisioningState: $($state ?? 'unknown') (elapsed ${elapsed}s)" -Level DEBUG
                if ($state -eq 'Succeeded' -or ($null -eq $state -and $existing.id)) {
                    # If provisioningState is not returned but the resource exists, treat as ready
                    $provisioned = $true
                    Write-Log "Service Group '$ServiceGroupName' is ready."
                    break
                }
                if ($state -eq 'Failed' -or $state -eq 'Canceled') {
                    throw "Service Group '$ServiceGroupName' provisioning $state. Cannot proceed."
                }
            }
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval
        }

        if (-not $provisioned) {
            throw "Service Group '$ServiceGroupName' did not reach 'Succeeded' state within ${maxWaitSec}s. This may be caused by RBAC propagation delays on the new Service Group scope. Please wait a few minutes and re-run the script — it will detect the existing Service Group and continue."
        }
    }

    # Set the canonical Service Group ARM ID
    if ($existing -and $existing.id) {
        $script:ServiceGroupId = $existing.id
    } else {
        $script:ServiceGroupId = "/providers/Microsoft.Management/serviceGroups/$ServiceGroupName"
    }
    Write-Log "Service Group ARM ID: $($script:ServiceGroupId)"

    # ── Step 3: Add members via Microsoft.Relationships ──────────────────────
    $memberIds = @($Resources | ForEach-Object { $_.ResourceId } | Where-Object { $_ })

    if ($memberIds.Count -eq 0) {
        Write-Log "No resources to add as members." -Level WARN
        return $existing
    }

    # Wait for RBAC propagation on the newly created Service Group scope.
    # The Microsoft.Relationships RP performs a linked authorization check against
    # the Service Group. After creation, RBAC on the SG scope can take 1-5 minutes
    # to propagate, causing LinkedAuthorizationFailed (HTTP 403). The portal works
    # because the natural UI navigation delay allows propagation to complete.
    if ($sgJustCreated -and -not $script:DryRunActive) {
        $rbacWaitSec = 30
        Write-Log "Waiting ${rbacWaitSec}s for RBAC propagation on newly created Service Group..." -Level INFO
        Start-Sleep -Seconds $rbacWaitSec
    }

    Write-Log "Adding $($memberIds.Count) resource(s) as Service Group members..."

    # Deterministic relationship ID per service group so PUT is idempotent
    $relationshipId = "sgm-$ServiceGroupName"

    $memberBody = @{
        properties = @{
            targetId = "/providers/Microsoft.Management/serviceGroups/$ServiceGroupName"
        }
    }

    $added   = 0
    $skipped = 0
    $failed  = 0

    # Batch-level retry: if the first member hits LinkedAuthorizationFailed, all
    # subsequent members will too (same RBAC propagation issue). Wait once and
    # retry the entire batch instead of burning per-call retries on each member.
    $maxBatchRetries   = 4
    $batchAttempt      = 0
    $linkedAuthBlocked = $true

    while ($linkedAuthBlocked -and $batchAttempt -lt $maxBatchRetries) {
        $batchAttempt++
        $added   = 0
        $skipped = 0
        $failed  = 0
        $linkedAuthBlocked = $false
        $pendingIds = if ($batchAttempt -eq 1) { $memberIds } else { $failedIds }
        $failedIds  = @()

        foreach ($rid in $pendingIds) {
            $memberParams = @{
                resourceId     = $rid
                relationshipId = $relationshipId
            }
            try {
                $memberResult = Invoke-ResiliencyApi -OperationName 'AddServiceGroupMember' `
                    -PathParams $memberParams -Body $memberBody -SkipOnDryRun
                if ($null -ne $memberResult) { $added++ } else { $skipped++ }
            } catch {
                $errText = $_.Exception.Message
                if ($errText -match 'LinkedAuthorizationFailed') {
                    $linkedAuthBlocked = $true
                    $failedIds += $rid
                } else {
                    Write-Log "Failed to add member $rid : $errText" -Level WARN
                    $failed++
                }
            }
        }

        if ($linkedAuthBlocked -and $batchAttempt -lt $maxBatchRetries) {
            $waitSec = 30 * $batchAttempt
            Write-Log "LinkedAuthorizationFailed on $($failedIds.Count) member(s). RBAC still propagating. Batch retry $batchAttempt/$maxBatchRetries in ${waitSec}s..." -Level WARN
            Start-Sleep -Seconds $waitSec
        } elseif ($linkedAuthBlocked) {
            Write-Log "LinkedAuthorizationFailed persisted after $maxBatchRetries batch retries. $($failedIds.Count) member(s) could not be added. Verify you have 'Microsoft.Relationships/serviceGroupMember/write' on the Service Group scope." -Level ERROR
            $failed += $failedIds.Count
        }
    }

    Write-Log "Members: $added added, $skipped skipped (dry-run), $failed failed."

    return $existing
}

#endregion

#region ───────────────────────── Set-ZonalResiliencyGoals ─────────────────────

function Set-ZonalResiliencyGoals {
    <#
    .SYNOPSIS Creates a Goal Template and then a Goal Assignment on the Service Group.
    .DESCRIPTION Goal assignment requires a goal template to exist first.
      1. PUT .../goalTemplates/{goalTemplateName}   (creates the template)
      2. PUT .../goalAssignments/{goalAssignmentName} (assigns the template)
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    $goalTemplateName   = "defaultTemplate"
    $goalAssignmentName = "defaultTemplate-asg"

    $pathParams = @{
        serviceGroupName   = $ServiceGroupName
        goalTemplateName   = $goalTemplateName
        goalAssignmentName = $goalAssignmentName
    }

    # ── Step 1: Check if goal assignment already exists ──────────────────────
    if (-not $Force) {
        Write-Log "Checking if goal assignment '$goalAssignmentName' already exists..."
        $existingAssignment = $null
        try {
            $existingAssignment = Invoke-ResiliencyApi -OperationName 'GetGoalAssignment' -PathParams $pathParams
        } catch {
            Write-Log "Goal assignment not found: $($_.Exception.Message)" -Level DEBUG
        }
        if ($existingAssignment -and $existingAssignment.id) {
            Write-Log "Goal assignment already exists. Skipping (use -Force to reassign)." -Level INFO
            return $existingAssignment
        }
    }

    # ── Step 2: Create Goal Template ─────────────────────────────────────────
    Write-Log "Creating Goal Template '$goalTemplateName'..."

    $goalTemplateBody = @{
        properties = @{
            goalType = 'Resiliency'
        }
    }

    $templateResult = Invoke-ResiliencyApi -OperationName 'CreateGoalTemplate' `
        -PathParams $pathParams -Body $goalTemplateBody -SkipOnDryRun

    if (-not $script:DryRunActive) {
        # Verify the template was created
        Write-Log "Verifying Goal Template creation..."
        $existingTemplate = $null
        try {
            $existingTemplate = Invoke-ResiliencyApi -OperationName 'GetGoalTemplate' -PathParams $pathParams
        } catch {
            Write-Log "Goal Template GET failed: $($_.Exception.Message)" -Level DEBUG
        }
        if (-not ($existingTemplate -and $existingTemplate.id)) {
            Write-Log "Goal Template may still be provisioning. Continuing..." -Level WARN
        } else {
            Write-Log "Goal Template created: $($existingTemplate.id)"
        }
    }

    # ── Step 3: Create Goal Assignment (references the template) ─────────────
    Write-Log "Creating Goal Assignment '$goalAssignmentName'..."

    $goalTemplateId = "/providers/Microsoft.Management/serviceGroups/$ServiceGroupName/providers/Microsoft.AzureResilienceManagement/goalTemplates/$goalTemplateName"

    $goalAssignmentBody = @{
        properties = @{
            goalTemplateId     = $goalTemplateId
            goalAssignmentType = 'Resiliency'
        }
    }

    $assignmentResult = Invoke-ResiliencyApi -OperationName 'CreateGoalAssignment' `
        -PathParams $pathParams -Body $goalAssignmentBody -SkipOnDryRun

    # ── Step 4: Poll until Goal Assignment is fully provisioned ──────────────
    if (-not $script:DryRunActive) {
        $maxWaitSec = 300
        $pollInterval = 15
        $elapsed = 0
        $provisioned = $false

        Write-Log "Waiting for Goal Assignment '$goalAssignmentName' to reach provisioningState 'Succeeded'..."
        while ($elapsed -lt $maxWaitSec) {
            $gaResult = $null
            try {
                $gaResult = Invoke-ResiliencyApi -OperationName 'GetGoalAssignment' -PathParams $pathParams
            } catch {
                Write-Log "  Polling GetGoalAssignment failed: $($_.Exception.Message)" -Level DEBUG
            }
            if ($gaResult -and $gaResult.id) {
                $state = $null
                if ($gaResult.properties -and (Get-Member -InputObject $gaResult.properties -Name 'provisioningState' -MemberType Properties)) {
                    $state = $gaResult.properties.provisioningState
                }
                Write-Log "  Goal Assignment provisioningState: $($state ?? 'unknown') (elapsed ${elapsed}s)" -Level DEBUG
                if ($state -eq 'Succeeded' -or ($null -eq $state -and $gaResult.id)) {
                    $provisioned = $true
                    Write-Log "Goal Assignment '$goalAssignmentName' is ready."
                    break
                }
                if ($state -eq 'Failed' -or $state -eq 'Canceled') {
                    throw "Goal Assignment '$goalAssignmentName' provisioning $state. Cannot proceed."
                }
            }
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval
        }

        if (-not $provisioned) {
            throw "Goal Assignment '$goalAssignmentName' did not reach 'Succeeded' state within ${maxWaitSec}s. Aborting."
        }
    }

    Write-Log "Goal assignment completed."
    return $assignmentResult
}

#endregion

#region ───────────────────────── Get-ZonalResiliencySummary ────────────────────

function Get-ZonalResiliencySummary {
    <#
    .SYNOPSIS Fetches high-level zonal resiliency summary via Azure Resource Graph query.
    .DESCRIPTION Queries the ARG extensibilityresources table for
                 microsoft.azureresiliencemanagement/unifiedresilienceitems
                 scoped to the current Service Group.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Fetching Zonal Resiliency summary via Azure Resource Graph..."

    $argQuery = @"
extensibilityresources
| where type == "microsoft.azureresiliencemanagement/unifiedresilienceitems"
| where id contains "$ServiceGroupName"
| project props = parse_json(properties)
| project ZRResources = props.recommendations.highAvailability.enabledResourceCount, nonZRResources = props.recommendations.highAvailability.notEnabledResourceCount, NotEvaluatedResources = props.recommendations.highAvailability.notEvaluatedResourceCount
"@

    try {
        Write-Log "[ARG Summary] Query:`n$argQuery" -Level DEBUG
        Write-Log "[ARG Summary] Command: Search-AzGraph -Query <query> -UseTenantScope" -Level DEBUG
        $results = Search-AzGraph -Query $argQuery -UseTenantScope
        Write-Log "ARG summary query returned $($results.Count) row(s)."

        if ($results -and $results.Count -gt 0) {
            $row = $results[0]
            $summary = [PSCustomObject]@{
                ZRResources          = $row.ZRResources
                NonZRResources       = $row.nonZRResources
                NotEvaluatedResources = $row.NotEvaluatedResources
            }
            return $summary
        } else {
            Write-Log "ARG query returned no results for Service Group '$ServiceGroupName'." -Level WARN
            return $null
        }
    } catch {
        Write-Log "ARG query failed: $($_.Exception.Message)" -Level ERROR
        Write-Log "Ensure the Az.ResourceGraph module is installed (Install-Module Az.ResourceGraph)." -Level WARN
        return $null
    }
}

#endregion

#region ───────────────────────── Get-PerResourceZonalPosture ──────────────────

function Get-PerResourceZonalPosture {
    <#
    .SYNOPSIS Fetches per-resource zonal posture details via Azure Resource Graph query.
    .DESCRIPTION Queries the ARG extensibilityresources table for goalAssignment
                 goalResource items scoped to the current Service Group, joined
                 with business continuity data to determine zone resilience status.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Fetching per-resource zonal posture via Azure Resource Graph..."

    $sgNameLower = $ServiceGroupName.ToLower()

    $argQuery = @'
ExtensibilityResources
| where type contains "Microsoft.AzureResilienceManagement/goalAssignments/goalResource"
| extend goalResourceId = tolower(tostring(properties.resourceArmId))
| extend serviceGroupName = tolower(split(id, "/")[4]),
        subscriptionId = tolower(split(goalResourceId, "/")[2]),
        resourceGroup = tolower(split(goalResourceId, "/")[4]),
        resourceName = tostring(split(goalResourceId, "/")[-1])
| extend segments = split(goalResourceId, "/")
| extend segmentCount = array_length(segments)
| extend rawResourceType = case(
    segmentCount > 9, tolower(strcat(segments[6], "/", segments[7], "/", segments[8])),
    segmentCount > 8, tolower(strcat(segments[6], "/", segments[7])),
    segmentCount > 7, tolower(segments[6]),
    "unknown"
)
| extend resourceType = case(
    rawResourceType == "microsoft.compute/virtualmachines", "Microsoft.Compute/virtualMachines",
    rawResourceType == "microsoft.containerservice/managedclusters", "Microsoft.ContainerService/managedClusters",
    rawResourceType == "microsoft.network/applicationgateways", "Microsoft.Network/applicationGateways",
    rawResourceType == "microsoft.network/loadbalancers", "Microsoft.Network/loadBalancers",
    rawResourceType == "microsoft.sql/managedinstances", "microsoft.sql/managedinstances",
    rawResourceType == "microsoft.documentdb/databaseaccounts", "microsoft.documentdb/databaseaccounts",
    rawResourceType == "microsoft.dbforpostgresql/flexibleservers", "microsoft.dbforpostgresql/flexibleServers",
    rawResourceType == "microsoft.storage/storageaccounts/fileservices/shares", "microsoft.storage/storageaccounts/fileServices/shares",
    rawResourceType == "microsoft.storage/storageaccounts/blobservices/containers", "microsoft.storage/storageaccounts/blobServices/containers",
    rawResourceType == "microsoft.storage/storageaccounts", "microsoft.storage/storageaccounts",
    rawResourceType == "microsoft.servicebus/namespaces", "microsoft.servicebus/namespaces",
    rawResourceType == "microsoft.containerregistry/registries", "microsoft.containerregistry/registries",
    rawResourceType == "microsoft.network/publicipaddresses", "microsoft.network/publicipaddresses",
    rawResourceType == "microsoft.network/azurefirewalls", "microsoft.network/azureFirewalls",
    rawResourceType == "microsoft.network/virtualnetworkgateways", "microsoft.network/virtualnetworkgateways",
    rawResourceType == "microsoft.compute/virtualmachinescalesets", "microsoft.compute/virtualmachinescalesets",
    rawResourceType == "microsoft.dbformysql/flexibleservers", "microsoft.dbformysql/flexibleservers",
    rawResourceType == "microsoft.web/hostingenvironments", "microsoft.web/hostingenvironments",
    rawResourceType == "microsoft.web/serverfarms", "microsoft.web/serverfarms",
    rawResourceType == "microsoft.cache/redis", "microsoft.cache/redis",
    goalResourceId matches regex @"microsoft.sql/servers(/[^/]+)?/databases", "microsoft.sql/servers/databases",
    "notsupported"
)
| where serviceGroupName == "YOURSGNAME"
| where resourceType in~ ('Microsoft.Compute/virtualMachines', 'Microsoft.ContainerService/managedClusters', 'Microsoft.Network/applicationGateways', 'Microsoft.Network/loadBalancers', 'microsoft.sql/managedinstances', 'microsoft.sql/servers/databases', 'microsoft.documentdb/databaseaccounts', 'microsoft.storage/storageaccounts', 'microsoft.dbforpostgresql/flexibleServers', 'microsoft.servicebus/namespaces', 'microsoft.containerregistry/registries', 'microsoft.network/publicipaddresses', 'microsoft.network/azureFirewalls', 'microsoft.network/virtualnetworkgateways', 'microsoft.dbformysql/flexibleservers', 'microsoft.compute/virtualmachinescalesets', 'microsoft.web/hostingenvironments', 'microsoft.web/serverfarms', 'microsoft.cache/redis', 'notsupported')
| project id, serviceGroupName, goalResourceId, resourceType, resourceName, goalResourceType = type, properties, subscriptionId, resourceGroup, location
| join kind=leftouter (
    recoveryservicesresources
    | where type == 'microsoft.azurebusinesscontinuity/unifiedprotecteditems'
    | where isnotempty(properties.protectedItems)
    | mv-expand protectedItem = properties.protectedItems
    | extend protectedId = tolower(tostring(properties.linkedResourceInformation.linkedResourceId))
    | where tostring(protectedItem.isHighlyAvailable.propertyValue) == "true"
    | project protectedId, solutionDisplayName = tostring(protectedItem.solutionDisplayName), protectedType = type, protectedItem, highAvailabilityStatus = tostring(protectedItem.isHighlyAvailable.propertyValue), name
) on $left.goalResourceId == $right.protectedId
| sort by resourceName asc
| project id, goalResourceId, serviceGroupName, resourceType, resourceName, Status = iff(highAvailabilityStatus == "true" or parse_json(properties).highAvailabilityAttestationStatus=="Attested", "Zone Resilient", iff(parse_json(properties).highAvailabilityGoalParticipation == "Excluded", "Not Evaluated", "Not Zone Resilient"))
'@

    # Substitute the service group name into the query
    $argQuery = $argQuery -replace 'YOURSGNAME', $sgNameLower

    try {
        Write-Log "[ARG Posture] Query:`n$argQuery" -Level DEBUG
        Write-Log "[ARG Posture] Command: Search-AzGraph -Query <query> -First 1000 -UseTenantScope" -Level DEBUG
        $results = Search-AzGraph -Query $argQuery -First 1000 -UseTenantScope
        Write-Log "ARG posture query returned $($results.Count) row(s)."
        return $results
    } catch {
        Write-Log "ARG posture query failed: $($_.Exception.Message)" -Level ERROR
        Write-Log "Ensure the Az.ResourceGraph module is installed (Install-Module Az.ResourceGraph)." -Level WARN
        return $null
    }
}

#endregion

#region ───────────────────────── Get-ServiceGroupRecommendations ──────────────

function Get-ServiceGroupRecommendations {
    <#
    .SYNOPSIS Fetches recommendations for the Service Group via Azure Resource Graph.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Fetching recommendations via ARG query..."

    $argQuery = @'
advisorresources | where type == "microsoft.advisor/recommendations"
| where (id startswith "/providers/microsoft.management/servicegroups" and properties.lastUpdated >ago(2h)) or (id !startswith "/providers/microsoft.management/servicegroups" and properties.lastUpdated >ago(1d))
| union (advisorresources | where properties.extendedProperties.resiliencyExperience == "True" | extend idStr = trim(" ", tolower(tostring(properties.resourceMetadata.resourceId))) | join kind = inner ( ExtensibilityResources  | where type contains "Microsoft.AzureResilienceManagement/goalAssignments/goalResources" | project idStr = trim(" ", tolower(tostring(properties.resourceArmId))), sgId = id ) on idStr | extend id = sgId)
| where id startswith "/providers/Microsoft.Management/serviceGroups/YOURSGNAME/providers"
| join kind = leftouter (advisorresources | where type =~ "microsoft.advisor/suppressions" | extend tokens = split(id, "/") | extend name = iff(array_length(tokens) > 3, tokens[(array_length(tokens) - 3)], "") | extend expirationTimeStamp = todatetime(iff(strcmp(tostring(properties.ttl), "-1") == 0, "9999-12-31", properties.expirationTimeStamp)) | project suppressionId = tostring(properties.suppressionId), name, expirationTimeStamp) on name
| extend status = iff(isnull(expirationTimeStamp) or isempty(expirationTimeStamp), "Active", "Dismissed")
| extend extendedProperties = properties.extendedProperties
| extend recommendationSubcategory = tostring(extendedProperties.recommendationSubCategory)
| extend resourceType = tostring(properties.impactedField)
| extend resourceId = tolower(substring(id, 0, strlen(id) - 81))
| project recommendationTypeId = tostring(properties.recommendationTypeId), recommendationSubcategory, offeringId=tostring(extendedProperties.recommendationOfferingId), shortDescription=tostring(properties.shortDescription.problem), shortDescription_solution=tostring(properties.shortDescription.solution), category=tostring(properties.category), impact = iff(tostring(properties.impact)=="High", 0, iff(tostring(properties.impact)=="Medium", 1, 2)), impactedField=tostring(properties.impactedField), resourceType, resourceId, status, impactedResources=1, lastUpdated = format_datetime(todatetime(properties.lastUpdated), "M/d/yyyy h:mm tt"), extendedProperties = properties.extendedProperties, resourceMetadata = properties.resourceMetadata
| extend recommendationCostImplication = tostring(extendedProperties.recommendationCostImplication)
| summarize impactedResources = count(), lastUpdated = max(lastUpdated), recommendationCostImplication = any(recommendationCostImplication) by recommendationTypeId, recommendationSubcategory, offeringId, shortDescription, shortDescription_solution, status, category, impact, impactedField, resourceType, resourceId
| where status == "Active"
| sort by tostring(recommendationSubcategory) asc
| project resourceType, shortDescription, impactedResources
'@ -replace 'YOURSGNAME', $ServiceGroupName

    Write-Log "Recommendations ARG query:`n$argQuery" -Level DEBUG
    Write-Log "Running: Search-AzGraph -UseTenantScope -Query <recommendations-kql>" -Level DEBUG

    $results = Search-AzGraph -Query $argQuery -UseTenantScope

    if ($results -and $results.Count -gt 0) {
        Write-Log "ARG returned $($results.Count) recommendation(s)."
    } else {
        Write-Log "ARG returned no recommendations." -Level WARN
    }

    return $results
}

#endregion

#region ───────────────────────── Export-ResiliencyExcelReport ──────────────────

function Export-ResiliencyExcelReport {
    <#
    .SYNOPSIS Exports collected data to an Excel workbook (or CSV fallback).
    #>
    [CmdletBinding()]
    param(
        [object]$Summary,
        [object]$Posture,
        [object]$Recommendations,
        [Parameter(Mandatory)][array]$TargetResources
    )

    Write-Log "Preparing report data..."

    # ── Sheet 1: Overview ──
    $portalBase = 'https://portal.azure.com/?feature.canmodifystamps=true&Microsoft_Azure_Resources=stagepreview&feature.isDrillMonitoringFeatureEnabled=true&exp.ServiceGroupsResilienceJobs=true&exp.ServiceGroupsRecoveryPlan=true&feature.isROMockEnabled=false&feature.isrhubatscaleviewenabled=true&feature.isrhubrecommendationsenabled=true&feature.isRhubAtScaleDrillsViewEnabled=true&feature.customportal=false&exp.ServiceGroupsDrill=true&microsoft_azure_resiliencyHub=gatedpreview&microsoft_azure_bcdrcenter=develop&feature.isCostImplicationUIEnabled=true&exp.ServiceGroupResilience=true&isOPS360DevModeEnabled=true&isEnhancedScoreEnabled=true&feature.canarytraffic=true'

    $overviewData = @(
        [PSCustomObject]@{ Field = 'SubscriptionId';    Value = $SubscriptionId }
        [PSCustomObject]@{ Field = 'ServiceGroupName';  Value = $ServiceGroupName }
        [PSCustomObject]@{ Field = 'ServiceGroupId';    Value = $script:ServiceGroupId }
        [PSCustomObject]@{ Field = 'Location';          Value = $ServiceGroupLocation }
        [PSCustomObject]@{ Field = 'ReportGenerated';   Value = $script:StartTimestamp.ToString('o') }
        [PSCustomObject]@{ Field = 'ScriptVersion';     Value = '1.0.0' }
        [PSCustomObject]@{ Field = 'Mode';              Value = $(if ($script:DryRunActive) { 'DryRun' } else { 'Live' }) }
        [PSCustomObject]@{ Field = '';                   Value = '' }
        [PSCustomObject]@{ Field = '── Resiliency Posture ──'; Value = '' }
        [PSCustomObject]@{ Field = 'TotalResources';        Value = $TargetResources.Count }
        [PSCustomObject]@{ Field = 'ZonalResilient';        Value = $(if ($Summary) { $Summary.ZRResources } else { 'N/A' }) }
        [PSCustomObject]@{ Field = 'NonZonalResilient';     Value = $(if ($Summary) { $Summary.NonZRResources } else { 'N/A' }) }
        [PSCustomObject]@{ Field = 'NotEvaluated';          Value = $(if ($Summary) { $Summary.NotEvaluatedResources } else { 'N/A' }) }
        [PSCustomObject]@{ Field = '';                   Value = '' }
        [PSCustomObject]@{ Field = '── Explore More in Portal ──'; Value = '' }
        [PSCustomObject]@{ Field = 'Get the most current resiliency summary for this application'; Value = 'ResiliencySummary' }
        [PSCustomObject]@{ Field = 'Create and execute a zone-down drill for this application';    Value = 'ZoneDownDrill' }
        [PSCustomObject]@{ Field = 'View resiliency summary at scale across all applications';     Value = 'At-scale summary' }
    )

    # Build hyperlink map for Overview sheet (row index -> URL)
    # Rows are 1-based in Excel; data starts at row 2 (row 1 = header)
    $script:OverviewHyperlinks = @{}
    $linkRow = $overviewData.Count + 1  # not used directly; we search by Value text
    $resiliencySummaryUrl = "${portalBase}#view/Microsoft_Azure_Resources/ServiceGroup.MenuView/~/goalsAndRecommendations/serviceGroupId/$ServiceGroupName"
    $zoneDownDrillUrl     = "${portalBase}#view/Microsoft_Azure_Resources/ServiceGroup.MenuView/~/drills/serviceGroupId/$ServiceGroupName"
    $atScaleUrl           = "${portalBase}#view/Microsoft_Azure_BCDRCenter/AbcCenterMenuBlade/~/resiliencyOverview"
    $script:OverviewHyperlinks['ResiliencySummary'] = $resiliencySummaryUrl
    $script:OverviewHyperlinks['ZoneDownDrill']     = $zoneDownDrillUrl
    $script:OverviewHyperlinks['At-scale summary']  = $atScaleUrl

    # ── Sheet 2: Resources (from ARG posture query) ──
    $overrideBaseUrl = 'https://portal.azure.com/?feature.canmodifystamps=true&Microsoft_Azure_Resources=stagepreview&feature.isDrillMonitoringFeatureEnabled=true&exp.ServiceGroupsResilienceJobs=true&exp.ServiceGroupsRecoveryPlan=true&feature.isROMockEnabled=false&feature.isrhubatscaleviewenabled=true&feature.isrhubrecommendationsenabled=true&feature.isRhubAtScaleDrillsViewEnabled=true&feature.customportal=false&exp.ServiceGroupsDrill=true&microsoft_azure_resiliencyHub=gatedpreview&microsoft_azure_bcdrcenter=develop&feature.isCostImplicationUIEnabled=true&exp.ServiceGroupResilience=true&isOPS360DevModeEnabled=true&isEnhancedScoreEnabled=true&feature.canarytraffic=true#view/Microsoft_Azure_ResiliencyHub/HighAvailabilityStrategyDetails.ReactView/resourceDetails~/%7B%22totalResources%22%3A78%2C%22haResources%22%3A5%2C%22nonHaResources%22%3A15%2C%22notEvaluatedResources%22%3A58%7D/goalAssignmentName/defaultTemplate-asg/serviceGroupId/' + $ServiceGroupName

    $resourceData = @()
    if ($Posture -and $Posture.Count -gt 0) {
        foreach ($row in $Posture) {
            $rid = $row.id
            $status = $row.Status
            $overrideOptions = ''
            $overrideUrl = ''
            if ($status -eq 'Not Zone Resilient') {
                $overrideOptions = 'Exclude from evaluation (or) Mark as resilient'
                $overrideUrl = $overrideBaseUrl
            }

            $resourceData += [PSCustomObject]@{
                ResourceName    = $row.resourceName
                ResourceId      = $rid
                ResourceType    = $row.resourceType
                ServiceGroup    = $row.serviceGroupName
                Status          = $status
                OverrideOptions = $overrideOptions
                OverrideUrl     = $overrideUrl
            }
        }
    } else {
        # Fallback: list discovered resources without posture data
        foreach ($res in $TargetResources) {
            $rid  = $res.ResourceId
            $name = $res.ResourceName ?? ($rid -split '/')[-1]
            $type = $res.ResourceType ?? (($rid -split '/providers/')[-1] -replace '/[^/]+$', '')

            $resourceData += [PSCustomObject]@{
                ResourceName    = $name
                ResourceId      = $rid
                ResourceType    = $type
                ServiceGroup    = $ServiceGroupName
                Status          = 'Unknown'
                OverrideOptions = ''
                OverrideUrl     = ''
            }
        }
    }

    # ── Sheet 3: Recommendations ──
    $recommendationData = @()
    if ($Recommendations -and $Recommendations.Count -gt 0) {
        foreach ($rec in $Recommendations) {
            $rt = ($rec.resourceType ?? '').Trim()
            # Look up remediation blurb by case-insensitive resource type match
            $guidance = ''
            foreach ($key in $script:RemediationBlurbs.Keys) {
                if ($key -ieq $rt) {
                    $guidance = $script:RemediationBlurbs[$key]
                    break
                }
            }
            $recommendationData += [PSCustomObject]@{
                ResourceType          = $rt
                ShortDescription      = $rec.shortDescription ?? ''
                ImpactedResources     = $rec.impactedResources ?? 0
                RemediationGuidance   = $guidance
            }
        }
    }

    # ── Sheet 4: ApiTrace ──
    $traceData = $script:ApiTraceList | Select-Object OperationName, Method, Path, ApiVersion, RequestBody, StatusCode, Error, CorrelationId, Timestamp

    # ── Write Output ──
    if ($script:HasImportExcel) {
        Export-ExcelWorkbook -OverviewData $overviewData -ResourceData $resourceData `
            -RecommendationData $recommendationData -TraceData $traceData
    } else {
        Export-CsvFallback -OverviewData $overviewData -ResourceData $resourceData `
            -RecommendationData $recommendationData -TraceData $traceData
    }
}

function Export-ExcelWorkbook {
    [CmdletBinding()]
    param($OverviewData, $ResourceData, $RecommendationData, $TraceData)

    $xlFile = $OutputPath
    Write-Log "Generating Excel workbook: $xlFile"

    # Remove existing file to avoid append issues
    if (Test-Path $xlFile) { Remove-Item $xlFile -Force }

    $commonParams = @{
        AutoSize      = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        AutoFilter    = $true
    }

    # Overview sheet — export then add hyperlinks
    $pkg = $OverviewData | Export-Excel -Path $xlFile -WorksheetName 'Overview' `
        -AutoSize -FreezeTopRow -BoldTopRow -PassThru
    $wsOverview = $pkg.Workbook.Worksheets['Overview']

    # Find the Value column (column 2) and apply hyperlinks to link texts
    $valCol = 2
    for ($r = 2; $r -le $wsOverview.Dimension.Rows; $r++) {
        $cellText = $wsOverview.Cells[$r, $valCol].Text
        if ($cellText -and $script:OverviewHyperlinks.ContainsKey($cellText)) {
            $url = $script:OverviewHyperlinks[$cellText]
            $wsOverview.Cells[$r, $valCol].Hyperlink = [System.Uri]::new($url)
            $wsOverview.Cells[$r, $valCol].Style.Font.UnderLine = $true
            $wsOverview.Cells[$r, $valCol].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        }
    }
    $pkg.Save()
    $pkg.Dispose()

    # Resources sheet — export data, then apply conditional formatting and hyperlinks
    if ($ResourceData.Count -gt 0) {
        # Export without OverrideUrl column (it's only used for the hyperlink target)
        $displayData = $ResourceData | Select-Object ResourceName, ResourceId, ResourceType, ServiceGroup, Status, OverrideOptions
        $pkg = $displayData | Export-Excel -Path $xlFile -WorksheetName 'Resources' @commonParams -PassThru
        $ws = $pkg.Workbook.Worksheets['Resources']

        # Find the Status column index (1-based)
        $statusCol = $null
        $overrideCol = $null
        for ($c = 1; $c -le $ws.Dimension.Columns; $c++) {
            $header = $ws.Cells[1, $c].Text
            if ($header -eq 'Status') { $statusCol = $c }
            if ($header -eq 'OverrideOptions') { $overrideCol = $c }
        }

        if ($statusCol) {
            # Apply cell-by-cell formatting on the Status column to avoid substring overlap
            # ("Not Zone Resilient" contains "Zone Resilient" so ContainsText rules conflict)
            for ($r = 2; $r -le $ws.Dimension.Rows; $r++) {
                $cellValue = $ws.Cells[$r, $statusCol].Text
                $fillColor = $null
                switch ($cellValue) {
                    'Zone Resilient'     { $fillColor = [System.Drawing.Color]::FromArgb(198, 239, 206) } # Green
                    'Not Zone Resilient' { $fillColor = [System.Drawing.Color]::FromArgb(189, 215, 238) } # Blue
                    default              { $fillColor = [System.Drawing.Color]::FromArgb(255, 235, 156) } # Yellow
                }
                if ($fillColor) {
                    $ws.Cells[$r, $statusCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $ws.Cells[$r, $statusCol].Style.Fill.BackgroundColor.SetColor($fillColor)
                }
            }
        }

        # Add hyperlinks on OverrideOptions column
        if ($overrideCol) {
            for ($r = 2; $r -le $ws.Dimension.Rows; $r++) {
                $cellText = $ws.Cells[$r, $overrideCol].Text
                if ($cellText -and $cellText -ne '') {
                    # Find matching row in ResourceData (r-2 because row 1 is header, array is 0-based)
                    $idx = $r - 2
                    if ($idx -lt $ResourceData.Count) {
                        $url = $ResourceData[$idx].OverrideUrl
                        if ($url) {
                            $ws.Cells[$r, $overrideCol].Hyperlink = [System.Uri]::new($url)
                            $ws.Cells[$r, $overrideCol].Style.Font.UnderLine = $true
                            $ws.Cells[$r, $overrideCol].Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
                        }
                    }
                }
            }
        }

        $pkg.Save()
        $pkg.Dispose()
    } else {
        @([PSCustomObject]@{ Message = 'No resource data available.' }) |
            Export-Excel -Path $xlFile -WorksheetName 'Resources' -AutoSize
    }

    # Recommendations sheet
    if ($RecommendationData.Count -gt 0) {
        $RecommendationData | Export-Excel -Path $xlFile -WorksheetName 'Recommendations' @commonParams
    } else {
        @([PSCustomObject]@{ Message = 'No recommendations available.' }) |
            Export-Excel -Path $xlFile -WorksheetName 'Recommendations' -AutoSize
    }

    # ApiTrace sheet
    if ($TraceData.Count -gt 0) {
        $TraceData | Export-Excel -Path $xlFile -WorksheetName 'ApiTrace' @commonParams
    }

    Write-Log "Excel workbook saved to: $xlFile" -Level INFO
}

function Export-CsvFallback {
    [CmdletBinding()]
    param($OverviewData, $ResourceData, $RecommendationData, $TraceData)

    $csvDir = [System.IO.Path]::ChangeExtension($OutputPath, $null).TrimEnd('.')
    if (-not (Test-Path $csvDir)) { New-Item -ItemType Directory -Path $csvDir -Force | Out-Null }
    Write-Log "ImportExcel not available. Exporting CSVs to folder: $csvDir"

    $OverviewData        | Export-Csv -Path (Join-Path $csvDir 'Overview.csv') -NoTypeInformation -Encoding UTF8
    $ResourceData        | Export-Csv -Path (Join-Path $csvDir 'Resources.csv') -NoTypeInformation -Encoding UTF8
    $RecommendationData  | Export-Csv -Path (Join-Path $csvDir 'Recommendations.csv') -NoTypeInformation -Encoding UTF8
    $TraceData           | Export-Csv -Path (Join-Path $csvDir 'ApiTrace.csv') -NoTypeInformation -Encoding UTF8

    Write-Log "CSV files saved to: $csvDir" -Level INFO
}

#endregion

#region ───────────────────────── Main Orchestration ───────────────────────────

function Invoke-Main {
    [CmdletBinding()]
    param()

    try {
        Write-Log "═══════════════════════════════════════════════════════════════"
        Write-Log " Azure Resiliency Hub Report Generator v1.0.0"
        Write-Log "═══════════════════════════════════════════════════════════════"

        # Step 1: Initialize
        Initialize-Context

        # Step 2: Discover resources
        $targetResources = Get-TargetResources
        if ($targetResources.Count -eq 0) {
            Write-Log "No resources to process. Generating empty report." -Level WARN
        }

        # Step 3: Ensure Service Group exists with member resources
        $sg = Ensure-ServiceGroup -Resources $targetResources

        # Step 4: Assign Zonal Resiliency goals
        $goals = Set-ZonalResiliencyGoals

        # Step 4b: Wait for ARG data to become available after goal assignment
        $argMaxWaitSec = 300
        $argPollInterval = 30

        # Step 5a: Fetch summary (retry until non-zero results or timeout)
        $summary = $null
        $argElapsed = 0
        Write-Log "Fetching Zonal Resiliency summary (will retry up to ${argMaxWaitSec}s for data)..."
        while ($argElapsed -lt $argMaxWaitSec) {
            $summary = Get-ZonalResiliencySummary
            if ($null -ne $summary) {
                Write-Log "Summary data retrieved successfully."
                break
            }
            Write-Log "  Summary not yet available, retrying in ${argPollInterval}s... (elapsed ${argElapsed}s)" -Level DEBUG
            Start-Sleep -Seconds $argPollInterval
            $argElapsed += $argPollInterval
        }
        if ($null -eq $summary) {
            Write-Log "Summary data not available after ${argMaxWaitSec}s. Continuing with N/A values." -Level WARN
        }

        # Step 5b: Fetch per-resource posture (retry until non-zero results or timeout)
        $posture = $null
        $argElapsed = 0
        Write-Log "Fetching per-resource posture (will retry up to ${argMaxWaitSec}s for data)..."
        while ($argElapsed -lt $argMaxWaitSec) {
            $posture = Get-PerResourceZonalPosture
            if ($posture -and $posture.Count -gt 0) {
                Write-Log "Posture data retrieved: $($posture.Count) resource(s)."
                break
            }
            Write-Log "  Posture not yet available, retrying in ${argPollInterval}s... (elapsed ${argElapsed}s)" -Level DEBUG
            Start-Sleep -Seconds $argPollInterval
            $argElapsed += $argPollInterval
        }
        if (-not $posture -or $posture.Count -eq 0) {
            Write-Log "Posture data not available after ${argMaxWaitSec}s. Resources sheet will be based on discovered resources." -Level WARN
        }

        # Step 5c: Fetch recommendations (with retry)
        $recommendations = $null
        $argElapsed = 0
        Write-Log "Fetching recommendations (will retry up to ${argMaxWaitSec}s for data)..."
        while ($argElapsed -lt $argMaxWaitSec) {
            $recommendations = Get-ServiceGroupRecommendations
            if ($recommendations -and $recommendations.Count -gt 0) {
                Write-Log "Recommendations data retrieved: $($recommendations.Count) item(s)."
                break
            }
            Write-Log "  Recommendations not yet available, retrying in ${argPollInterval}s... (elapsed ${argElapsed}s)" -Level DEBUG
            Start-Sleep -Seconds $argPollInterval
            $argElapsed += $argPollInterval
        }
        if (-not $recommendations -or $recommendations.Count -eq 0) {
            Write-Log "Recommendations data not available after ${argMaxWaitSec}s." -Level WARN
        }

        # Step 6: Generate report
        Export-ResiliencyExcelReport -Summary $summary `
            -Posture $posture -Recommendations $recommendations -TargetResources $targetResources

        # Final summary
        Write-Log "═══════════════════════════════════════════════════════════════"
        Write-Log " REPORT GENERATION COMPLETE"
        Write-Log "═══════════════════════════════════════════════════════════════"
        Write-Log "___"
        Write-Log "Next steps:"
        Write-Log "  1. Open the report: $OutputPath"
        Write-Log "API trace: $($script:ApiTraceList.Count) operations recorded in the ApiTrace sheet."

    } catch {
        Write-Log "FATAL: $($_.Exception.Message)" -Level ERROR
        Write-Log "Stack: $($_.ScriptStackTrace)" -Level ERROR
        throw
    }
}

# Entry point
Invoke-Main

#endregion

<#
═══════════════════════════════════════════════════════════════════════════════
  SAMPLE apiConfig.json SCHEMA
═══════════════════════════════════════════════════════════════════════════════

{
  "ApiCatalog": {
    "CreateServiceGroup": {
      "Method": "PUT",
      "PathTemplate": "/providers/Microsoft.Management/serviceGroups/{serviceGroupName}",
      "ApiVersion": "2024-02-01-preview",
      "Description": "Create or update a Service Group (custom override)"
    },
    "AddServiceGroupMember": {
      "Method": "PUT",
      "PathTemplate": "{resourceId}/providers/Microsoft.Relationships/serviceGroupMember/{relationshipId}",
      "ApiVersion": "2023-09-01-preview"
    },
    "GetSummary": {
      "ApiVersion": "2025-01-01-preview"
    },
    "GetResourcesPosture": {
      "Method": "POST",
      "ApiVersion": "2025-01-01-preview"
    },
    "GetRecommendations": {
      "Method": "POST",
      "ApiVersion": "2025-01-01-preview"
    },
    "AssignGoals": {
      "ApiVersion": "2025-01-01-preview"
    }
  },
}

═══════════════════════════════════════════════════════════════════════════════
  SAMPLE resources.json
═══════════════════════════════════════════════════════════════════════════════

[
  "/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/myRG/providers/Microsoft.Compute/virtualMachines/vm1",
  "/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/myRG/providers/Microsoft.Sql/servers/sql1/databases/db1"
]

OR:

[
  { "resourceId": "/subscriptions/.../providers/Microsoft.Compute/virtualMachines/vm1" },
  { "resourceId": "/subscriptions/.../providers/Microsoft.Storage/storageAccounts/sa1" }
]

═══════════════════════════════════════════════════════════════════════════════
  SAMPLE resources.csv
═══════════════════════════════════════════════════════════════════════════════

ResourceId
/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/myRG/providers/Microsoft.Compute/virtualMachines/vm1
/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/myRG/providers/Microsoft.Sql/servers/sql1/databases/db1

#>
