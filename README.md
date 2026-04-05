# Azure Resiliency Report Generator

A PowerShell script that automates Azure Resiliency assessments by creating a Service Group, assigning Zonal Resiliency goals, and producing an Excel report with per-resource posture, recommendations, and portal deep links.

## What the Script Does

`Invoke-AzResiliencyReport.ps1` performs the following steps end-to-end:

1. **Discovers resources** in a target subscription (or from an explicit JSON/CSV resource list, optionally scoped to a resource group).
2. **Creates or updates a Service Group** via the Azure Management REST API and adds discovered resources as members.
3. **Assigns Zonal Resiliency goals** by creating a goal template and goal assignment on the Service Group.
4. **Fetches zonal resiliency data** using Azure Resource Graph queries:
   - **Summary** — aggregate counts of zone-resilient, non-zone-resilient, and not-evaluated resources.
   - **Per-resource posture** — zone resilience status for each resource, joined with recovery services data.
   - **Recommendations** — active Advisor recommendations scoped to the Service Group, summarized by resource type.
5. **Generates an Excel workbook** with four sheets:
   - **Overview** — subscription, service group, posture counts, and clickable portal links for deeper exploration.
   - **Resources** — per-resource zone resilience status with conditional color coding (green = Zone Resilient, blue = Not Zone Resilient, yellow = other) and override option hyperlinks.
   - **Recommendations** — resource type, description, impacted resource count, and remediation guidance blurbs covering caveats like redeployment requirements, cost impact, and dependencies.
   - **ApiTrace** — full audit log of every REST API call made during the run.

If the ImportExcel module is unavailable (and auto-install fails), the script falls back to CSV output.

## Prerequisites

### Required (must be installed before running)

| Module | Purpose | Install Command |
|---|---|---|
| **Az.Accounts** | Authentication and `Invoke-AzRestMethod` | `Install-Module Az.Accounts -Scope CurrentUser` |
| **Az.ResourceGraph** | `Search-AzGraph` for posture, summary, and recommendations queries | `Install-Module Az.ResourceGraph -Scope CurrentUser` |

You must also be able to authenticate to Azure (`Connect-AzAccount`). If no Azure context is found, the script will invoke `Connect-AzAccount` interactively.

### Auto-installed

| Module | Purpose | Notes |
|---|---|---|
| **ImportExcel** | Excel workbook generation with formatting and hyperlinks | Automatically installed at runtime if missing. Falls back to CSV if installation fails. |

## Quick Start

```powershell
# 1. Install prerequisites (one-time)
Install-Module Az.Accounts -Scope CurrentUser
Install-Module Az.ResourceGraph -Scope CurrentUser

# 2. Run a scan for an application
.\Invoke-AzResiliencyReport.ps1 `
    -SubscriptionId "00000000-0000-0000-0000-000000000000" `
    -ServiceGroupName "MyApp" `
    -ResourceGroupName "myRG"
```

**Note** - Scopes of upto 500 resources are supported currently

The script will prompt for Azure login if needed, then generate a report at `.\ResiliencyReport_<SubscriptionId>_<timestamp>.xlsx`.

## Usage Examples

```powershell
# Scope to a specific resource group
.\Invoke-AzResiliencyReport.ps1 `
    -SubscriptionId "00000000-0000-0000-0000-000000000000" `
    -ServiceGroupName "MyApp" `
    -ResourceGroupName "myRG"

# Use an explicit resource list and custom output path
.\Invoke-AzResiliencyReport.ps1 `
    -SubscriptionId "00000000-0000-0000-0000-000000000000" `
    -ServiceGroupName "MyApp" `
    -ResourceListPath ".\my-resources.json" `
    -OutputPath ".\MyReport.xlsx"

# Dry-run mode (read-only, skips write operations)
.\Invoke-AzResiliencyReport.ps1 `
    -SubscriptionId "00000000-0000-0000-0000-000000000000" `
    -ServiceGroupName "MyApp" `
    -DryRun

# Override API endpoints via config file
.\Invoke-AzResiliencyReport.ps1 `
    -SubscriptionId "00000000-0000-0000-0000-000000000000" `
    -ServiceGroupName "MyApp" `
    -ApiConfigPath ".\apiConfig.json"
```

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-SubscriptionId` | Yes | — | Azure subscription ID (GUID format) |
| `-ServiceGroupName` | Yes | — | Name for the Service Group |
| `-ResourceGroupName` | No | — | Scope discovery to a specific resource group |
| `-ResourceListPath` | No | — | Path to a JSON or CSV file with ARM resource IDs |
| `-ServiceGroupLocation` | No | `eastus` | Azure region for the Service Group |
| `-OutputPath` | No | Auto-generated | Path for the output Excel file |
| `-DryRun` | No | `$false` | Skip write operations; read-only preview mode |
| `-ApiConfigPath` | No | — | JSON file to override API catalog entries (see `apiConfig.sample.json`) |
