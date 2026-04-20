#Requires -Modules Microsoft.Graph.Applications

<#
.SYNOPSIS
  Assigns Microsoft Graph and Exchange Online application permissions to an
  Azure Automation Managed Identity for the BambooHR-to-Entra-ID sync runbook.

.DESCRIPTION
  Connects to Microsoft Graph interactively, looks up the Managed Identity
  service principal, and assigns all required application role permissions.
  Idempotent — safely skips roles that are already assigned.

.NOTES
  Run this script once after creating the Azure Automation Account, or
  whenever new Graph permissions are needed.

  You must have Global Administrator or Privileged Role Administrator to
  grant application permissions.
#>

[CmdletBinding()]
param(
  [Parameter(HelpMessage = 'Azure Tenant ID. Leave empty to auto-detect after sign-in.')]
  [string] $TenantId,

  [Parameter(HelpMessage = 'Object ID of the Managed Identity (from Automation Account → Identity blade).')]
  [string] $MSIObjectId
)

# ─── Constants ──────────────────────────────────────────────────────────────
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$ExchangeAppRoleId = 'dc50a0fb-09a3-484d-be87-e023b12c6440'

# ─── Required Graph Application Permissions ─────────────────────────────────
$oAssignPermissions = @(
  # User management (read, create, update, disable, license, photo)
  'User.Read.All'
  'User.ReadWrite.All'
  'User-LifeCycleInfo.ReadWrite.All'
  'User.ReadBasic.All'
  'User.EnableDisableAccount.All'
  'UserAuthenticationMethod.ReadWrite.All'
  'User-LifeCycleInfo.ReadWrite.All'

  # Group management (membership, ownership, distribution lists)
  'Group.ReadWrite.All'
  'GroupMember.Read.All'
  'GroupMember.ReadWrite.All'

  # Directory (domains, service principals, org info)
  'Directory.ReadWrite.All'
  'Organization.Read.All'
  'Application.Read.All'

  # Mail (send notifications, set auto-replies)
  'Mail.ReadWrite'
  'Mail.ReadWrite.Shared'
  'Mail.Send'
  'Mail.Send.Shared'
  'MailboxSettings.ReadWrite'

  # Calendar (cancel meetings for terminated users)
  'Calendars.ReadWrite'

  # Device management (owned devices, Intune, Autopilot)
  'Device.ReadWrite.All'
  'DeviceManagementManagedDevices.ReadWrite.All'
  'DeviceManagementServiceConfig.ReadWrite.All'

  # OneDrive (grant manager access to terminated user's files)
  'Files.ReadWrite.All'
)

# ─── Connect to Graph ───────────────────────────────────────────────────────
$adminScopes = @(
  'Application.Read.All'
  'AppRoleAssignment.ReadWrite.All'
  'Directory.Read.All'
)

$connectParams = @{ Scopes = $adminScopes; NoWelcome = $true }
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
  $connectParams.TenantId = $TenantId
}

Write-Host 'Connecting to Microsoft Graph (browser sign-in may be required)...'
Connect-MgGraph @connectParams

$ctx = Get-MgContext
if (-not $ctx) {
  Write-Error 'Failed to connect to Microsoft Graph. Exiting.'
  return
}
if ([string]::IsNullOrWhiteSpace($TenantId)) {
  $TenantId = $ctx.TenantId
}
Write-Host "Connected to tenant: $TenantId"

# ─── Resolve Managed Identity ───────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($MSIObjectId)) {
  $MSIObjectId = Read-Host 'Enter the Object ID of the Managed Identity (Azure Automation → Identity blade)'
}
if ([string]::IsNullOrWhiteSpace($MSIObjectId)) {
  Write-Error 'MSIObjectId is required. Exiting.'
  return
}

try {
  $oMsi = Get-MgServicePrincipal -ServicePrincipalId $MSIObjectId -ErrorAction Stop
}
catch {
  Write-Error "Could not find Managed Identity with Object ID '$MSIObjectId'. Verify the ID in Azure Automation → Identity blade. Error: $($_.Exception.Message)"
  return
}
Write-Host "Managed Identity: $($oMsi.DisplayName) ($($oMsi.Id))"

# ─── Resolve service principals ─────────────────────────────────────────────
$oGraphSpn = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'" -ErrorAction Stop
$exchangeResourceId = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction Stop).Id

# ─── Assign Graph permissions ───────────────────────────────────────────────
$oAppRoles = $oGraphSpn.AppRoles |
  Where-Object { ($_.Value -in $oAssignPermissions) -and ($_.AllowedMemberTypes -contains 'Application') }

$assignedCount = 0
$skippedCount = 0

foreach ($AppRole in $oAppRoles) {
  $body = @{
    PrincipalId = $oMsi.Id
    ResourceId  = $oGraphSpn.Id
    AppRoleId   = $AppRole.Id
  }

  try {
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $oMsi.Id -BodyParameter $body -ErrorAction Stop | Out-Null
    Write-Host "  Assigned: $($AppRole.Value)"
    $assignedCount++
  }
  catch {
    if ($_.Exception.Message -like '*already exists*' -or $_.Exception.Message -like '*Permission being assigned already exists*') {
      Write-Host "  Already assigned: $($AppRole.Value)"
      $skippedCount++
    }
    else {
      Write-Warning "  Failed to assign $($AppRole.Value): $($_.Exception.Message)"
    }
  }
}

# ─── Assign Exchange Online permission ──────────────────────────────────────
try {
  New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $oMsi.Id -PrincipalId $oMsi.Id -AppRoleId $ExchangeAppRoleId -ResourceId $exchangeResourceId -ErrorAction Stop | Out-Null
  Write-Host '  Assigned: Exchange.ManageAsApp (Exchange Online)'
  $assignedCount++
}
catch {
  if ($_.Exception.Message -like '*already exists*') {
    Write-Host '  Already assigned: Exchange.ManageAsApp (Exchange Online)'
    $skippedCount++
  }
  else {
    Write-Warning "  Failed to assign Exchange permission: $($_.Exception.Message)"
  }
}

# ─── Summary ────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host "Done. $assignedCount new permission(s) assigned, $skippedCount already present."
Write-Host ''
Write-Host 'NOTE: If you use shared mailbox delegation, also assign the Managed Identity'
Write-Host 'the Exchange Administrator Entra role manually in the Azure portal:'
Write-Host '  Entra ID → Roles and Administrators → Exchange Administrator → Add assignments'
