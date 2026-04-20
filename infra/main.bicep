targetScope = 'resourceGroup'

@description('Name of the Azure Automation account.')
@minLength(6)
param automationAccountName string

@description('Basic SKU is the only supported Azure Automation account tier.')
@allowed([
  'Basic'
])
param automationAccountSku string = 'Basic'

@description('Admin notification email address. Leave empty to let the runbook derive a default from CompanyName.')
param adminEmailAddress string = ''

@description('BambooHR company subdomain used for API requests.')
@minLength(1)
param bambooHrCompanyName string

@description('Optional JSON overrides stored in the BHR_CustomizationsJson Automation variable.')
param customizationsJson string = ''

@description('Friendly company name used when deriving default email addresses and notifications.')
@minLength(1)
param companyName string

@description('Name of the hourly delta sync schedule.')
param deltaSyncScheduleName string = 'bhr-delta-sync-hourly'

@description('Interval, in hours, for the delta sync schedule.')
@minValue(1)
@maxValue(24)
param deltaSyncIntervalHours int = 1

@description('ISO 8601 start time for the hourly delta sync schedule.')
param deltaSyncStartTime string

@description('Azure CLI default package version for the PowerShell 7.4 runtime environment.')
param defaultAzureCliPackageVersion string = '2.64.0'

@description('Az PowerShell default package version for the PowerShell 7.4 runtime environment.')
param defaultAzPackageVersion string = '12.3.0'

@description('Weekdays for the weekly full sync schedule.')
param fullSyncWeekDays array = [
  'Saturday'
]

@description('Name of the weekly full sync schedule.')
param fullSyncScheduleName string = 'bhr-full-sync-weekly'

@description('ISO 8601 start time for the weekly full sync schedule.')
param fullSyncStartTime string

@description('Primary help desk email address surfaced in onboarding and notification flows.')
param helpDeskEmailAddress string = ''

@description('Azure region for the Automation account and runtime environment.')
param location string = resourceGroup().location

@description('Optional Entra license SKU GUID assigned to newly created users.')
param licenseId string = ''

@description('Notification email address for operational summaries. Leave empty to let the runbook derive a default from CompanyName.')
param notificationEmailAddress string = ''

@description('Whether the non-ARM endpoint surface (webhooks and agents) is reachable from the public internet.')
param publicNetworkAccess bool = true

@description('Name of the PowerShell 7.4 runtime environment.')
param runtimeEnvironmentName string = 'PowerShell-7.4'

@description('Name of the weekly runtime package maintenance schedule.')
param runtimeModuleUpdateScheduleName string = 'runtime-env-modules-weekly'

@description('ISO 8601 start time for the weekly runtime package maintenance schedule.')
param runtimeModuleUpdateStartTime string

@description('Weekdays for the weekly runtime package maintenance schedule.')
param runtimeModuleUpdateWeekDays array = [
  'Sunday'
]

@description('Time zone for all schedules. Use UTC unless you have an explicit regional requirement.')
param scheduleTimeZone string = 'UTC'

@description('Tags applied to all tagged resources.')
param tags object = {}

@description('Entra tenant ID used by the runbooks.')
@minLength(1)
param tenantId string

var recommendedRuntimePackages = [
  {
    name: 'ExchangeOnlineManagement'
    version: '3.9.2'
  }
  {
    name: 'Microsoft.Graph.Authentication'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Applications'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Calendar'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.DeviceManagement'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.DeviceManagement.Enrollment'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Files'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Groups'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Identity.DirectoryManagement'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Identity.SignIns'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Mail'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Users'
    version: '2.36.1'
  }
  {
    name: 'Microsoft.Graph.Users.Actions'
    version: '2.36.1'
  }
  {
    name: 'PackageManagement'
    version: '1.4.8.1'
  }
  {
    name: 'PowerShellGet'
    version: '2.2.5'
  }
]

var automationVariables = [
  {
    name: 'AdminEmailAddress'
    description: 'Administrative contact used by the runbooks.'
    value: adminEmailAddress
  }
  {
    name: 'AutomationAccountName'
    description: 'Automation account name used by helper runbooks.'
    value: automationAccountName
  }
  {
    name: 'BHRCompanyName'
    description: 'BambooHR company subdomain used for API requests.'
    value: bambooHrCompanyName
  }
  {
    name: 'BHR_CustomizationsJson'
    description: 'JSON-based optional configuration overrides for the runbooks.'
    value: customizationsJson
  }
  {
    name: 'CompanyName'
    description: 'Friendly company name used in provisioning messages.'
    value: companyName
  }
  {
    name: 'HelpDeskEmailAddress'
    description: 'Help desk contact used in onboarding messages.'
    value: helpDeskEmailAddress
  }
  {
    name: 'LicenseId'
    description: 'Optional Entra license SKU GUID assigned to new users.'
    value: licenseId
  }
  {
    name: 'NotificationEmailAddress'
    description: 'Operational notification email address.'
    value: notificationEmailAddress
  }
  {
    name: 'ResourceGroupName'
    description: 'Resource group that contains this Automation account.'
    value: resourceGroup().name
  }
  {
    name: 'RuntimeEnvironment'
    description: 'Runtime environment name used by helper runbooks.'
    value: runtimeEnvironmentName
  }
  {
    name: 'SubscriptionID'
    description: 'Subscription that contains this Automation account.'
    value: subscription().subscriptionId
  }
  {
    name: 'TenantId'
    description: 'Entra tenant ID used by the runbooks.'
    value: tenantId
  }
]

resource automationAccount 'Microsoft.Automation/automationAccounts@2024-10-23' = {
  name: automationAccountName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  tags: tags
  properties: {
    disableLocalAuth: true
    publicNetworkAccess: publicNetworkAccess
    sku: {
      name: automationAccountSku
    }
  }
}

resource runtimeEnvironment 'Microsoft.Automation/automationAccounts/runtimeEnvironments@2024-10-23' = {
  parent: automationAccount
  name: runtimeEnvironmentName
  location: location
  tags: tags
  properties: {
    description: 'PowerShell 7.4 runtime environment for BambooHR to Entra ID user provisioning.'
    defaultPackages: {
      Az: defaultAzPackageVersion
      AzureCLI: defaultAzureCliPackageVersion
    }
    runtime: {
      language: 'PowerShell'
      version: '7.4'
    }
  }
}

resource runtimeEnvironmentPackages 'Microsoft.Automation/automationAccounts/runtimeEnvironments/packages@2024-10-23' = [for package in recommendedRuntimePackages: {
  parent: runtimeEnvironment
  name: package.name
  properties: {
    contentLink: {
      uri: 'https://www.powershellgallery.com/api/v2/package/${package.name}/${package.version}'
      version: package.version
    }
  }
}]

resource automationAccountVariables 'Microsoft.Automation/automationAccounts/variables@2024-10-23' = [for item in automationVariables: {
  parent: automationAccount
  name: item.name
  properties: {
    description: item.description
    isEncrypted: false
    value: item.value
  }
}]

resource deltaSyncSchedule 'Microsoft.Automation/automationAccounts/schedules@2024-10-23' = {
  parent: automationAccount
  name: deltaSyncScheduleName
  properties: {
    description: 'Runs the standard delta reconciliation runbook on an hourly cadence.'
    frequency: 'Hour'
    interval: deltaSyncIntervalHours
    startTime: deltaSyncStartTime
    timeZone: scheduleTimeZone
  }
}

resource fullSyncSchedule 'Microsoft.Automation/automationAccounts/schedules@2024-10-23' = {
  parent: automationAccount
  name: fullSyncScheduleName
  properties: {
    advancedSchedule: {
      weekDays: fullSyncWeekDays
    }
    description: 'Runs the weekly full reconciliation runbook.'
    frequency: 'Week'
    interval: 1
    startTime: fullSyncStartTime
    timeZone: scheduleTimeZone
  }
}

resource runtimeModuleUpdateSchedule 'Microsoft.Automation/automationAccounts/schedules@2024-10-23' = {
  parent: automationAccount
  name: runtimeModuleUpdateScheduleName
  properties: {
    advancedSchedule: {
      weekDays: runtimeModuleUpdateWeekDays
    }
    description: 'Runs the runtime environment package maintenance runbook.'
    frequency: 'Week'
    interval: 1
    startTime: runtimeModuleUpdateStartTime
    timeZone: scheduleTimeZone
  }
}

output automationAccountId string = automationAccount.id
output automationAccountPrincipalId string = automationAccount.identity.principalId
output deltaSyncScheduleResourceId string = deltaSyncSchedule.id
output runtimeEnvironmentId string = runtimeEnvironment.id
output runtimeEnvironmentPackageNames array = [for package in recommendedRuntimePackages: package.name]
