[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$AutomationAccountName,

    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [securestring]$BambooHrApiKey,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$BhrCompanyName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Location,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string]$SubscriptionId,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F-]{36}$')]
    [string]$TenantId,

    [Parameter()]
    [string]$AdminEmailAddress,

    [Parameter()]
    [securestring]$BambooHrWebhookPrivateKey,

    [Parameter()]
    [ValidateRange(1, 3660)]
    [int]$DaysAhead,

    [Parameter()]
    [ValidateRange(0, 3650)]
    [int]$DaysToKeepAccountsAfterTermination,

    [Parameter()]
    [ValidateRange(1, 24)]
    [int]$DeltaSyncIntervalHours = 1,

    [Parameter()]
    [string]$DeltaSyncScheduleName = 'bhr-delta-sync-hourly',

    [Parameter()]
    [datetimeoffset]$DeltaSyncStartTime,

    [Parameter()]
    [bool]$EnableMobilePhoneSync = $false,

    [Parameter()]
    [bool]$ForceSharedMailboxPermissions = $false,

    [Parameter()]
    [string]$FullSyncScheduleName = 'bhr-full-sync-weekly',

    [Parameter()]
    [ValidateScript({
            foreach ($item in $_) {
                if ($item -notin @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')) {
                    throw "Unsupported weekday value '$($item)'."
                }
            }
            $true
        })]
    [string[]]$FullSyncWeekDays = @('Saturday'),

    [Parameter()]
    [datetimeoffset]$FullSyncStartTime,

    [Parameter()]
    [string]$HelpDeskEmailAddress,

    [Parameter()]
    [string]$LicenseId,

    [Parameter()]
    [ValidateRange(1, 366)]
    [int]$ModifiedWithinDays,

    [Parameter()]
    [string]$NotificationEmailAddress,

    [Parameter()]
    [bool]$PublicNetworkAccess = $true,

    [Parameter()]
    [string]$RuntimeEnvironmentName = 'PowerShell-7.4',

    [Parameter()]
    [string]$RuntimeModuleUpdateScheduleName = 'runtime-env-modules-weekly',

    [Parameter()]
    [datetimeoffset]$RuntimeModuleUpdateStartTime,

    [Parameter()]
    [ValidateScript({
            foreach ($item in $_) {
                if ($item -notin @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')) {
                    throw "Unsupported weekday value '$($item)'."
                }
            }
            $true
        })]
    [string[]]$RuntimeModuleUpdateWeekDays = @('Sunday'),

    [Parameter()]
    [string]$ScheduleTimeZone = 'UTC',

    [Parameter()]
    [hashtable]$Tags = @{},

    [Parameter()]
    [securestring]$TeamsCardUri,

    [Parameter()]
    [string]$TemplateFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'infra\main.bicep'),

    [Parameter()]
    [string]$UsageLocation,

    [Parameter()]
    [bool]$CreateWebhook = $false,

    [Parameter()]
    [bool]$CurrentOnly = $false,

    [Parameter()]
    [string]$CustomizationsJson,

    [Parameter()]
    [string]$CustomizationsJsonFilePath,

    [Parameter()]
    [string]$DefaultProfilePicPath,

    [Parameter()]
    [string]$EmailSignature,

    [Parameter()]
    [datetimeoffset]$WebhookExpiryTime = ([datetimeoffset]::UtcNow.AddYears(1)),

    [Parameter()]
    [string]$WebhookName = 'bamboohr-user-sync',

    [Parameter()]
    [string]$WelcomeLinksHtml,

    [Parameter()]
    [string]$WelcomeUserText
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AutomationApiVersion = '2024-10-23'

function ConvertTo-PlainText {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([string])]
    param(
        [Parameter()]
        [securestring]$SecureString
    )

    if ($null -eq $SecureString) {
        return $null
    }

    $bstr = [IntPtr]::Zero
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-NextHourlyStartTimeUtc {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([datetimeoffset])]
    param()

    $candidate = [datetimeoffset]::UtcNow.AddHours(1)
    return [datetimeoffset]::new(
        $candidate.Year,
        $candidate.Month,
        $candidate.Day,
        $candidate.Hour,
        0,
        0,
        [timespan]::Zero
    )
}

function Get-NextWeeklyStartTimeUtc {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([datetimeoffset])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$WeekDays,

        [Parameter(Mandatory)]
        [ValidateRange(0, 23)]
        [int]$Hour
    )

    $now = [datetimeoffset]::UtcNow
    $candidates = foreach ($weekDay in $WeekDays) {
        $targetDay = [System.DayOfWeek]::$weekDay
        $daysUntil = ([int]$targetDay - [int]$now.DayOfWeek + 7) % 7
        $candidateDate = [datetime]::UtcNow.Date.AddDays($daysUntil).AddHours($Hour)
        $candidate = [datetimeoffset]::new($candidateDate, [timespan]::Zero)
        if ($candidate -le $now) {
            $candidate = $candidate.AddDays(7)
        }

        $candidate
    }

    return $candidates | Sort-Object | Select-Object -First 1
}

function Get-CustomizationsJson {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$BoundParameters
    )

    if (-not [string]::IsNullOrWhiteSpace($CustomizationsJson) -and -not [string]::IsNullOrWhiteSpace($CustomizationsJsonFilePath)) {
        throw 'Specify CustomizationsJson or CustomizationsJsonFilePath, but not both.'
    }

    $customizationTable = @{}

    if (-not [string]::IsNullOrWhiteSpace($CustomizationsJsonFilePath)) {
        if (-not (Test-Path -LiteralPath $CustomizationsJsonFilePath -PathType Leaf)) {
            throw "Customizations JSON file was not found: $($CustomizationsJsonFilePath)"
        }

        $fileContent = Get-Content -LiteralPath $CustomizationsJsonFilePath -Raw
        if (-not [string]::IsNullOrWhiteSpace($fileContent)) {
            $fileValues = $fileContent | ConvertFrom-Json -Depth 10 -AsHashtable
            foreach ($key in $fileValues.Keys) {
                $customizationTable[$key] = $fileValues[$key]
            }
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($CustomizationsJson)) {
        $inlineValues = $CustomizationsJson | ConvertFrom-Json -Depth 10 -AsHashtable
        foreach ($key in $inlineValues.Keys) {
            $customizationTable[$key] = $inlineValues[$key]
        }
    }

    $directMappings = @(
        @{ Parameter = 'CurrentOnly'; Key = 'CurrentOnly'; Value = $CurrentOnly }
        @{ Parameter = 'DaysAhead'; Key = 'DaysAhead'; Value = $DaysAhead }
        @{ Parameter = 'DaysToKeepAccountsAfterTermination'; Key = 'DaysToKeepAccountsAfterTermination'; Value = $DaysToKeepAccountsAfterTermination }
        @{ Parameter = 'DefaultProfilePicPath'; Key = 'DefaultProfilePicPath'; Value = $DefaultProfilePicPath }
        @{ Parameter = 'EmailSignature'; Key = 'EmailSignature'; Value = $EmailSignature }
        @{ Parameter = 'EnableMobilePhoneSync'; Key = 'EnableMobilePhoneSync'; Value = $EnableMobilePhoneSync }
        @{ Parameter = 'ForceSharedMailboxPermissions'; Key = 'ForceSharedMailboxPermissions'; Value = $ForceSharedMailboxPermissions }
        @{ Parameter = 'ModifiedWithinDays'; Key = 'ModifiedWithinDays'; Value = $ModifiedWithinDays }
        @{ Parameter = 'UsageLocation'; Key = 'UsageLocation'; Value = $UsageLocation }
        @{ Parameter = 'WelcomeLinksHtml'; Key = 'WelcomeLinksHtml'; Value = $WelcomeLinksHtml }
        @{ Parameter = 'WelcomeUserText'; Key = 'WelcomeUserText'; Value = $WelcomeUserText }
    )

    foreach ($mapping in $directMappings) {
        if ($BoundParameters.ContainsKey($mapping.Parameter)) {
            $customizationTable[$mapping.Key] = $mapping.Value
        }
    }

    if ($customizationTable.Count -eq 0) {
        return ''
    }

    return $customizationTable | ConvertTo-Json -Depth 10 -Compress
}

function Get-RandomSecret {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([string])]
    param(
        [Parameter()]
        [ValidateRange(16, 128)]
        [int]$ByteCount = 32
    )

    $bytes = [byte[]]::new($ByteCount)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($bytes)
    $secret = [Convert]::ToBase64String($bytes)
    $secret = $secret.TrimEnd('=').Replace('+', '-').Replace('/', '_')

    return $secret
}

function Set-AutomationVariableValue {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'None')]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AutomationAccountName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ResourceGroupName,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter()]
        [bool]$Encrypted = $false
    )

    $existingVariable = Get-AzAutomationVariable -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $Name -ErrorAction SilentlyContinue

    if ($null -eq $existingVariable) {
        if ($PSCmdlet.ShouldProcess($Name, 'Create Automation variable')) {
            New-AzAutomationVariable -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $Name -Description $Description -Encrypted:$Encrypted -Value $Value -ErrorAction Stop | Out-Null
        }
        return
    }

    if ($PSCmdlet.ShouldProcess($Name, 'Update Automation variable')) {
        Set-AzAutomationVariable -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $Name -Description $Description -Encrypted:$Encrypted -Value $Value -ErrorAction Stop | Out-Null
    }
}

function Set-RunbookRuntimeEnvironment {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'None')]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AutomationAccountName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ResourceGroupName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RunbookName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RuntimeEnvironmentName,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F-]{36}$')]
        [string]$SubscriptionId
    )

    $path = "/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationAccountName)/runbooks/$($RunbookName)?api-version=$($script:AutomationApiVersion)"
    $payload = @{
        properties = @{
            runtimeEnvironment = $RuntimeEnvironmentName
            type = 'PowerShell'
        }
    } | ConvertTo-Json -Depth 5 -Compress

    if ($PSCmdlet.ShouldProcess($RunbookName, 'Associate the runbook with the configured runtime environment')) {
        Invoke-AzRestMethod -Method 'PATCH' -Path $path -Payload $payload -ErrorAction Stop | Out-Null
    }
}

function Register-RunbookSchedule {
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([void])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AutomationAccountName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ResourceGroupName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RunbookName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ScheduleName,

        [Parameter()]
        [hashtable]$Parameters = @{}
    )

    $existingRegistration = Get-AzAutomationScheduledRunbook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName $RunbookName -ScheduleName $ScheduleName -ErrorAction SilentlyContinue
    if ($null -ne $existingRegistration) {
        return
    }

    Register-AzAutomationScheduledRunbook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName $RunbookName -ScheduleName $ScheduleName -Parameters $Parameters -ErrorAction Stop | Out-Null
}

if (-not (Test-Path -LiteralPath $TemplateFilePath -PathType Leaf)) {
    throw "Template file was not found: $($TemplateFilePath)"
}

$weekDaySet = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')
foreach ($weekDay in ($FullSyncWeekDays + $RuntimeModuleUpdateWeekDays | Sort-Object -Unique)) {
    if ($weekDay -notin $weekDaySet) {
        throw "Unsupported weekday value '$($weekDay)'. Use full weekday names such as Monday or Saturday."
    }
}

if ($CreateWebhook -and -not $PublicNetworkAccess) {
    throw 'Webhook creation requires PublicNetworkAccess to remain enabled.'
}

$resolvedCustomizationsJson = Get-CustomizationsJson -BoundParameters $PSBoundParameters
$resolvedDeltaSyncStartTime = if ($PSBoundParameters.ContainsKey('DeltaSyncStartTime')) { $DeltaSyncStartTime } else { Get-NextHourlyStartTimeUtc }
$resolvedFullSyncStartTime = if ($PSBoundParameters.ContainsKey('FullSyncStartTime')) { $FullSyncStartTime } else { Get-NextWeeklyStartTimeUtc -WeekDays $FullSyncWeekDays -Hour 3 }
$resolvedRuntimeModuleUpdateStartTime = if ($PSBoundParameters.ContainsKey('RuntimeModuleUpdateStartTime')) { $RuntimeModuleUpdateStartTime } else { Get-NextWeeklyStartTimeUtc -WeekDays $RuntimeModuleUpdateWeekDays -Hour 4 }

$bambooHrApiKeyPlainText = ConvertTo-PlainText -SecureString $BambooHrApiKey
$teamsCardUriPlainText = ConvertTo-PlainText -SecureString $TeamsCardUri
$webhookPrivateKeyPlainText = ConvertTo-PlainText -SecureString $BambooHrWebhookPrivateKey
if ($CreateWebhook -and [string]::IsNullOrWhiteSpace($webhookPrivateKeyPlainText)) {
    $webhookPrivateKeyPlainText = Get-RandomSecret
}

$runbookDefinitions = @(
    @{
        Description = 'Primary BambooHR to Entra ID reconciliation runbook.'
        FileName = 'Start-BambooHRUserProvisioning.ps1'
        Name = 'Start-BambooHRUserProvisioning'
    }
    @{
        Description = 'Webhook-triggered targeted BambooHR reconciliation runbook.'
        FileName = 'Start-BambooHrWebhookSync.ps1'
        Name = 'Start-BambooHrWebhookSync'
    }
    @{
        Description = 'Runtime environment package maintenance runbook.'
        FileName = 'Update-AzureAutomationRuntimeEnvironmentPSModules.ps1'
        Name = 'Update-AzureAutomationRuntimeEnvironmentPSModules'
    }
)

$templateParameters = @{
    adminEmailAddress = $AdminEmailAddress
    automationAccountName = $AutomationAccountName
    bambooHrCompanyName = $BhrCompanyName
    companyName = $CompanyName
    customizationsJson = $resolvedCustomizationsJson
    deltaSyncIntervalHours = $DeltaSyncIntervalHours
    deltaSyncScheduleName = $DeltaSyncScheduleName
    deltaSyncStartTime = $resolvedDeltaSyncStartTime.ToString('o')
    fullSyncScheduleName = $FullSyncScheduleName
    fullSyncStartTime = $resolvedFullSyncStartTime.ToString('o')
    fullSyncWeekDays = $FullSyncWeekDays
    helpDeskEmailAddress = $HelpDeskEmailAddress
    licenseId = $LicenseId
    location = $Location
    notificationEmailAddress = $NotificationEmailAddress
    publicNetworkAccess = $PublicNetworkAccess
    runtimeEnvironmentName = $RuntimeEnvironmentName
    runtimeModuleUpdateScheduleName = $RuntimeModuleUpdateScheduleName
    runtimeModuleUpdateStartTime = $resolvedRuntimeModuleUpdateStartTime.ToString('o')
    runtimeModuleUpdateWeekDays = $RuntimeModuleUpdateWeekDays
    scheduleTimeZone = $ScheduleTimeZone
    tags = $Tags
    tenantId = $TenantId
}

if ($PSCmdlet.ShouldProcess("subscription $($SubscriptionId)", 'Connect to Azure context')) {
    $currentContext = Get-AzContext -ErrorAction SilentlyContinue
    if ($null -eq $currentContext -or $currentContext.Subscription.Id -ne $SubscriptionId -or $currentContext.Tenant.Id -ne $TenantId) {
        Connect-AzAccount -Subscription $SubscriptionId -Tenant $TenantId -ErrorAction Stop | Out-Null
    }
    else {
        Set-AzContext -Subscription $SubscriptionId -Tenant $TenantId -ErrorAction Stop | Out-Null
    }
}

if ($PSCmdlet.ShouldProcess($ResourceGroupName, 'Ensure the deployment resource group exists')) {
    $existingResourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
    if ($null -eq $existingResourceGroup) {
        New-AzResourceGroup -Name $ResourceGroupName -Location $Location -Tag $Tags -ErrorAction Stop | Out-Null
    }
}

$deploymentName = 'bhr-automation-' + (Get-Date -Format 'yyyyMMddHHmmss')
$deployment = $null
if ($PSCmdlet.ShouldProcess($AutomationAccountName, 'Deploy Azure Automation infrastructure')) {
    $deployment = New-AzResourceGroupDeployment -Name $deploymentName -ResourceGroupName $ResourceGroupName -TemplateFile $TemplateFilePath -TemplateParameterObject $templateParameters -ErrorAction Stop
}

if ($PSCmdlet.ShouldProcess($AutomationAccountName, 'Import, associate, and publish runbooks')) {
    foreach ($runbook in $runbookDefinitions) {
        $runbookPath = Join-Path -Path $PSScriptRoot -ChildPath $runbook.FileName
        if (-not (Test-Path -LiteralPath $runbookPath -PathType Leaf)) {
            throw "Runbook file was not found: $($runbookPath)"
        }

        Import-AzAutomationRunbook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $runbook.Name -Path $runbookPath -Description $runbook.Description -Type 'PowerShell' -LogProgress $true -LogVerbose $false -Force -ErrorAction Stop | Out-Null
        Set-RunbookRuntimeEnvironment -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName $runbook.Name -RuntimeEnvironmentName $RuntimeEnvironmentName -SubscriptionId $SubscriptionId
        Publish-AzAutomationRunbook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $runbook.Name -ErrorAction Stop | Out-Null
    }
}

if ($PSCmdlet.ShouldProcess($AutomationAccountName, 'Set encrypted Automation variables')) {
    Set-AutomationVariableValue -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name 'BambooHrApiKey' -Description 'BambooHR API key used by the runbooks.' -Encrypted $true -Value $bambooHrApiKeyPlainText

    if (-not [string]::IsNullOrWhiteSpace($teamsCardUriPlainText)) {
        Set-AutomationVariableValue -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name 'TeamsCardUri' -Description 'Teams webhook URI used for adaptive card summaries.' -Encrypted $true -Value $teamsCardUriPlainText
    }

    if (-not [string]::IsNullOrWhiteSpace($webhookPrivateKeyPlainText)) {
        Set-AutomationVariableValue -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name 'BambooHrWebhookPrivateKey' -Description 'Shared secret used to validate BambooHR webhook signatures.' -Encrypted $true -Value $webhookPrivateKeyPlainText
    }
}

if ($PSCmdlet.ShouldProcess($AutomationAccountName, 'Register recurring schedules')) {
    Register-RunbookSchedule -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName 'Start-BambooHRUserProvisioning' -ScheduleName $DeltaSyncScheduleName
    Register-RunbookSchedule -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName 'Start-BambooHRUserProvisioning' -ScheduleName $FullSyncScheduleName -Parameters @{ FullSync = $true }
    Register-RunbookSchedule -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -RunbookName 'Update-AzureAutomationRuntimeEnvironmentPSModules' -ScheduleName $RuntimeModuleUpdateScheduleName
}

$webhookUrl = $null
if ($CreateWebhook -and $PSCmdlet.ShouldProcess($WebhookName, 'Create Azure Automation webhook')) {
    $existingWebhook = Get-AzAutomationWebhook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $WebhookName -ErrorAction SilentlyContinue
    if ($null -ne $existingWebhook) {
        throw "A webhook named '$($WebhookName)' already exists. Choose a new webhook name or remove the existing webhook first."
    }

    $newWebhook = New-AzAutomationWebhook -AutomationAccountName $AutomationAccountName -ResourceGroupName $ResourceGroupName -Name $WebhookName -RunbookName 'Start-BambooHrWebhookSync' -ExpiryTime $WebhookExpiryTime.UtcDateTime -IsEnabled $true -ErrorAction Stop
    $webhookUrl = $newWebhook.WebhookURI
}

[pscustomobject]@{
    AutomationAccountId = if ($null -ne $deployment -and $null -ne $deployment.Outputs.automationAccountId) { $deployment.Outputs.automationAccountId.Value } else { $null }
    AutomationAccountName = $AutomationAccountName
    ResourceGroupName = $ResourceGroupName
    RuntimeEnvironmentName = $RuntimeEnvironmentName
    RuntimeEnvironmentId = if ($null -ne $deployment -and $null -ne $deployment.Outputs.runtimeEnvironmentId) { $deployment.Outputs.runtimeEnvironmentId.Value } else { $null }
    WebhookName = if ($CreateWebhook) { $WebhookName } else { $null }
    WebhookSigningSecret = if ($CreateWebhook) { $webhookPrivateKeyPlainText } else { $null }
    WebhookUrl = $webhookUrl
}
