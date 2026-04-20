<#
.SYNOPSIS
    Azure Automation runbook to update PowerShell modules for a Runtime Environment.
.DESCRIPTION
    Installs or updates PowerShell modules from PSGallery into the specified
    Azure Automation Runtime Environment via REST API.

    In non-interactive (default/runbook) mode, updates a hardcoded list of modules
    using Managed Identity authentication.

    In interactive mode (-Interactive $true), signs in via browser, lets you choose
    the subscription, automation account, and runtime environment, then compares
    installed module versions against PSGallery and prompts for selective or bulk
    updates.
.PARAMETER Interactive
    When $true, enables interactive mode with browser login, discovery prompts,
    version comparison, and per-module update approval. Default is $false for
    unattended runbook execution.
.PARAMETER SubscriptionID
    Azure subscription ID. In runbook mode, defaults to the SubscriptionID
    Automation Variable. In interactive mode, discovered via prompt if not provided.
.PARAMETER ResourceGroupName
    Resource group containing the Automation Account. In runbook mode, defaults to
    the ResourceGroupName Automation Variable. In interactive mode, discovered.
.PARAMETER AutomationAccountName
    Name of the Azure Automation Account. In runbook mode, defaults to the
    AutomationAccountName Automation Variable. In interactive mode, discovered.
.PARAMETER RuntimeEnvironment
    Name of the Runtime Environment to update. In runbook mode, defaults to
    PowerShell-7.4. In interactive mode, discovered via prompt if not provided.
.PARAMETER PollingTimeoutMinutes
    Maximum minutes to wait for package provisioning to complete. Default: 30.
.NOTES
    Requires Modules: Microsoft.PowerShell.PSResourceGet
    Interactive mode also requires: Az.Accounts, Az.Automation
    2025.01.02  - Initial Version - Andres Bohren
    2026.04.08  - Added param block, error handling, response validation,
                 polling timeout, and upgraded to GA API version 2024-10-23.
                - Added interactive mode with discovery, version comparison,
                 and selective update approval.
.EXAMPLE
    # Run as Azure Automation runbook (default - Managed Identity)
    .\Update-AzureAutomationRuntimeEnvironmentPSModules.ps1

    # Run interactively with full discovery
    .\Update-AzureAutomationRuntimeEnvironmentPSModules.ps1 -Interactive $true

    # Run interactively with pre-selected subscription and account
    .\Update-AzureAutomationRuntimeEnvironmentPSModules.ps1 -Interactive $true -SubscriptionID 'abc-123' -AutomationAccountName 'my-account' -ResourceGroupName 'my-rg'
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',
    Justification = 'Write-Host is used intentionally for interactive console UI with color support')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidLongLines', '',
    Justification = 'REST API paths and format strings exceed default line length')]
[CmdletBinding(ConfirmImpact = 'None')]
param(
    [Parameter()]
    [bool]$Interactive = $false,

    [Parameter()]
    [string]$SubscriptionID,

    [Parameter()]
    [string]$ResourceGroupName,

    [Parameter()]
    [string]$AutomationAccountName,

    [Parameter()]
    [string]$RuntimeEnvironment,

    [Parameter()]
    [ValidateRange(1, 120)]
    [int]$PollingTimeoutMinutes = 30
)

$ApiVersion = '2024-10-23'

###############################################################################
# Hardcoded module list (non-interactive / runbook mode)
###############################################################################
$RunbookModules = @(
    'Az.Accounts'
    'Az.Automation'
    'Az.Storage'
    'ExchangeOnlineManagement'
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.Beta.Security'
    'Microsoft.Graph.Calendar'
    'Microsoft.Graph.DeviceManagement'
    'Microsoft.Graph.DeviceManagement.Enrollment'
    'Microsoft.Graph.Files'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
    'Microsoft.Graph.Identity.SignIns'
    'Microsoft.Graph.Mail'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Users.Actions'
    'MicrosoftTeams'
    'PSTeams'
)

###############################################################################
# Helper Functions
###############################################################################

function Read-HostSelection {
    <#
    .SYNOPSIS
        Displays a numbered list and returns the user-selected item(s).
    .PARAMETER Items
        Array of objects to choose from.
    .PARAMETER DisplayProperty
        Property name to display. If null, displays the item itself.
    .PARAMETER Prompt
        Prompt text shown to the user.
    .PARAMETER AllowAll
        If true, adds an [A] All option that returns every item.
    #>
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Items,

        [Parameter()]
        [string]$DisplayProperty,

        [Parameter()]
        [string]$Prompt = 'Select an option',

        [Parameter()]
        [bool]$AllowAll = $false
    )

    Write-Host ''
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $display = if ($DisplayProperty) { $Items[$i].$DisplayProperty } else { $Items[$i] }
        Write-Host "  [$($i + 1)] $display"
    }
    if ($AllowAll) {
        Write-Host '  [A] All'
    }
    Write-Host ''

    do {
        $choice = Read-Host -Prompt $Prompt
        if ($AllowAll -and $choice -ieq 'A') {
            return $Items
        }
        $index = 0
        if ([int]::TryParse($choice, [ref]$index) -and $index -ge 1 -and $index -le $Items.Count) {
            return @($Items[$index - 1])
        }
        Write-Host '  Invalid selection. Please try again.' -ForegroundColor Yellow
    } while ($true)
}

function Submit-ModuleUpdate {
    <#
    .SYNOPSIS
        Submits a module update to the runtime environment via REST API PUT.
    .OUTPUTS
        True if HTTP response indicates success, false otherwise.
    #>
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EnvironmentBasePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Version
    )

    # Resolve the actual nupkg download URL via the PSGallery v2 API redirect.
    # The v2 endpoint returns 302 pointing to the current CDN host.
    $galleryDownloadUrl = "https://www.powershellgallery.com/api/v2/package/$($ModuleName)/$($Version)"
    try {
        $null = Invoke-WebRequest -Uri $galleryDownloadUrl -MaximumRedirection 0 -ErrorAction Stop
        $nupkgUri = $galleryDownloadUrl
    }
    catch {
        $redirectLocation = $_.Exception.Response.Headers.Location
        if ($redirectLocation) {
            $nupkgUri = "$($redirectLocation)"
        }
        else {
            $nupkgUri = $galleryDownloadUrl
        }
    }
    $body = @{
        properties = @{
            contentLink = @{ uri = $nupkgUri }
        }
    } | ConvertTo-Json -Depth 4

    $putPath = "$($EnvironmentBasePath)/packages/$($ModuleName)?api-version=$($ApiVersion)"
    $response = Invoke-AzRestMethod -Method 'PUT' -Path $putPath -Payload $body -ErrorAction Stop

    if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
        Write-Output "  Submitted $ModuleName $Version (HTTP $($response.StatusCode))"
        return $true
    }

    $errorDetail = $response.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
    $errorMessage = if ($errorDetail.error.message) { $errorDetail.error.message } else { $response.Content }
    Write-Error "  Failed to submit $ModuleName $Version (HTTP $($response.StatusCode)): $errorMessage"
    return $false
}

function Wait-PackageProvisioning {
    <#
    .SYNOPSIS
        Polls for provisioning completion of the specified modules only.
    .OUTPUTS
        Array of package status objects with name, Version, and ProvisioningState.
    #>
    [CmdletBinding(ConfirmImpact = 'None')]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EnvironmentBasePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ModuleNames,

        [Parameter()]
        [ValidateRange(1, 120)]
        [int]$TimeoutMinutes = 30
    )

    $timeoutEnd = (Get-Date).AddMinutes($TimeoutMinutes)
    $getPath = "$($EnvironmentBasePath)/packages?api-version=$($ApiVersion)"
    $packageStatus = $null

    do {
        Write-Output 'Checking package provisioning status...'
        try {
            $result = Invoke-AzRestMethod -Method 'GET' -Path $getPath -ErrorAction Stop

            if ($result.StatusCode -lt 200 -or $result.StatusCode -ge 300) {
                Write-Error "Failed to query packages (HTTP $($result.StatusCode)): $($result.Content)"
                break
            }

            $allPackages = ($result.Content | ConvertFrom-Json).value
            $packageStatus = $allPackages |
                Where-Object { $_.name -in $ModuleNames } |
                Select-Object name,
                    @{ Name = 'Version'; Expression = { $_.properties.version } },
                    @{ Name = 'ProvisioningState'; Expression = { $_.properties.provisioningState } }

            $pending = $packageStatus | Where-Object {
                $_.ProvisioningState -ne 'Succeeded' -and $_.ProvisioningState -ne 'Failed'
            }

            if ($null -eq $pending) {
                break
            }

            $pendingNames = ($pending | ForEach-Object { "$($_.name) ($($_.ProvisioningState))" }) -join ', '
            Write-Output "  Still pending: $pendingNames"

            if ((Get-Date) -ge $timeoutEnd) {
                Write-Error "Polling timed out after $TimeoutMinutes minutes. Still pending: $pendingNames"
                break
            }

            Start-Sleep -Seconds 15
        }
        catch {
            Write-Error "Error while polling package status: $_"
            break
        }
    } while ($true)

    return $packageStatus
}

###############################################################################
# Connection & Variable Resolution
###############################################################################

if ($Interactive) {
    # --- Prerequisite check ---
    $missingModules = @()
    foreach ($requiredModule in @('Az.Accounts', 'Az.Automation', 'Microsoft.PowerShell.PSResourceGet')) {
        if (-not (Get-Module -Name $requiredModule -ListAvailable)) {
            $missingModules += $requiredModule
        }
    }
    if ($missingModules.Count -gt 0) {
        Write-Error "Missing required modules for interactive mode: $($missingModules -join ', '). Install with: Install-Module -Name <ModuleName> -Scope CurrentUser"
        exit 1
    }

    # --- Interactive sign-in ---
    Write-Host 'Signing in to Azure...' -ForegroundColor Cyan
    try {
        $connectParams = @{ ErrorAction = 'Stop' }
        if ($SubscriptionID) {
            $connectParams['Subscription'] = $SubscriptionID
        }
        Connect-AzAccount @connectParams | Out-Null
    }
    catch {
        Write-Error "Failed to sign in to Azure: $_"
        exit 1
    }
    Write-Host 'Signed in successfully.' -ForegroundColor Green

    # --- Resolve subscription ---
    if (-not $SubscriptionID) {
        # Connect-AzAccount prompts for subscription when multiple exist;
        # read the selection from the current Az context.
        $currentContext = Get-AzContext
        if ($currentContext -and $currentContext.Subscription) {
            $SubscriptionID = $currentContext.Subscription.Id
            Write-Host "Using subscription: $($currentContext.Subscription.Name) ($($SubscriptionID))" -ForegroundColor Cyan
        }
        else {
            Write-Error 'No subscription selected during sign-in.'
            exit 1
        }
    }
    else {
        Set-AzContext -SubscriptionId $SubscriptionID -ErrorAction Stop | Out-Null
    }

    # --- Select Automation Account ---
    if (-not $AutomationAccountName -or -not $ResourceGroupName) {
        Write-Host 'Discovering Automation Accounts...' -ForegroundColor Cyan
        $accounts = Get-AzAutomationAccount -ErrorAction Stop | Sort-Object -Property AutomationAccountName
        if ($accounts.Count -eq 0) {
            Write-Error 'No Automation Accounts found in this subscription.'
            exit 1
        }
        elseif ($accounts.Count -eq 1) {
            $AutomationAccountName = $accounts[0].AutomationAccountName
            $ResourceGroupName = $accounts[0].ResourceGroupName
            Write-Host "Using account: $($AutomationAccountName) (RG: $($ResourceGroupName))" -ForegroundColor Cyan
        }
        else {
            Write-Host 'Select an Automation Account:' -ForegroundColor Cyan
            $displayItems = $accounts | ForEach-Object {
                [PSCustomObject]@{
                    Display = "$($_.AutomationAccountName) (RG: $($_.ResourceGroupName))"
                    Account = $_
                }
            }
            $selected = Read-HostSelection -Items $displayItems -DisplayProperty 'Display' -Prompt 'Account number'
            $AutomationAccountName = $selected[0].Account.AutomationAccountName
            $ResourceGroupName = $selected[0].Account.ResourceGroupName
        }
    }

    # --- Select Runtime Environment ---
    $accountBasePath = "/subscriptions/$($SubscriptionID)/resourceGroups/$($ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationAccountName)"

    if (-not $RuntimeEnvironment) {
        Write-Host 'Discovering Runtime Environments...' -ForegroundColor Cyan
        $rtePath = "$($accountBasePath)/runtimeEnvironments?api-version=$($ApiVersion)"
        $rteResult = Invoke-AzRestMethod -Method 'GET' -Path $rtePath -ErrorAction Stop

        if ($rteResult.StatusCode -lt 200 -or $rteResult.StatusCode -ge 300) {
            Write-Error "Failed to list runtime environments (HTTP $($rteResult.StatusCode)): $($rteResult.Content)"
            exit 1
        }

        $environments = ($rteResult.Content | ConvertFrom-Json).value
        $envNames = @($environments | ForEach-Object { $_.name } | Sort-Object)

        if ($envNames.Count -eq 0) {
            Write-Error 'No runtime environments found in this Automation Account.'
            exit 1
        }
        elseif ($envNames.Count -eq 1) {
            $RuntimeEnvironment = $envNames[0]
            Write-Host "Using runtime environment: $($RuntimeEnvironment)" -ForegroundColor Cyan
        }
        else {
            Write-Host 'Select a Runtime Environment:' -ForegroundColor Cyan
            $selected = Read-HostSelection -Items $envNames -Prompt 'Environment number'
            $RuntimeEnvironment = $selected[0]
        }
    }
}
else {
    # --- Runbook mode: resolve from Automation Variables ---
    if (-not $SubscriptionID) { $SubscriptionID = Get-AutomationVariable -Name 'SubscriptionID' }
    if (-not $ResourceGroupName) { $ResourceGroupName = Get-AutomationVariable -Name 'ResourceGroupName' }
    if (-not $AutomationAccountName) { $AutomationAccountName = Get-AutomationVariable -Name 'AutomationAccountName' }
    if (-not $RuntimeEnvironment) { $RuntimeEnvironment = 'PowerShell-7.4' }

    Write-Output 'Connecting to Azure with Managed Identity...'
    try {
        Connect-AzAccount -Identity -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to connect to Azure: $_"
        exit 1
    }
}

# Build base path after all variables are resolved
$BasePath = "/subscriptions/$($SubscriptionID)/resourceGroups/$($ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($AutomationAccountName)/runtimeEnvironments/$($RuntimeEnvironment)"

###############################################################################
# Determine which modules to update
###############################################################################

$modulesToUpdate = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($Interactive) {
    # --- Retrieve installed packages ---
    Write-Host ''
    Write-Host 'Retrieving installed packages...' -ForegroundColor Cyan
    $pkgPath = "$($BasePath)/packages?api-version=$($ApiVersion)"
    $pkgResult = Invoke-AzRestMethod -Method 'GET' -Path $pkgPath -ErrorAction Stop

    if ($pkgResult.StatusCode -lt 200 -or $pkgResult.StatusCode -ge 300) {
        Write-Error "Failed to list packages (HTTP $($pkgResult.StatusCode)): $($pkgResult.Content)"
        exit 1
    }

    $installedPackages = @(($pkgResult.Content | ConvertFrom-Json).value)

    if ($installedPackages.Count -eq 0) {
        Write-Host 'No packages installed in this runtime environment.' -ForegroundColor Yellow
        exit 0
    }

    # --- Build version comparison table ---
    $comparisonResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0

    foreach ($pkg in $installedPackages) {
        $i++
        $moduleName = $pkg.name
        $installedVersion = $pkg.properties.version
        Write-Progress -Activity 'Comparing module versions' -Status $moduleName -PercentComplete (($i / $installedPackages.Count) * 100)

        $latestVersion = $null
        $status = 'Unknown'

        try {
            $psResource = Find-PSResource -Name $moduleName -Repository PSGallery -ErrorAction Stop |
                Select-Object -First 1

            if ($psResource) {
                $latestVersion = $psResource.Version.ToString()
                if ($installedVersion -eq $latestVersion) {
                    $status = 'Up to date'
                }
                else {
                    try {
                        $installedVer = [System.Version]$installedVersion
                        $latestVer = [System.Version]$latestVersion
                        $status = if ($latestVer -gt $installedVer) { 'Update available' } else { 'Up to date' }
                    }
                    catch {
                        # Version parsing failed - fall back to string inequality
                        $status = if ($installedVersion -ne $latestVersion) { 'Update available' } else { 'Up to date' }
                    }
                }
            }
            else {
                $status = 'Not in PSGallery'
            }
        }
        catch {
            $status = 'Not in PSGallery'
        }

        $comparisonResults.Add([PSCustomObject]@{
            Module    = $moduleName
            Installed = $installedVersion
            Latest    = if ($latestVersion) { $latestVersion } else { '-' }
            Status    = $status
        })
    }
    Write-Progress -Activity 'Comparing module versions' -Completed

    # --- Display comparison table ---
    $sorted = $comparisonResults | Sort-Object -Property Module
    Write-Host ''
    Write-Host ('=' * 100) -ForegroundColor DarkGray
    Write-Host ('{0,-45} {1,-15} {2,-15} {3,-20}' -f 'Module', 'Installed', 'Latest', 'Status') -ForegroundColor White
    Write-Host ('-' * 100) -ForegroundColor DarkGray

    foreach ($item in $sorted) {
        $color = switch ($item.Status) {
            'Update available' { 'Yellow' }
            'Up to date'       { 'Green' }
            default            { 'Gray' }
        }
        Write-Host ('{0,-45} {1,-15} {2,-15} {3,-20}' -f $item.Module, $item.Installed, $item.Latest, $item.Status) -ForegroundColor $color
    }
    Write-Host ('=' * 100) -ForegroundColor DarkGray

    # --- Filter to updatable modules ---
    $updatable = @($comparisonResults | Where-Object { $_.Status -eq 'Update available' })

    if ($updatable.Count -eq 0) {
        Write-Host ''
        Write-Host 'All modules are up to date.' -ForegroundColor Green
        Disconnect-AzAccount | Out-Null
        exit 0
    }

    Write-Host ''
    Write-Host "$($updatable.Count) module(s) have updates available." -ForegroundColor Yellow
    Write-Host ''

    # --- Prompt for update strategy ---
    $choices = @(
        [System.Management.Automation.Host.ChoiceDescription]::new('&All', 'Update all outdated modules')
        [System.Management.Automation.Host.ChoiceDescription]::new('&Select individually', 'Choose which modules to update one at a time')
        [System.Management.Automation.Host.ChoiceDescription]::new('&Cancel', 'Exit without updating')
    )
    $decision = $host.UI.PromptForChoice('Module Updates', 'How would you like to proceed?', $choices, 0)

    switch ($decision) {
        0 {
            foreach ($mod in $updatable) {
                $modulesToUpdate.Add([PSCustomObject]@{ Name = $mod.Module; Version = $mod.Latest })
            }
        }
        1 {
            foreach ($mod in ($updatable | Sort-Object -Property Module)) {
                $moduleChoices = @(
                    [System.Management.Automation.Host.ChoiceDescription]::new('&Yes', "Update $($mod.Module)")
                    [System.Management.Automation.Host.ChoiceDescription]::new('&No', "Skip $($mod.Module)")
                )
                $moduleDecision = $host.UI.PromptForChoice(
                    $mod.Module,
                    "Update from $($mod.Installed) to $($mod.Latest)?",
                    $moduleChoices,
                    0
                )
                if ($moduleDecision -eq 0) {
                    $modulesToUpdate.Add([PSCustomObject]@{ Name = $mod.Module; Version = $mod.Latest })
                }
            }
        }
        2 {
            Write-Host 'Cancelled. No modules updated.' -ForegroundColor Yellow
            Disconnect-AzAccount | Out-Null
            exit 0
        }
    }

    if ($modulesToUpdate.Count -eq 0) {
        Write-Host 'No modules selected for update.' -ForegroundColor Yellow
        Disconnect-AzAccount | Out-Null
        exit 0
    }

    Write-Host ''
    Write-Host "Updating $($modulesToUpdate.Count) module(s)..." -ForegroundColor Cyan
}
else {
    # --- Runbook mode: resolve latest versions for hardcoded module list ---
    foreach ($Module in $RunbookModules) {
        try {
            Write-Output "Processing module: $Module"
            $psResource = Find-PSResource -Name $Module -Repository PSGallery -ErrorAction Stop |
                Select-Object -First 1

            if ($psResource) {
                $modulesToUpdate.Add([PSCustomObject]@{
                    Name    = $Module
                    Version = $psResource.Version.ToString()
                })
            }
            else {
                Write-Error "  Module $Module not found in PSGallery"
            }
        }
        catch {
            Write-Error "  Error finding module ${Module}: $_"
        }
    }
}

###############################################################################
# Submit Updates
###############################################################################
$failedModules = @()
$succeededModules = @()

foreach ($mod in $modulesToUpdate) {
    try {
        Write-Output "Submitting update: $($mod.Name) $($mod.Version)"
        $success = Submit-ModuleUpdate -EnvironmentBasePath $BasePath -ModuleName $mod.Name -Version $mod.Version

        if ($success) {
            $succeededModules += "$($mod.Name) $($mod.Version)"
        }
        else {
            $failedModules += $mod.Name
        }
    }
    catch {
        Write-Error "  Error submitting module $($mod.Name): $_"
        $failedModules += $mod.Name
    }
}

###############################################################################
# Poll for Package Installation Completion
###############################################################################
if ($succeededModules.Count -gt 0) {
    $submittedNames = @($modulesToUpdate | ForEach-Object { $_.Name })
    $packageStatus = Wait-PackageProvisioning -EnvironmentBasePath $BasePath -ModuleNames $submittedNames -TimeoutMinutes $PollingTimeoutMinutes

    $failedPackages = $packageStatus | Where-Object { $_.ProvisioningState -eq 'Failed' }
    if ($failedPackages) {
        $failedNames = ($failedPackages | ForEach-Object { $_.name }) -join ', '
        Write-Error "The following packages failed to install: $failedNames"
    }
}

###############################################################################
# Summary
###############################################################################
if ($failedModules.Count -gt 0) {
    Write-Error "The following modules failed during submission: $($failedModules -join ', ')"
}

if ($succeededModules.Count -gt 0) {
    Write-Output "Submitted successfully: $($succeededModules -join ', ')"
}

Write-Output 'Package update complete.'

###############################################################################
# Disconnect from Azure
###############################################################################
Write-Output 'Disconnecting from Azure...'
Disconnect-AzAccount