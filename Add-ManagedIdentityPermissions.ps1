#Requires -Modules Microsoft.Graph.Applications
$TenantId = "21546e07-2132-445d-b21e-9ca36bf847c2" # Azure Tenant ID, can be found at https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Overview
$MSIObjectId = "8e26b606-cf77-4089-ab92-0a91fb0d668c" # Object ID of the system-assigned or user-assigned managed service identity. (System-assigned use same name as resource).
$GraphAppId = "00000003-0000-0000-c000-000000000000" # Don't change this. This is the immutable application ID of the Microsoft Graph service principal.
#Exchange Permissions
$exchangeAppRoleID = "dc50a0fb-09a3-484d-be87-e023b12c6440" #Exchange Online Application Role ID -> always the same in every tenant 
$exchangeResourceID = (Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'").Id

$oAssignPermissions = @(
    "MailboxSettings.ReadWrite"
    "Organization.Read.All"
    "User.Read.All"
    "User.ReadBasic.All"
    "Group.ReadWrite.All"
    "GroupMember.Read.All"
    "GroupMember.ReadWrite.All"
    "Mail.ReadWrite"
    "Mail.ReadWrite.Shared"
    "Mail.Send"
    "Mail.Send.Shared"
    "User.EnableDisableAccount.All"
    "UserAuthenticationMethod.ReadWrite.All"
    "Exchange.ManageAsApp"
)

# If your account is restricted to only certain permissions, uncomment the following lines and use these to scope your connection.
 $MgRequiredScopes = @(
     "Application.Read.All"
     "AppRoleAssignment.ReadWrite.All"
     "Directory.Read.All"
 )

Connect-MgGraph -TenantId $TenantId -Scopes $MgRequiredScopes -NoWelcome #Uncomment NoWelcome if desired

$oMsi = Get-MgServicePrincipal -ServicePrincipalId $MSIObjectId
$oGraphSpn = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"

$oAppRole = $oGraphSpn.AppRoles | Where-Object { ($_.Value -in $oAssignPermissions) -and ($_.AllowedMemberTypes -contains "Application") }

foreach ($AppRole in $oAppRole) {
    $oAppRoleAssignment = @{
        "PrincipalId" = $oMSI.Id
        "ResourceId"  = $oGraphSpn.Id
        "AppRoleId"   = $AppRole.Id
    }
  
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $oAppRoleAssignment.PrincipalId `
        -BodyParameter $oAppRoleAssignment `
        -Verbose
}

New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $oMSI.Id -PrincipalId $oMSI.Id -AppRoleId $exchangeAppRoleID -ResourceId $exchangeResourceID

# You should manually assign the system managed identity an Entra role that has permissions to manage mailboxes if you use the shared mailbox permissions functionality.
# For example, assigning the managed account the Exchange Administrator role.