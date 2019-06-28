<#
USAGE

1. Dot source this file into your session: . .\Functions.ps1
2. (Optional) Fetch the properties you need from Azure AD (Use Powershell Get-AzureADUser to find the actual property names, not just the display name of those properties)
3. Connect to Azure AD (If not already connected): Connect-AzureAD
4. Fetch the properties to JSON file (examples):
    - NOTE: AzureADUserPropertiesAsJson writes a file called users.json to the current working directory
    Get-AzureADUserPropertiesAsJson -AllUsers -Properties Title,DisplayName,etc
    Get-AzureADUserPropertiesAsJson -SearchString "dustin" -Properties Title,DisplayName,etc
5. Upload the file to SharePoint and obtain the URL
- NOTE: The source file must be uploaded to the same SharePoint Online tenant where the process is started

6. Build hashtable for property map needed to map the source property name to the SharePoint property name
    $propertyMap = @{
        DisplayName = 'cn'
        Title = 'title'
        Department = 'Department'
        Mail = 'mail'
        Mobile = 'mobile'
        PhysicalDeliveryOfficeName = 'physicalDeliveryOfficeName'
    }

7. Queue properties import:
- NOTE: Returned value is the GUID of the import job
- NOTE: The connection to SharePoint online here does not support MFA
        either use an account that doesn't have MFA or an app password
  Update-SPAttributesFromJson -Credentials (Get-Credentials) -SharePointSharePointAdminUrl "" -JsonFileUrl "..." -PropertyMap $pm

Related: https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/bulk-user-profile-update-api-for-sharepoint-online
#>

function Get-AzureADUserPropertiesAsJson {
    param (
        # Get all users?
        [Parameter(Mandatory,Position=0,ParameterSetName="All")]
        [Switch] $AllUsers,

        # Get users based on Azure AD SearchString parameter?
        [Parameter(Mandatory,Position=0,ParameterSetName="Selection")]
        [String] $SearchString,

        # Properties from Azure AD to fetch into JSON output
        [Parameter(Mandatory,Position=1)]
        [String[]] $Properties,

        [Parameter(Position=2)]
        [String[]] $ExtensionProperties
    )

    $users = $null # Holding object

    if ($SearchString) {
        $users = Get-AzureADUser -SearchString $SearchString | Select-Object *
    } else {
        $users = Get-AzureADUser -All $true | Select-Object *
    }

    if ($users) {
        $jsonOutput = @{value = @() }
        foreach ($user in $users) {
            if ($user.Mail) {
                # Mandatory fields
                $userProperties = [ordered]@{
                    idName = $user.Mail
                }

                # Requested fields
                foreach ($item in $Properties) {
                    $userProperties += [ordered]@{
                        $item = $user.$item
                    }
                }

                # Extension Properties
                if ($ExtensionProperties) {
                    foreach ($property in $ExtensionProperties) {
                        $userEPs = $user.ExtensionProperty
                        $userProperties += [ordered]@{
                            $property = $userEPs.$property
                        }
                    }
                }

                $jsonOutput['value'] += $userProperties
            }
        }
        $jsonOutput | ConvertTo-Json | Out-File 'users.json'
        return $jsonOutput.Value | ForEach-Object { [pscustomobject] $_ } | Format-Table
    } else {
        Write-Error "No users returned from Azure AD."
    }
}

function Update-SPAttributesFromJson {
    param (
        [Parameter(Mandatory)]
        [String] $SPAdminUrl,

        # These credentials must work without MFA for both AzureAD and SharePoint Online
        [Parameter(Mandatory)]
        [PSCredential] $Credential,

        # Mapping between source file property name and the destination property name in SharePoint
        # Source (AzureAD) property name is the key, SharePoint property is the value
        [Parameter(Mandatory)]
        [Hashtable] $PropertyMap,

        [Parameter(Mandatory)]
        [String] $JsonFileUrl
    )

    # Get instances to the Office 365 tenant using CSOM
    $uri = New-Object System.Uri -ArgumentList $SPAdminUrl
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($uri)

    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.GetNetworkCredential().Password)
    $o365 = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($context)
    $context.Load($o365)

    # Type of user identifier ["Email", "CloudId", "PrincipalName"] in the user profile service
    $userIdType = [Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email

    # Name of user identifier property in the JSON
    $userLookupKey="idName"

    # Call to queue UPA property import 
    $workItemId = $o365.QueueImportProfileProperties($userIdType, $userLookupKey, $PropertyMap, $JsonFilePath);

    # Execute the CSOM command for queuing the import job
    $context.ExecuteQuery();

    # Output the unique identifier of the job
    Write-Output "Import job created with the following identifier: $($workItemId.Value)"
}

<#
$azureProperties = @('DisplayName','JobTitle','Department','Mobile','PhysicalDeliveryOfficeName','City','State','PostalCode','TelephoneNumber')
$azureEProperties = @('extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute1')

# AzureAD = SharePoint
$propertyMap = @{
    DisplayName = 'cn'
    JobTitle = 'title'
    Department = 'Department'
    Mail = 'mail'
    Mobile = 'mobile'
    PhysicalDeliveryOfficeName = 'physicalDeliveryOfficeName'
    City = ''
    State = ''
    PostalCode = ''
    TelephoneNumber = ''
}

Connect-AzureAD
Get-AzureADUserPropertiesAsJson -AllUsers -Properties $azureProperties -ExtensionProperties $azureExtensionProperties
Update-SPAttributesFromJson -SPAdminUrl '' -Credential (Get-Credential) -PropertyMap $propertyMap -JsonFileUrl ''
#>