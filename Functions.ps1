Import-Module SharePointPnPPowerShellOnline -Force -DisableNameChecking
Import-Module Microsoft.Online.SharePoint.PowerShell -Force -DisableNameChecking

function ConvertTo-Dictionary {
    [CmdletBinding()] param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [hashtable]
        $InputObject
    )

    process {
        $outputObject = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"

        foreach ($entry in $InputObject.GetEnumerator()) {
            $newKey = $entry.Key -as [System.String]
            
            if ($null -eq $newKey) {
                throw 'Could not convert key "{0}" of type "{1}" to type "{2}"' -f
                      $entry.Key,
                      $entry.Key.GetType().FullName,
                      "System.String"
            } elseif ($outputObject.ContainsKey($newKey)) {
                throw "Duplicate key `"$newKey`" detected in input object."
            }

            $outputObject.Add($newKey, $entry.Value)
        }
        return $outputObject
    }
}

function Get-AzureADUserPropertiesAsJson {
    [CmdletBinding()] param (
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
        [String[]] $ExtensionProperties,

        [Parameter()]
        [String] $OutFile = 'userdata.json'
    )

    $users = $null # Holding object

    if ($SearchString) {
        $users = Get-AzureADUser -SearchString $SearchString | Select-Object *
    } else {
        $users = Get-AzureADUser -All $true | Select-Object *
    }

    if ($users) {
        $output = @()
        foreach ($user in $users) {
            $userObject = New-Object PSObject
            if ($user.Mail) {
                # Mandatory fields
                $userObject | Add-Member -NotePropertyName "idName" -NotePropertyValue $user.Mail

                # Requested fields
                foreach ($item in $Properties) {
                    if ($null -eq ($user.$item)) {
                        $userObject | Add-Member -NotePropertyName $item -NotePropertyValue ""
                    } else {
                        $userObject | Add-Member -NotePropertyName $item -NotePropertyValue $user.$item
                    }
                }

                # Extension Properties
                if ($ExtensionProperties) {
                    foreach ($property in $ExtensionProperties) {
                        $userEPs = $user.ExtensionProperty
                        if ($null -eq ($userEPs.$property)) {
                            $userObject | Add-Member -NotePropertyName $property -NotePropertyValue ""
                        } else {
                            $userObject | Add-Member -NotePropertyName $property -NotePropertyValue $userEPs.$property
                        }
                    }
                }
                
                $output += $userObject
            }
        }
        $jsonOutput = @{value = $output} | ConvertTo-Json
        $jsonOutput | Out-File $OutFile
        return $output
    } else {
        Write-Error "No users returned from Azure AD."
    }
}

function Update-SPAttributesFromJson {
    [CmdletBinding()] param (
        [Parameter(Mandatory)]
        [System.Uri] $SharePointAdminSiteUrl,

        [Parameter(Mandatory)]
        [PSCredential] $Credential,

        # Mapping between source file property name and the destination property name in SharePoint
        # Source (AzureAD) property name is the key, SharePoint property is the value
        [Parameter(Mandatory)]
        [Hashtable] $PropertyMap,

        [Parameter(Mandatory)]
        [System.Uri] $JsonFileUrl
    )

    $propertyDictMap = ConvertTo-Dictionary -InputObject $PropertyMap

    $username = $Credential.UserName
    $password = $Credential.GetNetworkCredential().Password | ConvertTo-SecureString -AsPlainText -Force
    
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SharePointAdminSiteUrl)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)

    $office365 = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($context)
    $context.Load($office365)

    # Type of user identifier ["Email", "CloudId", "PrincipalName"] in the user profile service
    $userIdType = [Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email

    # Name of user identifier property in the JSON
    $userLookupKey="idName"

    # Call to queue UPA property import 
    $workItemId = $office365.QueueImportProfileProperties($userIdType, $userLookupKey, $propertyDictMap, $JsonFileUrl);

    # Execute the CSOM command for queuing the import job
    $context.ExecuteQuery();

    # Return the unique identifier of the job
    if ($workItemId) {
        return $workItemId
    }
    return $null
}

function Export-CredentialToFile {
    [CmdletBinding()] param (
        [Parameter(Mandatory)]
        [PSCredential] $Credential,

        [Parameter(Mandatory)]
        [String] $Path
    )

    $Credential | Export-CliXml -Path $Path
}

function Import-CredentialFromFile {
    [CmdletBinding()] param (
        [Parameter(Mandatory)]
        [String] $Path
    )

    [PSCredential] $credential = Import-Clixml -Path $Path
    return $credential
}

function Write-FileToSharePoint {
    [CmdletBinding()] param (
        [Parameter(Mandatory)]
        [System.Uri] $SharePointSiteUrl,

        [Parameter(Mandatory)]
        [PSCredential] $Credential,

        [Parameter(Mandatory)]
        [String] $SourceFile,

        [Parameter(Mandatory)]
        [String] $DocumentLibraryName
    )

    Connect-PnPOnline -Url $SharePointSiteUrl -Credentials $Credential

    Add-PnPFile -Path $SourceFile -Folder $DocumentLibraryName
}

<#
# SETUP
Install-Module AzureAd -Scope AllUsers -Force
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope AllUsers -Force
Install-Module SharePointPnPPowerShellOnline -Scope AllUsers -Force
$Credential = Get-Credential
Export-CredentialToFile -Credential $Credential -Path 'C:\UpdateSharePointProperties\Credential.cred'
#>

#
# RUN
$azureProperties = @('DisplayName','GivenName','Surname','JobTitle','Department','Mobile','PhysicalDeliveryOfficeName','StreetAddress','City','State','PostalCode','TelephoneNumber')
$azureExtensionProperties = @('extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute1','extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute10','extension_5412726b57b245199a74ff6529fff9d2_wWWHomePage')

# AzureAD = SharePoint
$propertyMap = @{
    DisplayName = 'PreferredName'
    GivenName = 'FirstName'
    Surname = 'LastName'
    JobTitle = 'title'
    Department = 'Department'
    Mail = 'workemail'
    Mobile = 'CellPhone'
    PhysicalDeliveryOfficeName = 'Office'
    TelephoneNumber = 'workphone'
    StreetAddress = 'street'
    City = 'city'
    State = 'state'
    PostalCode = 'zip'
    extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute1 = 'ext'
    extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute10 = 'LinkedIn'
    extension_5412726b57b245199a74ff6529fff9d2_wWWHomePage = 'WWW'
}

$WorkingFolder = 'C:\UpdateSharePointProperties'
try {
    Start-Transcript "$WorkingFolder\Transcript.txt"
    $Credential = Import-CredentialFromFile -Path "$WorkingFolder\SPAdmin.cred"
    Connect-AzureAD -Credential $Credential -ErrorAction Stop
    $userdata = Get-AzureADUserPropertiesAsJson -AllUsers -Properties $azureProperties -ExtensionProperties $azureExtensionProperties -OutFile "$($WorkingFolder)\userdata.json" -ErrorAction Stop
    $userdata | Select-Object DisplayName,extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute1,PhysicalDeliveryOfficeName,JobTitle,Mail,TelephoneNumber,Mobile | Export-Csv "$($WorkingFolder)\userdata.csv" -NoTypeInformation
    # Change the columns to friendly names for the CSV with calculated properties!
    #$userdata | Select-Object @{expression={$_.DisplayName}; label="Display Name"},@{expression={$_.extension_5412726b57b245199a74ff6529fff9d2_extensionAttribute1}; label="Extension"},@{expression={$_.PhysicalDeliveryOfficeName}; label="Office"},@{expression={$_.JobTitle}; label="Title"},Mail,TelephoneNumber,Mobile | Export-Csv "$($WorkingFolder)\userdata.csv" -NoTypeInformation
    Write-FileToSharePoint -SharePointSiteUrl 'https://<tenant>.sharepoint.com/sites/directory' -Credential $Credential -SourceFile "$($WorkingFolder)\userdata.json" -DocumentLibraryName 'Documents' -ErrorAction Stop | Out-Null
    Write-FileToSharePoint -SharePointSiteUrl 'https://<tenant>.sharepoint.com/sites/directory' -Credential $Credential -SourceFile "$($WorkingFolder)\userdata.csv" -DocumentLibraryName 'Documents' | Out-Null
    Update-SPAttributesFromJson -SharePointAdminSiteUrl 'https://<tenant>-admin.sharepoint.com' -Credential $Credential -PropertyMap $propertyMap -JsonFileUrl 'https://<tenant>.sharepoint.com/sites/directory/Documents/userdata.json'
    Stop-Transcript
} catch {
    throw $_
} finally {
    Disconnect-AzureAD
    Stop-Transcript
}