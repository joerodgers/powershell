<#

 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. 
 
 This sample assumes that you are familiar with the programming language being demonstrated and the 
 tools used to create and debug procedures. Microsoft support professionals can help explain the 
 functionality of a particular procedure, but they will not modify these examples to provide added 
 functionality or construct procedures to meet your specific needs. if you have limited programming 
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
 line at (800) 936-5200. 


----------------------------------------------------------
History
----------------------------------------------------------
 07-26-2017 - Created
  
==============================================================#>

Install-Module -Name AzureADPreview
Import-Module  -Name AzureADPreview

# make sure the module loaded 
if (-not (Get-Module -Name AzureADPreview))
{
    Write-Error -Message "The AzureADPreview module has not been loaded."
    return
} 

# make sure the correct version of the module loaded
if (-not (Get-Command -Name "Get-AzureADDirectorySetting" -ErrorAction SilentlyContinue))
{
    Write-Error -Message "The required version of the AzureADPreview module has not been loaded."
    return
}

# user will need to be a tenant admin
Connect-AzureAD

$directorySettings = Get-AzureADDirectorySetting | ? { $_.DisplayName -eq "Group.Unified" }
 
if( -not $directorySettings )
{
    # get the template 
    $template = Get-AzureADDirectorySettingTemplate | ? {$_.DisplayName -eq 'Group.Unified'}
 
    # create the settings from the template 
    $directorySetting = $template.CreateDirectorySetting()
    New-AzureADDirectorySetting -DirectorySetting $directorySetting

    $directorySettings = Get-AzureADDirectorySetting | ? { $_.DisplayName -eq "Group.Unified" }
}  

if($directorySettings)
{
    $directorySettings["EnableGroupCreation"]           = "true"
    $directorySettings["ClassificationList"]            = "Confidential,Secret,Top Secret,Cosmic Top Secret"
    $directorySettings["ClassificationDescriptions"]    = "Confidential:Confidential Information,Secret:Secret Information,Top Secret:Top Secret Information,Cosmic Top Secret:The stuff about aliens" 
    $directorySettings["DefaultClassification"]         = "Confidential"
    $directorySettings["AllowGuestsToBeGroupOwner"]     = "false"
    $directorySettings["AllowGuestsToAccessGroups"]     = "true" 
    $directorySettings["UsageGuidelinesUrl"]            = "https://contoso.sharepoint.com/sites/support/documents/usage-guidelines.docx" 
    $directorySettings["GuestUsageGuidelinesUrl"]       = "https://contoso.sharepoint.com/sites/support/documents/usage-guidelines.docx" 
    $directorySettings["AllowToAddGuests"]              = "true"
    # $directorySettings["PrefixSuffixNamingRequirement"] = "o365grp-" # this one doesn't seem to work yet

    # adjust to fit your needs
    # $domainGroup = Get-AzureADGroup | ? { $_.DisplayName -eq "O365GroupAdmins" }
    # $directorySettings["GroupCreationAllowedGroupId"] = $$domainGroup.ObjectId 

    Set-AzureADDirectorySetting -Id $directorySettings.Id -DirectorySetting $directorySettings
    
    # show the settings after the update
    Get-AzureADDirectorySetting | ? { $_.DisplayName -eq "Group.Unified" } | SELECT -ExpandProperty Values  | FT -AutoSize
}
else
{
    Write-Warning "No AD directory settings for 'Group.Unified' found, no updates have been made." 
}
