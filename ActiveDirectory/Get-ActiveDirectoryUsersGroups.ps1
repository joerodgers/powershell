<#
     This sample code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
     THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
     INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
 
 ----------------------------------------------------------
 History
 ----------------------------------------------------------
 01-04-2017 - Created

==============================================================#>


function Get-ActiveDirectoryUserForMappingFile
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$false)][string]$Filter = "(objectCategory=User)"
    )

    begin
    {
        $results = @()

        $domain   = New-Object System.DirectoryServices.DirectoryEntry
        $searcher = New-Object System.DirectoryServices.DirectorySearcher 
    }
    process
    {
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot  = $domain
        $searcher.PageSize    = 1000
        $searcher.Filter      = $Filter 
        $searcher.SearchScope = "Subtree"    
        
        $searchResults = $searcher.FindAll()

        foreach( $user in $searchResults )
        {
            $upn  =  $user.Properties["userPrincipalName"][0]
            $mail =  $user.Properties["mail"][0]

            # only include mappings that have both UPN and mail values
            if( $upn -and $mail )
            {
                $results += New-Object PSObject -Property @{
                    "SourceUser" = $upn
                    "Target UPN" = $mail
                }
            }
        }
    }
    end
    {
        $results
    }
}

function Get-ActiveDirectoryGroupForMappingFile
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$false)][string]$Filter = "(objectCategory=Group)"
    )

    begin
    {
        $results = @()

        $domain   = New-Object System.DirectoryServices.DirectoryEntry
        $searcher = New-Object System.DirectoryServices.DirectorySearcher 
    }
    process
    {
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.SearchRoot  = $domain
        $searcher.PageSize    = 1000
        $searcher.Filter      = $Filter 
        $searcher.SearchScope = "Subtree"    
        
        $searchResults = $searcher.FindAll()

        foreach( $user in $searchResults )
        {
            $groupName =  $user.Properties["name"][0]
            $bytes = $user.Properties["objectSid"][0]
            $securityIdentifier = New-Object System.Security.Principal.SecurityIdentifier( $bytes, 0 )

            if( $groupName -and $securityIdentifier.ToString() )
            {
                $results += New-Object PSObject -Property @{
                    "GroupName" = $groupName
                    "GroupSID"  = $securityIdentifier.ToString()
                }
            }
        }
    }
    end
    {
        $results
    }
}

# (object type is a user) and (not disabled) and (has a email mail value) and (starts with a* or n*)
$userFilter = "(& (objectClass=user) (!userAccountControl:1.2.840.113556.1.4.803:=2) (!mail=*) (|(cn=n*)(cn=a*)) )"
Get-ActiveDirectoryUserForMappingFile -Filter $userFilter | SELECT "SourceUser", "Target UPN" | Export-Csv -Path "DomainUserMappingFile_$($(Get-Date).ToString('yyyy-MM-dd')).csv" -NoTypeInformation

# object type is a group
$groupFilter = "(objectCategory=Group)"
Get-ActiveDirectoryGroupForMappingFile -Filter $groupFilter | SELECT "GroupName", "GroupSID" | Export-Csv -Path "DomainGroupMappingFile_$($(Get-Date).ToString('yyyy-MM-dd')).csv" -NoTypeInformation


