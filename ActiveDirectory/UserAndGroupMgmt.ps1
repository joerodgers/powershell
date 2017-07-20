 <#

     This sample code is provided for the purpose of illustration only and is not intended to be used in a production environment.  

     THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
     INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 ----------------------------------------------------------
 History
 ----------------------------------------------------------
 06-07-2017 - Created

==============================================================#>

[CmdletBinding(DefaultParameterSetName="SiteCollectionScope")]
param 
(
    [parameter(Position=0, Mandatory=$true, ParameterSetName="GetUserMembership")]
    [parameter(Position=0, Mandatory=$true, ParameterSetName="AddUserToGroup")]
    [parameter(Position=0, Mandatory=$true, ParameterSetName="RemoveUserFromGroup")]
    [string[]]$UserDistinguishedName,
    
    [parameter(Position=0, Mandatory=$true, ParameterSetName="AddUserToGroup")]
    [parameter(Position=0, Mandatory=$true, ParameterSetName="RemoveUserFromGroup")]
    [parameter(Position=0, Mandatory=$true, ParameterSetName="AddUserToGroupInputFile")]
    [string[]]$GroupDistinguishedName,

    [parameter(Position=0, Mandatory=$true, ParameterSetName="AddUserToGroupInputFile")]
    [string]$UserSamAccountNameFilePath,

    [parameter(Position=0, Mandatory=$true, ParameterSetName="AddUserToGroup")]
    [switch]$AddUser,

    [parameter(Position=0, Mandatory=$true, ParameterSetName="RemoveUserFromGroup")]
    [switch]$RemoveUser

)


Add-Type -AssemblyName System.DirectoryServices.AccountManagement


function Get-ActiveDirectoryUserMembership
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$DistinguishedName
    )

    begin
    {
        $membership = @()
    }
    process
    {
        try
        {
            $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DistinguishedName")

            $propertyValueCollection = $directoryEntry.Properties["memberof"]

            foreach( $value in $propertyValueCollection.GetEnumerator() )
            {
                if( $value )
                {
                    try
                    {
                        $membership += Get-ActiveDirectoryGroup -DistinguishedName $value
                    }
                    catch
                    {
                        Write-Error "Error searching for user principal group $value. Exception: $($_.Exception)"
                    }
                }
            }
        }
        catch
        {
            Write-Error "Error searching for user memberships $DistinguishedName. Exception: $($_.Exception)"
        }
        finally
        {
            if( $directoryEntry )
            {
                $directoryEntry.Close()
                $directoryEntry.Dispose()
            }        
        }
    }
    end
    {
        $membership
    }
}

function Get-ActiveDirectoryGroupMember
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$DistinguishedName,
        [Parameter(Mandatory=$false)][switch]$Recurse
    )

    begin
    {
        $members = @()
    }
    process
    {
        try
        {
            $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DistinguishedName")

            $propertyValueCollection = $directoryEntry.Properties["member"]

            foreach( $value in $propertyValueCollection.GetEnumerator())
            {
                if( $value )
                {
                    try
                    {
                        $member = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$value")

                        if( $member.SchemaClassName -eq "group" )
                        {
                            if( $Recurse )
                            {
                                $members += Get-ActiveDirectoryGroupMember -DistinguishedName $member.distinguishedName -Recurse:$Recurse.IsPresent
                            }
                            else
                            {
                                $members += Get-ActiveDirectoryGroup -DistinguishedName $member.distinguishedName
                            }                    
                        }
                        else # user
                        {
                            $members += Get-ActiveDirectoryUser -DistinguishedName $member.distinguishedName
                        }
                    }
                    catch
                    {
                        Write-Error "Error searching for principal $value. Exception: $($_.Exception)"
                    }
                    finally
                    {
                        if( $member )
                        {
                            $member.Close()
                            $member.Dispose()
                        }
                    }
                }
            }
        }
        catch
        {
            Write-Error "Error searching for group principal $DistinguishedName. Exception: $($_.Exception)"
        }
        finally
        {
            if( $directoryEntry )
            {
                $directoryEntry.Close()
                $directoryEntry.Dispose()
            }        
        }
    }
    end
    {
        $members
    }
}

function Get-ActiveDirectoryUser
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false,position=0,ParameterSetName='DistinguishedName')]
        [string]$DistinguishedName,

        [Parameter(Mandatory=$false,position=0,ParameterSetName='SamAccountName')]
        [string]$SamAccountName,

        [Parameter(Mandatory=$false,position=0,ParameterSetName='Name')]
        [string]$Name,
    
        [Parameter(Mandatory=$false,position=0,ParameterSetName='UserPrincipalName')]
        [string]$UserPrincipalName
    )

    begin
    {
        $principal = $null
    }
    process
    {
        $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext( [System.DirectoryServices.AccountManagement.ContextType]::Domain)
    
        try
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                "DistinguishedName"
                {
                    $principal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::DistinguishedName, $DistinguishedName)
                }
                "SamAccountName"
                {
                    $principal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName, $SamAccountName)
                }
                "Name"
                {
                    $principal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::Name, $Name)
                }
                "UserPrincipalName"
                {
                    $principal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::UserPrincipalName, $UserPrincipalName)
                }
            }
        }
        catch
        {
            Write-Error "Error getting user principal by $($PSCmdlet.ParameterSetName). Exception: $($_.Exception)"

            throw $_.Exception    
        }
        finally
        {
            $context.Dispose()
        }
    }
    end
    {
        $principal
    }
}

function Get-ActiveDirectoryGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false,position=0,ParameterSetName='DistinguishedName')]
        [string]$DistinguishedName,

        [Parameter(Mandatory=$false,position=0,ParameterSetName='SamAccountName')]
        [string]$SamAccountName,

        [Parameter(Mandatory=$false,position=0,ParameterSetName='Name')]
        [string]$Name
    )

    begin
    {
        $principal = $null
    }
    process
    {
        $context = New-Object System.DirectoryServices.AccountManagement.PrincipalContext( [System.DirectoryServices.AccountManagement.ContextType]::Domain)
    
        try
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                "DistinguishedName"
                {
                    $principal = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::DistinguishedName, $DistinguishedName)
                }
                "SamAccountName"
                {
                    $principal = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName, $SamAccountName)
                }
                "Name"
                {
                    $principal = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity( $context, [System.DirectoryServices.AccountManagement.IdentityType]::Name, $Name)
                }
            }
        }
        catch
        {
            Write-Error "Error getting group principal by $($PSCmdlet.ParameterSetName). Exception: $($_.Exception)"

            throw $_.Exception    
        }
        finally
        {
            $context.Dispose()
        }
    }
    end
    {
        $principal
    }
}

function Add-PrincipalToGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$PrincipalDistinguishedName,
        [Parameter(Mandatory=$true)][string]$GroupDistinguishedName
    )

    try
    {
        Write-Verbose -Message "$($PSCmdlet.CommandRuntime) - Attempting to add $PrincipalDistinguishedName to $GroupDistinguishedName"

        $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$GroupDistinguishedName")
        $directoryEntry.Properties["member"].Add($PrincipalDistinguishedName) | Out-Null
        $directoryEntry.CommitChanges()
        
        Write-Verbose -Message "$($PSCmdlet.CommandRuntime) - Successfully added $PrincipalDistinguishedName to $GroupDistinguishedName"
    }
    catch
    {
        Write-Error "Error adding principal $PrincipalDistinguishedName to group $GroupDistinguishedName.  Exception: $($_.Exception)"

        throw $_.Exception         
    }
    finally
    {
        if( $directoryEntry )
        {
            $directoryEntry.Close()
            $directoryEntry.Dispose()
        }
    }
}

function Remove-PrincipalFromGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$PrincipalDistinguishedName,
        [Parameter(Mandatory=$true)][string]$GroupDistinguishedName
    )

    try
    {
        Write-Verbose -Message "$($PSCmdlet.CommandRuntime) - Attempting to remove $PrincipalDistinguishedName from $GroupDistinguishedName"

        $directoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$GroupDistinguishedName")
        $directoryEntry.Properties["member"].Remove($PrincipalDistinguishedName) | Out-Null
        $directoryEntry.CommitChanges()

        Write-Verbose -Message "$($PSCmdlet.CommandRuntime) - Successfully removed $PrincipalDistinguishedName from $GroupDistinguishedName"
    }
    catch
    {
        Write-Error "Error removing principal $PrincipalDistinguishedName from group $GroupDistinguishedName.  Exception: $($_.Exception)"

        throw $_.Exception         
    }
    finally
    {
        if( $directoryEntry )
        {
            $directoryEntry.Close()
            $directoryEntry.Dispose()
        }
    }
}


# Test getting Active Directory User

    # Get-ActiveDirectoryUser -DistinguishedName "CN=Adam Barr,OU=Demo Users,DC=contoso,DC=com"
    # Get-ActiveDirectoryUser -SamAccountName "adamb"
    # Get-ActiveDirectoryUser -Name "Adam Barr"
    # Get-ActiveDirectoryUser -UserPrincipalName "adamb@contoso.com"


# Test getting Active Directory Group

    # Get-ActiveDirectoryGroup -DistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com"
    # Get-ActiveDirectoryGroup -SamAccountName "DomainGroup"
    # Get-ActiveDirectoryGroup -Name "DomainGroup"


# Test adding & removing a user to/from a domain group

    # Add-PrincipalToGroup -PrincipalDistinguishedName "CN=Adam Barr,OU=Demo Users,DC=contoso,DC=com" -GroupDistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com"
    # Add-PrincipalToGroup -PrincipalDistinguishedName "CN=IT Operations,OU=Demo Users,DC=contoso,DC=com" -GroupDistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com"


# Test adding & removing a group to/from a domain group

    # Remove-PrincipalFromGroup -PrincipalDistinguishedName "CN=Adam Barr,OU=Demo Users,DC=contoso,DC=com" -GroupDistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com"
    # Remove-PrincipalFromGroup -PrincipalDistinguishedName "CN=IT Operations,OU=Demo Users,DC=contoso,DC=com" -GroupDistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com"


# Test listing group members

    # Get-ActiveDirectoryGroupMember -DistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com" | SELECT SamAccountName
    # Get-ActiveDirectoryGroupMember -DistinguishedName "CN=DomainGroup,OU=2013ServiceAccounts,DC=contoso,DC=com" -Recurse | SELECT DistinguishedName


# Test listing user memberships

    # Get-ActiveDirectoryUserMembership -DistinguishedName "CN=Adam Barr,OU=Demo Users,DC=contoso,DC=com" | SELECT DistinguishedName





# Request was to list user group membership

    if( $PSCmdlet.ParameterSetName -eq "GetUserMembership")
    {
        $UserDistinguishedName | % { $userDn = $_; Get-ActiveDirectoryUserMembership -DistinguishedName $_ } | FT @{ Name="UserDistinguishedName"; Expression={$userDn}}, @{ Name="GroupDistinguishedName"; Expression={$_.DistinguishedName}} -AutoSize
        return
    }


# request was to add or remove user(s) to/from a domain group(s)


    if( $PSCmdlet.ParameterSetName -eq "AddUserToGroup" -or $PSCmdlet.ParameterSetName -eq "RemoveUserFromGroup" )
    {
        $UserDistinguishedName | % { 
    
            $userDn = $_

            if( $PSCmdlet.ParameterSetName -eq "AddUserToGroup" )
            {
                $GroupDistinguishedName | % { Add-PrincipalToGroup -PrincipalDistinguishedName $userDn -GroupDistinguishedName $_  }
            }
            elseif( $PSCmdlet.ParameterSetName -eq "RemoveUserFromGroup" )
            {
                $GroupDistinguishedName | % { Remove-PrincipalFromGroup -PrincipalDistinguishedName $userDn -GroupDistinguishedName $_  }
            }
        }

        return
    }



# request was to add or remove user(s) listed in a txt file to/from a domain group(s)

    if( $PSCmdlet.ParameterSetName -eq "AddUserToGroupInputFile" )
    {
        if( Test-Path -Path $UserSamAccountNameFilePath -PathType Leaf )
        {
            Get-Content -Path $UserSamAccountNameFilePath | % {
        
                $samAccountName = $_.ToString().Trim()

                if( $samAccountName )
                {
                    $principal = Get-ActiveDirectoryUser -SamAccountName $samAccountName

                    if( $principal )
                    {
                        $GroupDistinguishedName | % { Add-PrincipalToGroup -PrincipalDistinguishedName $principal.DistinguishedName -GroupDistinguishedName $_  }
                    }
                    else
                    {
                        Write-Warning "User not found: $samAccountName"
                    }
                }
        
            }
        }
        else
        {
            Write-Error "File not found: $UserSamAccountNameFilePath"
        }
   
        return
    }

