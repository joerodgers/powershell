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
 01-25-2017 - Created
 07-18-2017 - Rewritten to handle complex transformations
 07-21-2017 - Updated to improve perf for large AD query result sets
  

    This script will sync Active Directory attribute values to SharePoint Online's User Profile service.  The 
    sync method in the script is using the bulk user profile API.  At a high level, the sync is performed by 
    generating a JSON file with the property values for each user in AD.  The JSON file is uploaded to a team 
    site library.  The code creates a import job, using the JSON file as input.  In testing, it takes ~20 minutes 
    for the the import job to start updating profiles.

    More Details:
    https://support.microsoft.com/en-us/help/3168272/information-about-user-profile-synchronization-in-sharepoint-online

==============================================================#>


$CSOMAssemblyPath = "C:\Microsoft.SharePointOnline.CSOM.16.1.6008.1200\lib\net45"


# connection info for tenant admin site

    $tenantAdminUrl  = "https://contoso-admin.sharepoint.com"
    $tenantUserName  = "admin@contoso.onmicrosoft.com"
    $tenantSecurePwd =  ConvertTo-SecureString 'pass@word1' -AsPlainText -Force


# connection info for the upload site

    $uploadSiteLibrary   = "Documents"
    $uploadSiteUrl       = "https://contoso.sharepoint.com/sites/teamsite"
    $uploadSiteUserName  = "admin@contoso.onmicrosoft.com"
    $uploadSiteSecurePwd =  ConvertTo-SecureString 'pass@word1' -AsPlainText -Force


# DELETE BEFORE PUBLISHING

    $tenantAdminUrl  = "https://josrod-admin.sharepoint.com"
    $tenantUserName  = "joe.rodgers@josrod.onmicrosoft.com"
    $tenantSecurePwd =  ConvertTo-SecureString 'pass@word45' -AsPlainText -Force

    $uploadSiteLibrary   = "Documents"
    $uploadSiteUrl       = "https://josrod.sharepoint.com/sites/teamsite"
    $uploadSiteUserName  = "joe.rodgers@josrod.onmicrosoft.com"
    $uploadSiteSecurePwd =  ConvertTo-SecureString 'pass@word45' -AsPlainText -Force


# (object type is a user) and (not disabled) and (has an email value) and (CN starts with a* or n*)

    $ldapFilter = "(& (objectClass=user) (!userAccountControl:1.2.840.113556.1.4.803:=2) (cn=a*))"


# what field is the unique identifier for the users in SPO. Keep this in sync with the $userIdType variable below

    $userIdentifierAttributeName = "userPrincipalName"
    $userIdentifierAttributeName = "mail"


# this is a mapping of the AD property name to the SPO property name.  You can define up to 500 property mappings in one import.

    $propertyMappingObjects =  @( 

        # department L5/L6/L7/L8
        [PSCustomObject] @{
            Name                       = "Department Hierarchy Import" 
            SourcePropertyName         = $null 
            TargetPropertyName         = $null
            DataTransformationFunction = "Export-DepartmentNumber"
        },

        # Work Location Type
        [PSCustomObject] @{
            Name                       = "Work Location Type Import" 
            SourcePropertyName         = $null
            TargetPropertyName         = $null
            DataTransformationFunction = "Export-WorkLocationType"
        },    

        # Employee Number
        [PSCustomObject] @{
            Name                       = "Employee Number Import" 
            SourcePropertyName         = "mailNickName"
            TargetPropertyName         = "EmployeeNumber"
            DataTransformationFunction = $null
        },    

        # Workforce type
        [PSCustomObject] @{
            Name                       = "Workforce Type Import" 
            SourcePropertyName         = $null 
            TargetPropertyName         = $null
            DataTransformationFunction = "Export-WorkforceType"
        },    

        # Office
        [PSCustomObject] @{
            Name                       = "Office Import" 
            SourcePropertyName         = "physicaldeliveryofficename" 
            TargetPropertyName         = "Office"
            DataTransformationFunction = $null
        },    

        # Work Location
        [PSCustomObject] @{
            Name                       = "Work Location Import" 
            SourcePropertyName         = "l" 
            TargetPropertyName         = "WorkLocation"
            DataTransformationFunction = $null
        },    

        # mobile
        [PSCustomObject] @{
            Name                       = "Mobile Phone Import" 
            SourcePropertyName         = "mobile" 
            TargetPropertyName         = "CellPhone"
            DataTransformationFunction = $null
        },    

        # state/province
        [PSCustomObject] @{
            Name                       = "State/Province Import" 
            SourcePropertyName         = "st" 
            TargetPropertyName         = "StateProvince"
            DataTransformationFunction = $null
        },    
 
        # country
        [PSCustomObject] @{
            Name                       = "Country Import" 
            SourcePropertyName         = "Country" 
            TargetPropertyName         = "Country"
            DataTransformationFunction = $null
        },    

        # telephone number
        [PSCustomObject] @{
            Name                       = "Telephone Number Import" 
            SourcePropertyName         = "telephoneNumber" 
            TargetPropertyName         = "WorkPhone"
            DataTransformationFunction = $null
        }
    )

# this defines this var as script wide variable so functions access the value

    $TRACE_LOG_PATH = "C:\_temp\$($MyInvocation.MyCommand.Name)_$($(Get-Date).ToString('yyyyMMdd')).csv"


# this is the location that the .txt (JSON) file will be temporarly created, it the upload succeeds the file is deleted

    $tempFileDirectory = "C:\_Temp"


# fill these values in if you want to send an email if any critical failures occur during processing

    $notificationEmailTo   = @()
    $notificationEmailFrom = ""
    $notificationEmailSMTP = ""


# if the upload fails, you can attempt to re-run this sync script with an exisiting JSON file, otherwise leave this null or blank

    $existingJSONFilePath = $null # "e:\json\UserProfilePropertySync-2017-07-12_03.19.27.txt"
    
    # If you are using an existing JSON file, you have to define the property mappings for the timer job.
    # If you are using not using an existing JSON file, these settings will be built on the fly and anything here will be ignored.
    $propertyFieldMappings = New-Object -Type "System.Collections.Generic.Dictionary[String,String]"
    #$propertyFieldMappings.Add( "CostCenter2",   "CostCenter"    )
    #$propertyFieldMappings.Add( "WorkforceType", "WorkforceType" )
    #$propertyFieldMappings.Add( "DepartmentL5",  "DepartmentL5"  )
    #$propertyFieldMappings.Add( "DepartmentL6",  "DepartmentL6"  )
    #$propertyFieldMappings.Add( "DepartmentL7",  "DepartmentL7"  )
    #$propertyFieldMappings.Add( "DepartmentL8",  "DepartmentL8"  )



<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
$CorrelationId = [System.Guid]::NewGuid()

function Write-TraceLogEntry
{
    <#
    .Synopsis

       Writes the specified message to a log file.  The log file is path is either specified in the cmdlet or by the global variable $Global:TRACE_LOG_PATH

    .EXAMPLE

        Write-TraceLogEntry -Message "Log this to the file" -TraceLevel "Verbose" -Correlation "bad3f4e2-27c9-43ec-af16-e77fe9001ba6"

    .EXAMPLE

        Write-TraceLogEntry -Message "Log this to the file" -TraceLevel "Verbose" -Correlation "bad3f4e2-27c9-43ec-af16-e77fe9001ba6" -LogPath "C:\Temp\LogFile.log"
    #>

    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$Message,
        [parameter(Mandatory=$true)][ValidateSet("Verbose", "Low", "Medium", "High", "Critical")][string]$TraceLevel,
        [parameter(Mandatory=$false)][string]$LogPath = $TRACE_LOG_PATH,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
    }
    process
    {
        $logEntry = [PSCustomObject] @{
            Date          = $(Get-Date).ToString("G")
            TraceLevel    = $TraceLevel
            Message       = $Message
            CorrelationId = $CorrelationId
        }

        try
        {
            if($LogPath)
            {
                $mutex = New-Object System.Threading.Mutex($false, "Trace Log Mutex")
                $mutexAcquired = $mutex.WaitOne(5000) # wait 5 seconds before timing out
                
                if( $mutexAcquired )
                {
                    if( -not (Test-Path -Path $LogPath -PathType Leaf) )
                    {
                        ($logEntry | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[0] | Out-File -FilePath $LogPath -Append
                    }

                    # log the message to the log
                    ($logEntry | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1] | Out-File -FilePath $LogPath -Append
                }
                else
                {
                    Write-Error "Mutex could not be aquired."
                    $_
                }
            }

            switch( $TraceLevel )
            {
                "Critical"
                {
                    Write-TraceLogEntry -Message "Flipped critical flag to true" -TraceLevel Verbose

                    Write-Host ($logEntry | Select-Object Date, TraceLevel, Message | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1] -ForegroundColor Red
                }
                "High"
                {
                    Write-Host ($logEntry | Select-Object Date, TraceLevel, Message | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1] -ForegroundColor Yellow
                }
                "Medium"
                {
                    Write-Host ($logEntry | Select-Object Date, TraceLevel, Message | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1]
                }
                "Low"
                {
                    Write-Host ($logEntry | Select-Object Date, TraceLevel, Message | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1]
                }
                "Verbose"
                {
                    # don't log verbose messages to the screen
                    # Write-Host ($logEntry | SELECT Date, TraceLevel, Message | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation)[1]
                }
            }

        }
        catch
        {
            Write-Error "Unable to log message. Exception: $($_.Exception)"
        }
        finally
        {
            if( $mutexAcquired -and $mutex )
            {
                $mutex.ReleaseMutex()
                $mutex.Dispose()
            }
        }
    }
    end
    {
    }
}

function Send-EmailNotification
{
    <#
    .Synopsis

       Sends an email.  If either To, From or Subject fields are empty, the contents of the mail are written to the trace file.

    .EXAMPLE

        Send-EmailNotification -To "john.doe@contoso.com" -From "jane.doe@contoso.com" -Subject "Script suffered a critical failure." -Body "HTML BODY" -SmtpServer "mail.contoso.com"
   
    #>

    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$false)][string[]]$To,
        [parameter(Mandatory=$false)][string]$From,
        [parameter(Mandatory=$false)][string[]]$Cc,
        [parameter(Mandatory=$false)][string]$Subject,
        [parameter(Mandatory=$false)][string]$Body,
        [parameter(Mandatory=$false)][string]$SMTPServer,
        [parameter(Mandatory=$false)][string[]]$Attachments,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Sending Email Notification" -TraceLevel Verbose -CorrelationId $CorrelationId
        
        if($To -and $From -and $Subject)
        {
            $params = @{
                To         = $To
                From       = $From
                Subject    = $Subject
            }

            if($Body)
            {
                $params.Add("Body", $Body)
                $params.Add("BodyAsHtml", $true)
            }

            if($SMTPServer)
            {
                $params.Add("SMTPServer", $SMTPServer)
            }

            if($Attachments)
            {
                $params.Add("Attachments", $Attachments)
            }

            if($Cc)
            {
                $params.Add("Cc", $cc)
            }

            Send-MailMessage @params
            
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Email notification sent" -TraceLevel Verbose -CorrelationId $CorrelationId
        }
        else
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Email notification not sent, missing required email params." -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    } 
}

function Get-ActiveDirectoryUsers
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$Filter,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId

        $domain   = New-Object System.DirectoryServices.DirectoryEntry
        $searcher = New-Object System.DirectoryServices.DirectorySearcher 
        $searchResults = $null
    }
    process
    {
        try 
        {
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.SearchRoot  = $domain
            $searcher.PageSize    = 1000
            $searcher.Filter      = $Filter 
            $searcher.SearchScope = "Subtree"    
            
            $stopWatch = Measure-Command -Expression { $searchResults = $searcher.FindAll() } 

            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - AD query execution minutes: $($stopWatch.TotalMinutes)" -TraceLevel Verbose -CorrelationId $CorrelationId
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Exception querying AD. Exception: $($_.Exception)" -TraceLevel Critical -CorrelationId $CorrelationId
            throw $_.Excpetion
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
        $searchResults
    }
}

function Import-ClientSideObjectModelAssemblies
{
    <#
    .Synopsis

       Loads all of the SharePoint Online Assemblies into memory.  Returns TRUE if assembly loading was succesful, otherwise FALSE.

    .EXAMPLE

        Import-ClientSideObjectModelAssemblies -Path "C:\O365_CSOM\Microsoft.SharePointOnline.CSOM.16.1.5715.1200\lib\net45"

    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AssemblyPath,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
        $success = $true
    }
    process
    {
        try
        {
            Get-ChildItem -Path $AssemblyPath -ErrorAction Stop | Where-Object { $_.Name -match "Microsoft.*.dll" -and $_.Name -notmatch ".Runtime.Windows.dll" } | ForEach-Object { 
                
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Loading Assembly: $($_.FullName)" -TraceLevel Verbose -CorrelationId $CorrelationId
                
                [System.Reflection.Assembly]::LoadFrom( $_.FullName ) | Out-Null
            }
        }
        catch
        {
            $success = $false
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Failed to load CSOM Assemblies"             -TraceLevel Critical -CorrelationId $CorrelationId
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Error Details:     $($_.Exception.Message)" -TraceLevel Critical -CorrelationId $CorrelationId
            return
        }
    }
    end
    {
        $success
    }
}

function Export-GenericActiveDirectoryProperty
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$false)][string]$SourcePropertyName,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        try
        {
            # pull the user identifier property
            $userIdentifier = $ActiveDirectoryUser.Properties[$IdentifierPropertyName][0] 

            # ensure the user has been initialized in the export mappings
            if( -not $ActiveDirectoryUserExportMappings.ContainsKey( $userIdentifier ) )
            {
                $ActiveDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
            }

            # pull the current property mapping values for the user
            $profileValues = $ActiveDirectoryUserExportMappings[$userIdentifier]

            # get the property value from the Active Directory User Object
            $activeDirectoryPropertyValue = $ActiveDirectoryUser.Properties[$SourcePropertyName][0]

            if($activeDirectoryPropertyValue)
            {
                if( $activeDirectoryPropertyValue.GetType() -eq [Byte[]] )
                {
                    # convert from byte[] to string
                    $activeDirectoryPropertyValue = $([System.Text.Encoding]::ASCII.GetString($activeDirectoryPropertyValue, 0, $byteArray.Count))
                }

                # trim off excess spaces from the string value
                $activeDirectoryPropertyValue = $activeDirectoryPropertyValue.Trim()

                if( $activeDirectoryPropertyValue)
                {
                    # Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Property Name: $SourcePropertyName, Property Value: $activeDirectoryPropertyValue" -TraceLevel Verbose -CorrelationId $CorrelationId
                    $profileValues.Add( $SourcePropertyName, $activeDirectoryPropertyValue)
                }
                else 
                {
                    # $profileValues.Add( $activeDirectoryPropertyName, $null )
                }
            }
            else
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Export property '$SourcePropertyName' is empty for user '$userIdentifier'." -TraceLevel Verbose -CorrelationId $CorrelationId
            }
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unable to export property '$SourcePropertyName' for user '$userIdentifier'. Exception: $($_.Exception)" -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}

function Export-DepartmentNumber
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$true)][System.Collections.Generic.Dictionary[String,String]]$PropertyFieldMap,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        try
        {
            # pull the user identifier property
            $userIdentifier = $ActiveDirectoryUser.Properties[$IdentifierPropertyName][0] 

            # ensure the user has been initialized in the export mappings
            if( -not $ActiveDirectoryUserExportMappings.ContainsKey( $userIdentifier ) )
            {
                $ActiveDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
            }

            # pull the current property mapping values for the user
            $profileValues = $ActiveDirectoryUserExportMappings[$userIdentifier]
            
            foreach( $departmentNumber in $ActiveDirectoryUser.Properties["DepartmentNumber"] )
            {
                if( $departmentNumber )
                {
                    $level = $departmentNumber.Substring(0,2)
                
                    switch( $level )
                    {
                        "L5"
                        {
                            $profileValues.Add("DepartmentL5", $departmentNumber.Substring(2).Trim() )

                            if( -not $PropertyFieldMap.ContainsKey( "DepartmentL5" ) )
                            {
                                $PropertyFieldMap.Add( "DepartmentL5", "DepartmentL5" )
                            }
                        }
                        "L6"
                        {
                            $profileValues.Add("DepartmentL6", $departmentNumber.Substring(2).Trim() )

                            if( -not $PropertyFieldMap.ContainsKey( "DepartmentL6" ) )
                            {
                                $PropertyFieldMap.Add( "DepartmentL6", "DepartmentL6" )
                            }
                        }
                        "L7"
                        {
                            $profileValues.Add("DepartmentL7", $departmentNumber.Substring(2).Trim() )

                            if( -not $PropertyFieldMap.ContainsKey( "DepartmentL7" ) )
                            {
                                $PropertyFieldMap.Add( "DepartmentL7", "DepartmentL7" )
                            }
                        }
                        "L8"
                        {
                            $profileValues.Add("DepartmentL8", $departmentNumber.Substring(2).Trim() )

                            if( -not $PropertyFieldMap.ContainsKey( "DepartmentL8" ) )
                            {
                                $PropertyFieldMap.Add( "DepartmentL8", "DepartmentL8" )
                            }
                        }
                        default
                        {
                            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unexepcted 'DepartmentNumber' value for user '$userIdentifier'. Value: $($departmentNumber.Substring(0,2))" -TraceLevel Verbose -CorrelationId $CorrelationId
                        }
                    }
                }
            }
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unable to export property 'DepartmentNumber' for user '$userIdentifier'. Exception: $($_.Exception)" -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}

function Export-WorkforceType
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$true)][System.Collections.Generic.Dictionary[String,String]]$PropertyFieldMap,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        try
        {
            # pull the user identifier property
            $userIdentifier = $ActiveDirectoryUser.Properties[$IdentifierPropertyName][0] 

            # ensure the user has been initialized in the export mappings
            if( -not $ActiveDirectoryUserExportMappings.ContainsKey( $userIdentifier ) )
            {
                $ActiveDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
            }

            # pull the current property mapping values for the user
            $profileValues = $ActiveDirectoryUserExportMappings[$userIdentifier]
            
            # If the first character of the AD.EmployeeNumber is 'A', then the Workforce type is 'FTE'. 
            # If the First letter of the Employee Number is 'N', then workforce type is 'Contingent'. 
            # Otherwise, the Workforce type is 'Other'            
            
            if( -not $profileValues.ContainsKey("WorkforceType") )
            {
                $employeeNumber = $ActiveDirectoryUser.Properties["CN"][0]

                if( $employeeNumber )
                {
                    $firstLetter = $employeeNumber.Substring(0,1)
                    
                    if( -not $PropertyFieldMap.ContainsKey( "WorkforceType" ) )
                    {
                        $PropertyFieldMap.Add( "WorkforceType", "WorkforceType" )
                    }

                    switch( $firstLetter )
                    {
                        "A"
                        {
                            $profileValues.Add("WorkforceType", "FTE" )
                        }
                        "N"
                        {
                            $profileValues.Add("WorkforceType", "Contengent" )
                        }
                        default
                        {
                            $profileValues.Add("WorkforceType", "Other" )
                        }
                    }
                }
            }
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unable to export property 'WorkforceType' for user '$userIdentifier'. Exception: $($_.Exception)" -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}

function Export-WorkLocationType
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$true)][System.Collections.Generic.Dictionary[String,String]]$PropertyFieldMap,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {        
        try
        {
            # pull the user identifier property
            $userIdentifier = $ActiveDirectoryUser.Properties[$IdentifierPropertyName][0] 

            # ensure the user has been initialized in the export mappings
            if( -not $ActiveDirectoryUserExportMappings.ContainsKey( $userIdentifier ) )
            {
                $ActiveDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
            }

            # pull the current property mapping values for the user
            $profileValues = $ActiveDirectoryUserExportMappings[$userIdentifier]

            if( $profileValues -and -not $profileValues.ContainsKey("WorkLocationType"))
            {
                $physicaldeliveryofficename = $ActiveDirectoryUser.Properties["physicaldeliveryofficename"][0]

                if( $physicaldeliveryofficename )
                {
                    # If the last character of USP.Office = 'H', then this value is 'Teleworker'. 
                    # If no letter or 'S', then 'Office'
                    $lastLetter = $physicaldeliveryofficename.Substring($physicaldeliveryofficename.Length-1)

                    if( -not $PropertyFieldMap.ContainsKey( "WorkLocationType" ) )
                    {
                        $PropertyFieldMap.Add( "WorkLocationType", "WorkLocationType" )
                    }

                    switch( $lastLetter )
                    {
                        "H"
                        {
                            $profileValues.Add("WorkLocationType", "Teleworker" )
                        }
                        default
                        {
                            $profileValues.Add("WorkLocationType", "Office" )
                        }
                    }
                }
            }
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unable to export property 'WorkLocationType' for user '$userIdentifier'. Exception: $($_.Exception)" -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}

function Invoke-ScriptBlockPropertyExport
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$true)][string]$SourcePropertyName,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$true)][ScriptBlock]$ScriptBlock,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        try 
        {
            $userIdentifier = $ActiveDirectoryUser.Properties[$IdentifierPropertyName][0] 

            if( -not $ActiveDirectoryUserExportMappings.ContainsKey( $userIdentifier ) )
            {
                $ActiveDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
            }

            $profileValues = $ActiveDirectoryUserExportMappings[$userIdentifier]

            if( -not $profileValues.ContainsKey($SourcePropertyName) )
            {
                $profileValues.Add( $SourcePropertyName , $ScriptBlock.Invoke($null) )
            }
        }
        catch 
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unable to export property '$SourcePropertyName' for user '$userIdentifier'. Exception: $($_.Exception)" -TraceLevel High -CorrelationId $CorrelationId
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}

function Publish-File
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$false)][Microsoft.SharePoint.Client.ClientContext]$ClientContext,
        [parameter(Mandatory=$false)][string]$ListTitle = "Documents",
        [parameter(Mandatory=$false)][string]$FilePath,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
        $fileUrl = ""
    }
    process
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Attempting to upload file: $FilePath to $ListTitle" -TraceLevel Verbose -CorrelationId $CorrelationId

        try
        {
            $fi = New-Object System.IO.FileInfo($FilePath)
            $fs = New-Object System.IO.FileStream($fi.FullName, "Open")

            $list = $ClientContext.Web.Lists.GetByTitle($ListTitle)
            $ClientContext.Load($ClientContext.Site)
            $ClientContext.Load($ClientContext.Web)
            $ClientContext.Load($list.RootFolder)
            Invoke-ClientContextWithRetry -ClientContext $ClientContext -CorrelationId $CorrelationId


            # https://tenant.sharepoint.com/sites/teamsite/subsite
            $webUri = New-Object System.Uri( $ClientContext.Web.Url) 
            
            # https://tenant.sharepoint.com
            $serverUri = New-Object System.Uri( "$($webUri.Scheme)://$($webUri.Host)" )

            # /sites/teamsite/subsite/documents/sample.txt
            $listRelativeUri = New-Object System.Uri( "$($list.RootFolder.ServerRelativeUrl)/$($fi.Name)", [System.UriKind]::Relative) 
     
            # https://tenant.sharepoint.com/sites/teamsite/subsite/documents/sample.txt
            $fileUri = New-Object System.Uri( $serverUri, $listRelativeUri )

            [Microsoft.SharePoint.Client.File]::SaveBinaryDirect($clientContext, $fileUri.AbsolutePath, $fs, $true)

            $fileUrl = $fileUri.ToString()

            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Upload of $($fileUri.AbsolutePath) complete." -TraceLevel Verbose -CorrelationId $CorrelationId
        }
        catch
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Upload of $($fileUri.AbsolutePath) failed. Error: $($_.Exception)" -TraceLevel Critical -CorrelationId $CorrelationId
        }
        finally
        {
            if($fs)
            {
                $fs.Close()
                $fs.Dispose()
            }
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
        $fileUrl
    }
}

function Invoke-ClientContextWithRetry
{
    <#
    .Synopsis

       Executes a CSOM query.  If the query fails to execute because of a WebException, the code will retry the query. The number of retries and delays 
       between the query are either provided to the cmdlet or the default values are used.

    .EXAMPLE

        Invoke-ClientContextWithRetry -ClientContext $ClientContext -CorrelationId "bad3f4e2-27c9-43ec-af16-e77fe9001ba6"

    .EXAMPLE

        Invoke-ClientContextWithRetry -ClientContext $ClientContext -Delay 30 -Retry 2
    #>

    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$ClientContext,
        [parameter(Mandatory=$false)][int]$Delay = $DEFAULT_CONTEXT_QUERY_RETRY_DELAY_SECONDS,
        [parameter(Mandatory=$false)][int]$RetryAttempts = $DEFAULT_CONTEXT_QUERY_RETRY_ATTEMPTS,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
        $attemps = 1
    }
    process
    {
        do
        {
            try
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Context Url: $($ClientContext.Url)" -TraceLevel Verbose -CorrelationId $CorrelationId

                $ClientContext.ExecuteQuery()
                return
            }
            catch [System.Net.WebException]
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution attempt: $attemps failed" -TraceLevel High -CorrelationId $CorrelationId
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Exception Info: $($_.Exception)"    -TraceLevel High -CorrelationId $CorrelationId

                $attemps++
                Start-Sleep -Seconds $Delay
            }
            catch
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Unexpected execution exception: $($_.Exception)" -TraceLevel Critical -CorrelationId $CorrelationId
                throw $_.Exception
            }
        }
        while( $attemps -lt $RetryAttempts )

        # we never had a successful execution
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Retry threshold ($RetryAttempts) exceeded." -TraceLevel High -CorrelationId $CorrelationId
        throw [System.Exception] "Invoke-ContextQuery attempts exceeded."
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    } 
}

function New-ClientContextWithRetry
{
    <#
    .Synopsis

       Creates a new client context object and loads the Site and Web objects.

    .EXAMPLE

        $tenantContext = New-ClientContextWithRetry -ContextUrl "https://tenant-admin.sharepoint.com" -Credential $tenatAdminCredential 

    .EXAMPLE

        $tenantContext = New-ClientContextWithRetry -ContextUrl "https://tenant.sharepoint.com/sites/teamsite" -Credential $tenatAdminCredential 
    #>

    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$ContextUrl,
        [Parameter(Mandatory=$true, ParameterSetName = "PSCredential")][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$true, ParameterSetName = "SPOCredential")][Microsoft.SharePoint.Client.SharePointOnlineCredentials]$SharePointOnlineCredential,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )
    
    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
        $context = $null
    }
    process
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Creating Context: $ContextUrl" -TraceLevel Verbose -CorrelationId $CorrelationId

        $context = New-Object Microsoft.SharePoint.Client.ClientContext($ContextUrl)
        switch ($psCmdlet.ParameterSetName) 
        {
            "PSCredential"
            {
                $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)
            }
            "SPOCredential"
            {
                $context.Credentials = $SharePointOnlineCredential
            }
        }

        $context.Load($context.Web)
        $context.Load($context.Site)

        Invoke-ClientContextWithRetry -ClientContext $context -CorrelationId $CorrelationId

        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Context Created: $ContextUrl" -TraceLevel Verbose  -CorrelationId $CorrelationId
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
        $context
    }
}

function Invoke-MappingFunctionsForActiveDirectoryUser
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$IdentifierPropertyName,
        [parameter(Mandatory=$true)][object]$ActiveDirectoryUser,
        [parameter(Mandatory=$true)][object[]]$PropertyMappingObjects,
        [parameter(Mandatory=$true)][HashTable]$ActiveDirectoryUserExportMappings,
        [parameter(Mandatory=$true)][System.Collections.Generic.Dictionary[String,String]]$PropertyFieldMap,
        [parameter(Mandatory=$false)][System.Guid]$CorrelationId = [System.Guid]::NewGuid()
    )

    begin
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Execution." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
    process
    {
        foreach( $propertyMappingObject in $PropertyMappingObjects )
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Building property mappings for '$($propertyMappingObject.Name)'" -TraceLevel Verbose -CorrelationId $CorrelationId

            # check if this is a simple 1:1 mapping
            if( $propertyMappingObject.SourcePropertyName -and $propertyMappingObject.TargetPropertyName -and -not $propertyMappingObject.DataTransformationFunction )
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - $($propertyMappingObject.Name) is running under Export-GenericActiveDirectoryProperty." -TraceLevel Verbose -CorrelationId $CorrelationId

                if( -not $PropertyFieldMap.ContainsKey( $propertyMappingObject.SourcePropertyName ) )
                {
                    # add the 1:1 (source:target) property mappings to our hash for the timerjob
                    $PropertyFieldMap.Add( $propertyMappingObject.SourcePropertyName, $propertyMappingObject.TargetPropertyName )
                }

                Export-GenericActiveDirectoryProperty `
                    -IdentifierPropertyName            $IdentifierPropertyName `
                    -ActiveDirectoryUser               $ActiveDirectoryUser `
                    -SourcePropertyName                $propertyMappingObject.SourcePropertyName `
                    -ActiveDirectoryUserExportMappings $ActiveDirectoryUserExportMappings `
                    -CorrelationId                     $CorrelationId
            }
            elseif( $propertyMappingObject.SourcePropertyName -and $propertyMappingObject.TargetPropertyName -and $propertyMappingObject.DataTransformationFunction.GetType() -eq [System.Management.Automation.ScriptBlock])
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - $($propertyMappingObject.Name) is running under Invoke-ScriptBlockPropertyExport." -TraceLevel Verbose -CorrelationId $CorrelationId

                # this is simple static transformation, each property mapping maps to the value returned by a script block

                if( -not $PropertyFieldMap.ContainsKey( $propertyMappingObject.SourcePropertyName ) )
                {
                    # add the property mappings to our hash for the timerjob
                    $PropertyFieldMap.Add( $propertyMappingObject.SourcePropertyName, $propertyMappingObject.TargetPropertyName )
                }
                
                Invoke-ScriptBlockPropertyExport `
                    -IdentifierPropertyName            $IdentifierPropertyName `
                    -ActiveDirectoryUser               $ActiveDirectoryUser `
                    -SourcePropertyName                $propertyMappingObject.SourcePropertyName `
                    -ActiveDirectoryUserExportMappings $ActiveDirectoryUserExportMappings `
                    -ScriptBlock                       $propertyMappingObject.DataTransformationFunction `
                    -CorrelationId                     $CorrelationId
            
            }
            elseif($propertyMappingObject.DataTransformationFunction -and $propertyMappingObject.DataTransformationFunction.GetType() -eq [string])
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - $($propertyMappingObject.Name) is running under $($propertyMappingObject.DataTransformationFunction)." -TraceLevel Verbose -CorrelationId $CorrelationId

                # this is complex transformation that has a specific function that does the transformation. These function are effectively an implementation of a standard
                # interface that all custom transformation functions should adhere to.

                # complex transformation
                & $propertyMappingObject.DataTransformationFunction `
                    -IdentifierPropertyName            $IdentifierPropertyName `
                    -ActiveDirectoryUser               $ActiveDirectoryUser `
                    -ActiveDirectoryUserExportMappings $ActiveDirectoryUserExportMappings `
                    -PropertyFieldMap                  $PropertyFieldMap `
                    -CorrelationId                     $CorrelationId
            }
            else 
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - $($propertyMappingObject.Name) did not map to an export function." -TraceLevel High -CorrelationId $CorrelationId
            }        
        }
    }
    end
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Execution Ended." -TraceLevel Verbose -CorrelationId $CorrelationId
    }
}


# load assemblies and modules
if(-not $(Import-ClientSideObjectModelAssemblies -AssemblyPath $CSOMAssemblyPath -CorrelationId $CorrelationId) )
{
    Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Required CSOM assemblies failed to load." -TraceLevel Critical -CorrelationId $CorrelationId
    Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -Body "See $TRACE_LOG_PATH on $env:COMPUTERNAME for exception details."
    return
}

# define the type of field being used as the user identifier ["Email", "CloudId", "PrincipalName"] in the User Profile Service

    $userIdType = [Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email


# build a context to the tenant admin site

    $tenantAdminCredential  = New-Object System.Management.Automation.PSCredential($tenantUserName, $tenantSecurePwd)
    $tenantContext = New-ClientContextWithRetry -ContextUrl $tenantAdminUrl -Credential $tenantAdminCredential -CorrelationId $CorrelationId


# build a context to the json file upload site

    $uploadSiteCredential  = New-Object System.Management.Automation.PSCredential( $uploadSiteUserName, $uploadSiteSecurePwd )
    $siteContext = New-ClientContextWithRetry -ContextUrl $uploadSiteUrl -Credential $uploadSiteCredential -CorrelationId $CorrelationId


# if we are not using an exisitng JSON file, export the data from AD for the define properties

    if( -not $existingJSONFilePath -or -not (Test-Path -Path $existingJSONFilePath))
    {
        $activeDirectoryUserObjects = @()
        $activeDirectoryUserExportMappings = @{}
        $propertyFieldMappings = New-Object -Type "System.Collections.Generic.Dictionary[String,String]"


        # query active directory once for each distinct LDAP filter and store all the results, we don't want to make huge duplicate queries more than once

            try 
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Querying user data from Active Directory." -TraceLevel Low -CorrelationId $CorrelationId
                $activeDirectoryUserObjects = Get-ActiveDirectoryUsers -Filter $ldapFilter -CorrelationId $CorrelationId
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Active Directory query returned $($activeDirectoryUserObjects.Count) objects." -TraceLevel Low -CorrelationId $CorrelationId
            }
            catch 
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Active Directory query failed.  Exception: $($_.Exception)." -TraceLevel Critical -CorrelationId $CorrelationId
                Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
                return
            }

            if( -not $activeDirectoryUserObjects -or $activeDirectoryUserObjects.Count -eq 0)
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Active Directory query returned 0 records." -TraceLevel Critical -CorrelationId $CorrelationId
                Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
                return
            }

            # enumerate each of the users returned by the AD query.
            foreach( $activeDirectoryUserObject in $activeDirectoryUserObjects )
            {
                # get the value of the identifier property from the user 
                $userIdentifer = $activeDirectoryUserObject.Properties[$userIdentifierAttributeName][0]

                # filter out all the users that have a blank value for the identifier or has already been added to the hashtable
                if( $userIdentifer -and -not $activeDirectoryUserExportMappings.ContainsKey( $userIdentifer ) )
                {
                    Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Processing AD User Object '$userIdentifer'" -TraceLevel Low -CorrelationId $CorrelationId

                    # add the user to the hashtable and include the user's identifier attribute so it is dumped to the .json file
                    $activeDirectoryUserExportMappings.Add( $userIdentifer, @{ "Identifier" = $userIdentifer }) 
                
                    Invoke-MappingFunctionsForActiveDirectoryUser `
                        -IdentifierPropertyName            $userIdentifierAttributeName `
                        -ActiveDirectoryUser               $activeDirectoryUserObject `
                        -PropertyMappingObjects            $propertyMappingObjects `
                        -ActiveDirectoryUserExportMappings $activeDirectoryUserExportMappings `
                        -PropertyFieldMap                  $propertyFieldMappings `
                        -CorrelationId                     $CorrelationId
                }
                else
                {
                    if( $userIdentifer )
                    {
                        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Skipping duplicate user: $userIdentifer." -TraceLevel Verbose -CorrelationId $CorrelationId
                    }
                    else 
                    {
                        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Skipping user, identifier property value is empty." -TraceLevel Verbose -CorrelationId $CorrelationId
                    }
                }

            } # end foreach AD users

            if( $activeDirectoryUserExportMappings.Count -eq 0 )
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - LDAP query found no users found with specified property values." -TraceLevel High -CorrelationId $CorrelationId
                Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
                return
            }


      
        # transform the object into JSON

            # need this format to meet the required JSON format expected by the SPO import job
            $userProfileParentObject = New-Object PSObject -Property @{
                value = $activeDirectoryUserExportMappings.Values | Where-Object { $_.Count -gt 1 } # filter out all users that don't have any properties to map
            }

                <#
                    REQUIRED JSON FORMAT

                    {
                        "value": [
                            {
                                "IdName":    "vesaj@contoso.com",
                                "Property1": "Helsinki",
                                "Property2": "Viper"
                            },
                            {
                                "IdName":    "bjansen@contoso.com",
                                "Property1": "Brussels",
                                "Property2": "Beetle"
                            }
                        ]
                    }

                #>


            # save the results to a temp file and upload the file to SPO library

                # temp file path
                $tempFilePath = Join-Path -Path $tempFileDirectory -ChildPath "UserPropertySync-$($(Get-Date).ToString("yyyy-MM-dd_hh.mm.ss")).txt"
    
                # convert to JSON and save to temp dir
                $userProfileParentObject | ConvertTo-Json | Set-Content -Path $tempFilePath -Force
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - JSON data file: $tempFilePath." -TraceLevel Medium -CorrelationId $CorrelationId
    }
    else
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Using existing JSON file at $existingJSONFilePath" -TraceLevel Low -CorrelationId $CorrelationId
        $tempFilePath = $existingJSONFilePath

        if( -not $propertyFieldMappings -or $propertyFieldMappings.Count -lt 1 )
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - The `$propertyFieldMappings array must be initialized if you are importing an existing JSON file." -TraceLevel Critical -CorrelationId $CorrelationId
            Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
            return
        }
    }

    # upload to SPO

        $importFileFullUrl = Publish-File -ClientContext $siteContext -ListTitle $uploadSiteLibrary -FilePath $tempFilePath -CorrelationId $CorrelationId

        if($importFileFullUrl) 
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - JSON file uploaded." -TraceLevel Low -CorrelationId $CorrelationId

            # delete the temp file 
            # Remove-Item -Path $tempFilePath -Force -ErrorAction SilentlyContinue

            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - JSON file deleted" -TraceLevel Verbose -CorrelationId $CorrelationId
        }
        else
        {
            Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - JSON file upload failed, exiting script." -TraceLevel High -CorrelationId $CorrelationId
            Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
            return
        }

        $propertyFieldMappings.GetEnumerator() | ForEach-Object { Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Property Field Mappings: Key: $($_.Key), Value: $($_.Value)" -TraceLevel Verbose -CorrelationId $CorrelationId }


# queue UPA property import job

    Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Starting Import Job" -TraceLevel Medium -CorrelationId $CorrelationId

    try
    {
        $o365Tenant = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($tenantContext)
        $tenantContext.Load($o365Tenant)

        # "Identifier" is the name of the field in the JSON file that uniquely identifies the user to the import service.  This value is hardcoded in the script, so don't change it.
        $importJobId = $o365Tenant.QueueImportProfileProperties( $userIdType, "Identifier", $propertyFieldMappings, $importFileFullUrl )
        Invoke-ClientContextWithRetry -ClientContext $tenantContext -CorrelationId $CorrelationId

        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Import Job Id: $($importJobId.Value)" -TraceLevel Medium -CorrelationId $CorrelationId
    }
    catch
    {
        Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Exception starting import job. Exception: $($_.Exception)" -TraceLevel Critical -CorrelationId $CorrelationId
        Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
        return
    }

# Wait for property import job to finish. This takes at least ~20 minutes to complete

    if( $importJobId.Value -ne $null )
    {
        $importJob = $null

        while(-not $importJob -or $importJob.State -ne "Succeeded" -and $importJob.State -ne "Error" )
        {
            Start-Sleep -Seconds 60

            $importJob = $o365Tenant.GetImportProfilePropertyJob( [GUID]$importJobId.Value )
            $tenantContext.Load($importJob)
            Invoke-ClientContextWithRetry -ClientContext $tenantContext -CorrelationId $CorrelationId

            if($importJob.State -ne "Error")
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Import Job Status: $($importJob.State)" -TraceLevel Medium -CorrelationId $CorrelationId
            }
            else
            {
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Import Job Status: $($importJob.State). Status: $($importJob.Error)" -TraceLevel High    -CorrelationId $CorrelationId
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - JSON File Url: $($importJob.SourceUri)"                              -TraceLevel Verbose -CorrelationId $CorrelationId
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Import Error Message: $($importJob.ErrorMessage)"                    -TraceLevel Verbose -CorrelationId $CorrelationId
                Write-TraceLogEntry -Message "$($MyInvocation.MyCommand.Name) - Import Error Log Location: $($importJob.LogFolderUri)"               -TraceLevel Verbose -CorrelationId $CorrelationId

                Send-EmailNotification -To $notificationEmailTo -From $notificationEmailFrom -Subject "Script Execution Failure: $($MyInvocation.MyCommand.Name)" -SMTPServer $notificationEmailSMTP -CorrelationId $CorrelationId
            }
        }
    }


<#

    #Show status of all all import jobs

    $jobs = $o365Tenant.GetImportProfilePropertyJobs()
    $tenantContext.Load($jobs)
    $tenantContext.ExecuteQuery();

    foreach ($item in $jobs)
    {
       $item | SELECT JobId, State, Error, ErrorMessage
    }

#>

