<#

----------------------------------------------------------
History
----------------------------------------------------------
 08-02-2017 - Created from SearchBenchmarkWithSQLIO
 08-03-2017 - Added ability to archive the perf logs prior to data cleanup
 08-03-2017 - Fixed a bug with .dat file creation
 08-14-2017 - Removed some PSSession calls to reduce chances of remote connetion failures

==============================================================#>

# name and location of the HTML report
$reportFile = "SearchIndexDriveIOTestResults$($(Get-Date).ToString('yyyy-MM-dd')).html"


# path must exist, can be local or UNC.  If you don't want to backup the perf logs, leave this empty or null
$logArchiveDirectoryPath = "\\dc01\_backups"


# index server info.  These directories will be deleted after the script finishes
$indexServerInfo = @( 
    
    [PSCustomObject] @{
        ServerName   = "SP2013-01"
        TestPath     = "E:\SQLIO2_SP01"
        TestDuration = 10  # seconds
        DATFileSize  = 1   # GB
    },

    [PSCustomObject] @{
        ServerName   = "SP2013-02"
        TestPath     = "E:\SQLIO2_SP02"
        TestDuration = 10 # seconds 
        DATFileSize  = 1  # GB
    }

)



<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>


function Initialize-IndexServer
{
    <#
    .Synopsis
       If SQLIO.exe is not installed and does not exist in the test folder on the index servers, this function will download and install SQLIO from the internet, then copy sqlio.exe to the test folder.  
       If the .dat file needed for the SQLIO.exe test does not exist or is not the defined size, this function will create a new .dat file.

    .EXAMPLE
       Initialize-IndexServer -ServerInfo $ServerInfo
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
        #[parameter(Mandatory=$true)][object[]]$Session
    )

    try
    {
        # Invoke-Command -Session $Session -ErrorAction Stop  -ScriptBlock {
        Invoke-Command -ComputerName $ServerInfo.ServerName -ErrorAction Stop  -ScriptBlock {
            param
            (
                [parameter(Mandatory=$true)][object[]]$ServerInfo
            )

            $ServerInfo | ? { $_.ServerName -eq $env:COMPUTERNAME } | SELECT -First 1 | % {

                try
                {

                    if (-not (Test-Path -Path $_.TestPath -PathType Container ) )
                    {
                        New-Item -Path $_.TestPath -ItemType Directory | Out-Null
                    }

                    if( -not (Test-Path -Path "$($_.TestPath)\sqlio.exe") )
                    {
                        # check the default installation location
                        if (-not (Test-Path -Path "C:\Program Files (x86)\SQLIO\sqlio.exe"))
                        {
                            # download sqlio.msi
                            Invoke-WebRequest "http://download.microsoft.com/download/f/3/f/f3f92f8b-b24e-4c2e-9e86-d66df1f6f83b/SQLIO.msi" -OutFile "$($_.TestPath)\SQLIO.msi"

                            if( Test-Path -Path "$($_.TestPath)\SQLIO.msi" )
                            {
                                # install sqlio
                                Start-Process -FilePath "$($_.TestPath)\SQLIO.msi" -ArgumentList "/quiet" -Wait
                            }
                            else
                            {
                                throw "File not found: $($_.TestPath)\SQLIO.msi"
                            }
                        }

                        if (Test-Path -Path "C:\Program Files (x86)\SQLIO\sqlio.exe")
                        {
                            # copy the sqlio.exe to the test path location
                            Copy-Item -Path "C:\Program Files (x86)\SQLIO\sqlio.exe" -Destination $($_.TestPath)
                        }
                        else
                        {
                            throw "File Not Found: 'C:\Program Files (x86)\SQLIO\sqlio.exe'"
                        }
                    }   

                    $datFilePath = Join-Path -Path $_.TestPath -ChildPath "sqlio_test.dat"

                    if( -not (Test-Path -Path $datFilePath -PathType Leaf) )
                    {
                        # no existing .dat file, create one
                        Start-Process -FilePath "FSUTIL.EXE" -ArgumentList "file createnew    $datFilePath $($($_.DATFileSize) * 1GB)" -Wait
                        Start-Process -FilePath "FSUTIL.EXE" -ArgumentList "file setvaliddata $datFilePath $($($_.DATFileSize) * 1GB)" -Wait
                    }
                    elseif ( (Get-Item -Path $datFilePath).Length -ne ($_.DATFileSize * 1GB) )
                    {
                        # existing .dat file is not the required size, delete and recreate it.
                        Remove-Item -Path $datFilePath -Force
                        Start-Process -FilePath "FSUTIL.EXE" -ArgumentList "file createnew $datFilePath $($($_.DATFileSize) * 1GB)" -Wait
                    }

                    if( -not (Test-Path -Path $datFilePath -PathType Leaf) )
                    {
                        # creation failed, we need to bail out
                        throw "SQLIO DAT File Not Found: '$datFilePath'"
                    }

                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $true
                        Details  = ""
                    }


                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $true
                        Details  = ""
                    }
                }
                catch
                {
                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $false
                        Details  = $_.Exception
                    }
                }
            }

        } -ArgumentList @(,$ServerInfo)
    }
    catch
    {
        if( $_.Exception[0].Message -match "Connecting to remote server (?<ComputerName>\S*) failed" )
        {
            return New-Object PSObject @{
                Computer = $Matches["ComputerName"]
                Success  = $false
                Details  = $_.Exception.Message
            }
        }

        return New-Object PSObject @{
            Computer = "Unknown"
            Success  = $false
            Details  = $_.Exception.Message
        }
    }
}

function Start-IndexServerPerformanceTest
{
    <#
    .Synopsis
       Executes a set of SQLIO.exe performance tests against the drives on the index servers 

    .EXAMPLE
       Start-IndexServerPerformanceTest -ServerInfo $ServerInfo
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
        # [parameter(Mandatory=$true)][object[]]$Session
    )
    
    try
    {
        # Invoke-Command -Session $Session -ErrorAction Stop  -ScriptBlock {
        Invoke-Command -ComputerName $ServerInfo.ServerName -ErrorAction Stop  -ScriptBlock {
            param
            (
                [parameter(Mandatory=$true)][object[]]$ServerInfo
            )

            $ServerInfo | ? { $_.ServerName -eq $env:COMPUTERNAME } | SELECT -First 1 | % { 
        
                try
                {
                    $logPath = Join-Path -Path $_.TestPath -ChildPath "LogFiles"
                    $datPath = Join-Path -Path $_.TestPath -ChildPath "sqlio_test.dat"

                    Remove-Item -Path $logPath -Recurse -Force -ErrorAction SilentlyContinue
                    New-Item    -Path $logPath -ItemType Directory -Force | Out-Null

                    $performanceTests = [ordered]@{
                        "64k-read"    = "-s$($_.TestDuration) -kR -t4 -o25 -b64     -frandom -LS -BN $datPath"
                        "256k-write"  = "-s$($_.TestDuration) -kW -t4 -o25 -b256    -frandom -LS -BN $datPath"
                        "100mb-read"  = "-s$($_.TestDuration) -kR -t1 -o1  -b100000 -frandom -LS -BN $datPath"
                        "100mb-write" = "-s$($_.TestDuration) -kW -t1 -o1  -b100000 -frandom -LS -BN $datPath"
                    }

                    foreach( $perfTest in $performanceTests.GetEnumerator() )
                    {
                        "$($_.TestPath)\SQLIO.exe $($perfTest.Value)" | Add-Content -Path "$logPath\$($env:COMPUTERNAME)_sqlio_execution_args.txt" # dump the sqlio.exe args to a trace file 

                        Start-Process -FilePath "$($_.TestPath)\SQLIO.exe" -ArgumentList $perfTest.Value -RedirectStandardOutput "$logPath\$($env:COMPUTERNAME)_$($perfTest.Key).txt" -Wait -WorkingDirectory $_.TestPath
                        # Start-Sleep -Seconds 60 # Note: Microsoft recommends that you allow time between each sqlio command to let the I/O system return to an idle state
                    }

                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $true
                        Details  = ""
                    }
                }
                catch
                {
                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $false
                        Details  = $_.Exception
                    }
                }
            }
        } -ArgumentList @(,$ServerInfo)
    }
    catch
    {
        if( $_.Exception[0].Message -match "Connecting to remote server (?<ComputerName>\S*) failed" )
        {
            return New-Object PSObject @{
                Computer = $Matches["ComputerName"]
                Success  = $false
                Details  = $_.Exception.Message
            }
        }

        return New-Object PSObject @{
            Computer = "Unknown"
            Success  = $false
            Details  = $_.Exception.Message
        }
    }
}

function Read-PerformanceTestLogs
{
    <#
    .Synopsis
       Collections the performance data from each of index servers and returns an object containing the performance data

    .EXAMPLE
       Read-PerformanceTestLogs -ServerInfo $ServerInfo
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
        #[parameter(Mandatory=$true)][object[]]$Session
    )

    foreach( $info in $ServerInfo )
    {
        try
        {
            # unc path to logs directory
            $uncPathToLogFiles = "\\$($info.ServerName)\$($info.TestPath)\LogFiles".Replace(":", "$")

            # unc path to .dat file
            $uncPathToDATFile = "\\$($info.ServerName)\$($info.TestPath)\sqlio_test.dat".Replace(":", "$")

            $performanceResults = [PSCustomObject] @{
                ServerName   = $info.ServerName
                TestPath     = $info.TestPath
                LogPath      = $uncPathToLogFiles
                DATFileSize  = $(Get-Item -Path $uncPathToDATFile).Length / 1GB
                TestDuration = $info.TestDuration
                Read_64K     = 0
                Write_256k   = 0
                Read_100mb   = 0
                Write_100mb  = 0
                Success      = $true
                Details      = ""
            }

            foreach( $file in Get-ChildItem -Path $uncPathToLogFiles -File )
            {
                $content = Get-Content -Path $file.FullName

                switch -wildcard ($file.Name)
                {
                    "*_100mb-read.txt"
                    {
                        $performanceResults.Read_100mb  = $content | Select-String "MBs/Sec" | % { [int]$_.ToString().Replace("MBs/sec: ", "") }
                    }
                    "*_100mb-write.txt"
                    {
                        $performanceResults.Write_100mb = $content | Select-String "MBs/sec" | % { [int]$_.ToString().Replace("MBs/sec: ", "") }
                    }
                    "*_256k-write.txt"
                    {
                        $performanceResults.Write_256k  = $content | Select-String "IOs/Sec" | % { [int]$_.ToString().Replace("IOs/sec: ", "") }
                    }
                    "*_64k-read.txt"
                    {
                        $performanceResults.Read_64K    = $content | Select-String "IOs/sec" | % { [int]$_.ToString().Replace("IOs/sec: ", "") }
                    }
                } 
            }

            $performanceResults
        }
        catch
        {
            $performanceResults.Success = $false
            $performanceResults.Details = $_.Exception
        }
    }
    <#
    try
    {
        Invoke-Command -Session $Session -ErrorAction Stop  -ScriptBlock {
            param
            (
                [parameter(Mandatory=$true)][object[]]$ServerInfo
            )

            $ServerInfo | ? { $_.ServerName -eq $env:COMPUTERNAME } | SELECT -First 1 | % { 
        
                try
                {
                    $logPath = Join-Path -Path $_.TestPath -ChildPath "LogFiles"
                    $datPath = Join-Path -Path $_.TestPath -ChildPath "sqlio_test.dat"

                    $performanceResults = [PSCustomObject] @{
                        TestPath     = $_.TestPath
                        LogPath      = $logPath
                        DATFileSize  = $(Get-Item -Path $datPath).Length / 1GB
                        TestDuration = $_.TestDuration
                        Read_64K     = 0
                        Write_256k   = 0
                        Read_100mb   = 0
                        Write_100mb  = 0
                        Success      = $true
                        Details      = ""
                    }
        
                    foreach( $file in Get-ChildItem -Path $logPath -File )
                    {
                        $content = Get-Content -Path $file.FullName

                        switch -wildcard ($file.Name)
                        {
                            "*_100mb-read.txt"
                            {
                                $performanceResults.Read_100mb  = $content | Select-String "MBs/Sec" | % { [int]$_.ToString().Replace("MBs/sec: ", "") }
                            }
                            "*_100mb-write.txt"
                            {
                                $performanceResults.Write_100mb = $content | Select-String "MBs/sec" | % { [int]$_.ToString().Replace("MBs/sec: ", "") }
                            }
                            "*_256k-write.txt"
                            {
                                $performanceResults.Write_256k  = $content | Select-String "IOs/Sec" | % { [int]$_.ToString().Replace("IOs/sec: ", "") }
                            }
                            "*_64k-read.txt"
                            {
                                $performanceResults.Read_64K    = $content | Select-String "IOs/sec" | % { [int]$_.ToString().Replace("IOs/sec: ", "") }
                            }
                        } 
                    }
                }
                catch
                {
                    $performanceResults.Success = $false
                    $performanceResults.Details = $_.Exception
                }

                $performanceResults  
            }
        } -ArgumentList @(,$ServerInfo)
    }
    catch
    {
        if( $_.Exception[0].Message -match "Connecting to remote server (?<ComputerName>\S*) failed" )
        {
            return New-Object PSObject @{
                Computer = $Matches["ComputerName"]
                Success  = $false
                Details  = $_.Exception.Message
            }
        }

        return New-Object PSObject @{
            Computer = "Unknown"
            Success  = $false
            Details  = $_.Exception.Message
        }
    }
    #>
}

function Get-HardwareSpecifications
{
    <#
    .Synopsis
       Collections the hardware specs from each of index servers and returns an object containing the data

    .EXAMPLE
       Get-HardwareSpecifications -ServerInfo $ServerInfo
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
        #[parameter(Mandatory=$true)][object[]]$Session
    )
    
    foreach ($info in $ServerInfo )
    {
        try
        {
            $hardwareSpecifications = New-Object PSObject -Property @{
                ServerName        = $info.ServerName
                TotalRAM          = (Get-WMIObject -Class Win32_PhysicalMemory -ComputerName $info.ServerName | Measure-Object -Property Capacity                  -Sum).Sum / 1GB
                LogicalProcessors = (Get-WMIObject -Class Win32_Processor      -ComputerName $info.ServerName | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
                DiskInfo          = @()
                Success           = $true
                Details           = ""
            }    

            Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = 3" -ComputerName $info.ServerName | % { 

                $hardwareSpecifications.DiskInfo += New-Object PSObject -Property @{
                    DriveLetter = $_.DeviceID
                    Capacity    = [Math]::Round($_.Size / 1GB)
                    FreeSpace   = [Math]::Round($_.FreeSpace / 1Gb)
                    PercentFree = [Math]::Round(($_.FreeSpace/1GB) / ($_.Size/1GB) * 100 ,2)
                }
            }

            $hardwareSpecifications
        }
        catch
        {
            $hardwareSpecifications.Success = $false
            $hardwareSpecifications.Details = $_.Exception
        }
    }

    <#
    try
    {
        Invoke-Command -Session $Session -ErrorAction Stop  -ScriptBlock {
            param
            (
                [parameter(Mandatory=$true)][object[]]$ServerInfo
            )

            $ServerInfo | ? { $_.ServerName -eq $env:COMPUTERNAME } | SELECT -First 1 | % { 
        
                try
                {
                    $hardwareSpecifications = New-Object PSObject -Property @{
                        TotalRAM          = (Get-WMIObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity                  -Sum).Sum / 1GB
                        LogicalProcessors = (Get-WMIObject -Class Win32_Processor      | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
                        DiskInfo          = @()
                        Success           = $true
                        Details           = ""
                    }

                    Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = 3" | % { 

                        $hardwareSpecifications.DiskInfo += New-Object PSObject -Property @{
                            DriveLetter = $_.DeviceID
                            Capacity    = [Math]::Round($_.Size / 1GB)
                            FreeSpace   = [Math]::Round($_.FreeSpace / 1Gb)
                            PercentFree = [Math]::Round(($_.FreeSpace/1GB) / ($_.Size/1GB) * 100 ,2)
                        }
                    }
                }
                catch
                {
                    $hardwareSpecifications.Success = $false
                    $hardwareSpecifications.Details = $_.Exception
                }
            
                $hardwareSpecifications
            }
        } -ArgumentList @(,$ServerInfo)
    }
    catch
    {
        if( $_.Exception[0].Message -match "Connecting to remote server (?<ComputerName>\S*) failed" )
        {
            return New-Object PSObject @{
                Computer = $Matches["ComputerName"]
                Success  = $false
                Details  = $_.Exception.Message
            }
        }

        return New-Object PSObject @{
            Computer = "Unknown"
            Success  = $false
            Details  = $_.Exception.Message
        }
    }
    #>
}

function New-HTMLReport
{
    <#
    .Synopsis
       Creates a HTML file summarizing the performace results  

    .EXAMPLE
       New-HTMLReport -PerformanceTestResults $performanceTestResults -HardwareSpecifications $hardwareSpecifications
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$PerformanceTestResults,
        [parameter(Mandatory=$true)][object[]]$HardwareSpecifications
    )

    begin
    {
        function ConvertTo-ColorCodedTable
        {
            [cmdletbinding()]
            param
            (
                [parameter(Mandatory=$true)][string]$PropertyName,
                [parameter(Mandatory=$true)][int]$PropertyValue
            )

            $template = "<table class='{0}'><tr><td style='border-bottom: none'>$($PropertyValue.ToString('N0'))</td></tr></table>"


            switch($PropertyName)
            {
                "Read_64K"
                {
                    if( $PropertyValue -gt 300 )
                    {
                        return $template -f "green"
                    }

                    return $template -f "red"
                }
                "Write_256k"
                {
                    if( $PropertyValue -gt 100 )
                    {
                        return $template -f "green"
                    }

                    return $template -f "red"
                }
                "Read_100mb"
                {
                    if( $PropertyValue -gt 200 )
                    {
                        return $template -f "green"
                    }

                    return $template -f "red"
                }
                "Write_100mb"
                {
                    if( $PropertyValue -gt 200 )
                    {
                        return $template -f "green"
                    }

                    return $template -f "red"
                }
            }

        }

        function ConvertTo-TestDetailsTable
        {
            [cmdletbinding()]
            param
            (
                [parameter(Mandatory=$true)][string]$Path,
                [parameter(Mandatory=$true)][int]$Size,
                [parameter(Mandatory=$true)][int]$Seconds
            )

            New-Object PSObject -Property @{
                Path     = $path
                Size     = "$Size GB"
                Duration = "$Seconds seconds"
            } | ConvertTo-Html -Fragment -As List | Out-String
        }

        function ConvertTo-HardwareDetailsTable
        {
            [cmdletbinding()]
            param
            (
                [parameter(Mandatory=$true)][int]$LogicalProcessors,
                [parameter(Mandatory=$true)][int]$TotalRAM
            )

            New-Object PSObject -Property @{
                "Logical Processors" = $LogicalProcessors
                "Total RAM"          = "$TotalRAM GB"
            } | ConvertTo-Html -Fragment -As List | Out-String
        }

        function ConvertTo-DiskInfomationTable
        {
            [cmdletbinding()]
            param
            (
                [parameter(Mandatory=$true)][object[]]$DiskInformation
            )

            $DiskInformation | SORT DriveLetter | SELECT @{N="Drive"; E={$_.DriveLetter}}, @{N="Capacity"; E={"$($_.Capacity) GB"}}, @{N="Free Space"; E={"$($_.FreeSpace) GB"}}, @{N="Percent Free"; E={ "$($_.PercentFree)%"}} | ConvertTo-Html -Fragment -As Table| Out-String 
        }

        $headerFormat = "`t`t`t<tr><th>{0}</th><th>{1}</th><th>{2}</th><th>{3}</th><th>{4}</th><th>{5}</th><th>{6}</th><th>{7}</th><th></th></tr>`n"
        $rowFormat    = "`t`t`t<tr style='height:1px'><td>{0}</td><td style='height:inherit'>{1}</td><td style='height:inherit'>{2}</td><td style='height:inherit'>{3}</td><td style='height:inherit'>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>`n"

        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine( "<html>" )                                           | Out-Null
        $sb.AppendLine( "<Title>Search Benchmark: SharePoint 2013</title>" ) | Out-Null
        $sb.AppendLine( "<head>" )                                           | Out-Null
        $sb.AppendLine( "<style>" )                                          | Out-Null

        $sb.AppendLine( "body    { font-size: 10.5pt; font-family: Calibri, Candara, Segoe, `"Segoe UI`", Optima, Arial, sans-serif; }" )           | Out-Null
        $sb.AppendLine( "table   { width: 100%; text-align: center; padding: 5px; }" )                                                              | Out-Null
        $sb.AppendLine( "td, th  { border-bottom: solid 1px #D8D8D8; }" )                                                                           | Out-Null
        $sb.AppendLine( ".red    { background-color: #F78181; color: black; height: 100%; width: 100%; text-align: center; border-bottom: none }" ) | Out-Null
        $sb.AppendLine( ".green  { background-color: #BEF781; color: black; height: 100%; width: 100%; text-align: center; border-bottom: none }" ) | Out-Null

        $sb.AppendLine( "</style>" ) | Out-Null
        $sb.AppendLine( "</head"   ) | Out-Null
        $sb.AppendLine( "<body>"   ) | Out-Null
        $sb.AppendLine( "<h3>Reported Generated on $($env:COMPUTERNAME) at $(Get-Date)</h3>" ) | Out-Null
        $sb.AppendLine( "<table >" ) | Out-Null

    }
    process
    {
        $sb.AppendFormat( $headerFormat, "Computer", "64KB Read (IOPS)", "256KB Write (IOPS)", "100MB Read(MB/s)", "100MB Write(MB/s)", "Test Details", "Hardware", "Disk Information") | Out-Null
        $sb.AppendFormat( $rowFormat,    "Recommeded Minimum", 300, 100, 200, 200, "", "", "" ) | Out-Null

        foreach( $performanceTestResult in $PerformanceTestResults | SORT ServerName )
        {
            $hardwareSpecification = $HardwareSpecifications | ? { $_.ServerName -eq $performanceTestResult.ServerName }

            $sb.AppendFormat( $rowFormat, 
                $performanceTestResult.ServerName, 
                (ConvertTo-ColorCodedTable -PropertyName "Read_64K"    -PropertyValue $performanceTestResult.Read_64K),
                (ConvertTo-ColorCodedTable -PropertyName "Write_256k"  -PropertyValue $performanceTestResult.Write_256k),
                (ConvertTo-ColorCodedTable -PropertyName "Read_100mb"  -PropertyValue $performanceTestResult.Read_100mb),
                (ConvertTo-ColorCodedTable -PropertyName "Write_100mb" -PropertyValue $performanceTestResult.Write_100mb),
                (ConvertTo-TestDetailsTable -Path $performanceTestResult.TestPath -Size $PerformanceTestResult.DATFileSize -Seconds $performanceTestResult.TestDuration), 
                (ConvertTo-HardwareDetailsTable -LogicalProcessors $hardwareSpecification.LogicalProcessors -TotalRAM $hardwareSpecification.TotalRAM),
                (ConvertTo-DiskInfomationTable -DiskInformation $hardwareSpecification.DiskInfo)) | Out-Null
        }
    }
    end
    {
        $sb.AppendLine("`t`t</table>") | Out-Null
        $sb.AppendLine("`t<body>")     | Out-Null
        $sb.AppendLine("<html>")       | Out-Null

        $sb.ToString()
    } 
}

function Remove-PerformanceTestDirectory
{
    <#
    .Synopsis
       Deletes the test directory on each index server

    .EXAMPLE
       Remove-PerformanceTestDirectory -ServerInfo $indexServerInfo
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
        #[parameter(Mandatory=$true)][object[]]$Session
    )

    foreach ($info in $ServerInfo )
    {
        try
        {
            # unc path to logs directory
            $uncPathToLogFiles = "\\$($info.ServerName)\$($info.TestPath)".Replace(":", "$")

            if( Test-Path -Path $uncPathToLogFiles )
            {
                #Remove-Item -Path $uncPathToLogFiles -Force -Recurse                      
            }
        }
        catch
        {
            Write-Error "Error deleting $uncPathToLogFiles"
        }
    }
    <#
    try
    {
        Invoke-Command -Session $Session -ErrorAction Stop  -ScriptBlock {
            param
            (
                [parameter(Mandatory=$true)][object[]]$ServerInfo
            )

            $ServerInfo | ? { $_.ServerName -eq $env:COMPUTERNAME } | SELECT -First 1 | % { 
        
                try
                {
                    if( Test-Path -Path $_.TestPath )
                    {
                        Remove-Item -Path $_.TestPath -Force -Recurse                      
                    }
                }
                catch
                {
                    return New-Object PSObject @{
                        Computer = $env:COMPUTERNAME
                        Success  = $true
                        Details  = $_.Exception
                    }
                }

                return New-Object PSObject @{
                    Computer = $env:COMPUTERNAME
                    Success  = $true
                    Details  = $_.Exception
                }
            }
        } -ArgumentList @(,$ServerInfo)
    }
    catch
    {
        if( $_.Exception[0].Message -match "Connecting to remote server (?<ComputerName>\S*) failed" )
        {
            return New-Object PSObject @{
                Computer = $Matches["ComputerName"]
                Success  = $false
                Details  = $_.Exception.Message
            }
        }

        return New-Object PSObject @{
            Computer = "Unknown"
            Success  = $false
            Details  = $_.Exception.Message
        }
    }
    #>
}

function Compress-PerformanceTestLogs
{
    <#
    .Synopsis
       Compresses the performance logs

    .EXAMPLE
        Compress-PerformanceTestLogs -ServerInfo $indexServerInfo -ArchiveDirectoryPath "\\server\archive"
    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo,
        [parameter(Mandatory=$false)][string]$ArchiveDirectoryPath
    )

    try
    {
        Add-Type -AssemblyName "System.IO.Compression.FileSystem"

        $ServerInfo | % { 

            $computerName = $_.ServerName

            # unc path to logs directory
            $uncPathToLogFiles = "\\$($_.ServerName)\$($_.TestPath)\LogFiles".Replace(":", "$")
            
            # unc path to archive directory
            $zipFilePath = Join-Path -Path $ArchiveDirectoryPath -ChildPath "$($computerName)_LogFileArchive_$($(Get-Date).ToString('yyyy-MM-dd_hhmmss')).zip"

            if( Test-Path -Path $ArchiveDirectoryPath )
            {
                [System.IO.Compression.ZipFile]::CreateFromDirectory( $uncPathToLogFiles, $zipFilePath )
            }
            else
            {
                throw "Directory Not Found: $ArchiveDirectoryPath"
            }

            return New-Object PSObject @{
                Computer = $computerName
                Success  = $true
                Details  = ""
            }
        }
    }
    catch
    {
        return New-Object PSObject @{
            Computer = $computerName
            Success  = $false
            Details  = $_.Exception
        }
    }}

function Initialize-PSSession
{
    <#
    .Synopsis

    .EXAMPLE

    #>
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][object[]]$ServerInfo
    )

    New-PSSession -ComputerName $ServerInfo.ServerName -ErrorAction Stop
}


# init the servers

    Write-Host "$(Get-Date) - Server Initialization Starting" -ForegroundColor Green
    $initializationResult = Initialize-IndexServer -ServerInfo $indexServerInfo
    $initializationResult | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Initialization Failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; exit }
    Write-Host "$(Get-Date) - Server Initialization Complete" -ForegroundColor Green


# execute the tests

    Write-Host "$(Get-Date) - Performance Test Execution Starting" -ForegroundColor Green
    $executionResult = Start-IndexServerPerformanceTest -ServerInfo $indexServerInfo
    $executionResult | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Performance Test Execution Failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; exit }
    Write-Host "$(Get-Date) - Performance Test Execution Complete" -ForegroundColor Green


# collect the test results 

    Write-Host "$(Get-Date) - Performance Test Data Collection Starting" -ForegroundColor Green
    $performanceTestResults = Read-PerformanceTestLogs -ServerInfo $indexServerInfo
    $performanceTestResults | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Performance Test Data Collection Failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; exit }
    Write-Host "$(Get-Date) - Performance Test Data Collection Complete" -ForegroundColor Green


# collect the server hardware specs

    Write-Host "$(Get-Date) - Hardware Data Collection Starting" -ForegroundColor Green
    $hardwareSpecifications = Get-HardwareSpecifications -ServerInfo $indexServerInfo
    $hardwareSpecifications | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Hardware Data Collection Failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; }
    Write-Host "$(Get-Date) - Hardware Data Collection Complete" -ForegroundColor Green


# build an HTML report

    Write-Host "$(Get-Date) - Report Generation Starting" -ForegroundColor Green
    New-HTMLReport -PerformanceTestResults $performanceTestResults -HardwareSpecifications $hardwareSpecifications | Out-File -FilePath $reportFile
    Write-Host "$(Get-Date) - Report Generation Complete - $reportFile" -ForegroundColor Green


# archive the performance logs 

    if( $logArchiveDirectoryPath -and (Test-Path -Path $logArchiveDirectoryPath -PathType Container) )
    {
        Write-Host "$(Get-Date) - Log file compression staring." -ForegroundColor Green
        $compressionResults = Compress-PerformanceTestLogs -ServerInfo $indexServerInfo -ArchiveDirectoryPath $logArchiveDirectoryPath
        $compressionResults | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Archival of the performance logs failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; exit }
        Write-Host "$(Get-Date) - Log file compression complete." -ForegroundColor Green
    }
    elseif ($logArchiveDirectoryPath)
    {
        Write-Host "$(Get-Date) - Log archive path '$logArchiveDirectoryPath' does not exist, performance logs will be deleted." -ForegroundColor Yellow
    }


# cleanup test data

    Write-Host "$(Get-Date) - Performance Test Data Deletion Starting" -ForegroundColor Green
    $cleanupResult = Remove-PerformanceTestDirectory -ServerInfo $indexServerInfo
    $cleanupResult | ? { -not $_.Success } | % { Write-Host "$(Get-Date) - Performance Test Data Deletion Failed on server: '$($_.Computer)'`nException Details: $($_.Details)" -ForegroundColor Red; exit }
    Write-Host "$(Get-Date) - Performance Test Data Deletion Complete" -ForegroundColor Green
    


