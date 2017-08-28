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

HISTORY

    08-28-2017 - Created

==============================================================#>

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

function Get-CustomResultSource
{
    <#
    .SYNOPSIS

       This cmdlet will return Site Url, Web Url, Name, Creation Date, Status and IsDefault value for each custom search resource source for
       the specificed site collection.

    .EXAMPLE
        
        Get-CustomResultSource -Site $site -SearchServiceApplication $ssa

    .FUNCTIONALITY

        PowerShell Language
    #>
    [cmdletbinding()]
    [OutputType([object[]])]
    param
    (
        # SPSite to search for custom result sources
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.SPSite]$Site,

        # SearchServiceApplication associated with the site collection
        [Parameter(Mandatory=$true)][Microsoft.Office.Server.Search.Administration.SearchServiceApplication]$SearchServiceApplication
    )

    begin
    {
        $siteLevel = [Microsoft.Office.Server.Search.Administration.SearchObjectLevel]::SPSite
        $webLevel  = [Microsoft.Office.Server.Search.Administration.SearchObjectLevel]::SPWeb
        $defaultSource = $null
    }
    process
    {
        # we can't read from readlocked sites, so skip them
        if (-not $site.IsReadLocked )
        {
            $owner  = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectOwner( $siteLevel, $site.RootWeb )
            $filter = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectFilter( $owner )
            $filter.IncludeHigherLevel = $false

            $federationManager = New-Object Microsoft.Office.Server.Search.Administration.Query.FederationManager($SearchServiceApplication)
            $siteSources = $federationManager.ListSourcesWithDefault( $filter, $false, [ref]$defaultSource )

            # filter out all built in and non-site level result sources
            $siteSources | ? { -not $_.BuiltIn -and $_.Owner.Level -eq $siteLevel } | SELECT @{ Name="SiteUrl"; Expression={ $site.Url}}, @{ Name="WebUrl"; Expression={ $site.Url}}, Name, CreatedDate, @{ Name="Status"; Expression={ if ($_.Active) { return "Active"}else{ return "Inactive"} }}, @{ Name="IsDefault"; Expression={ $_.Id -eq $defaultSource.Id}}


            foreach ($web in $site.AllWebs | ? { -not $_.IsAppWeb })
            {
                $owner  = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectOwner( $webLevel, $web )
                $filter = New-Object Microsoft.Office.Server.Search.Administration.SearchObjectFilter( $owner )
                $filter.IncludeHigherLevel = $false

                $federationManager = New-Object Microsoft.Office.Server.Search.Administration.Query.FederationManager($SearchServiceApplication)
                $webSources = $federationManager.ListSourcesWithDefault( $filter, $false, [ref]$defaultSource )

                # filter out all built in and non-web level result sources
                $webSources | ? { -not $_.BuiltIn -and $_.Owner.Level -eq $webLevel } | SELECT @{ Name="SiteUrl"; Expression={ $site.Url}}, @{ Name="WebUrl"; Expression={ $web.Url}}, Name, CreatedDate, @{ Name="Status"; Expression={ if ($_.Active) { return "Active"}else{ return "Inactive"} }}, @{ Name="IsDefault"; Expression={ $_.Id -eq $defaultSource.Id}}
            }
        }
    }
    end
    {
    }    
}

# array to store results
$customResultSources = @()

# get the search service
$ssa = Get-SPEnterpriseSearchServiceApplication | SELECT -First 1

# get the custom result sources for all sites in the farm
Get-SPSite -Limit All | % { $customResultSources += Get-CustomResultSource -Site $_ -SearchServiceApplication $ssa }

# save the results to the ULS logs directory in CSV format
$customResultSources  | Export-Csv -Path "$($(Get-SPDiagnosticConfig).LogLocation)\CustomResultSources$($(Get-Date).ToString('yyyyMMdd')).csv" -NoTypeInformation


