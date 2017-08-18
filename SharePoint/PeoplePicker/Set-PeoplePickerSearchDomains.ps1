<#

----------------------------------------------------------
 History
 ----------------------------------------------------------
 10-01-2014 - Created
 03-20-2015 - Added ability to set the "disable-netbios-dc-resolve" farm property 

==============================================================#>

$webApplicationUrl = "https://hhroot.contoso.com"

# https://support.microsoft.com/en-us/help/2874332/users-in-any-domain-in-the-trusted-forest-receive-a-no-exact-match-was
$disableNetbiosDCResolution = $false;

$searchDomains = @() # array to store custom domains, do not modify

# NORTHAMERICA
$searchDomains += New-Object PSObject -Property @{
    DomainName      = "northamerica.contoso.com"
    ShortDomainName = "";
    CustomFilter    = "";
    IsForest        = $false
    LoginName       = ""
}

# EUROPE
$searchDomains += New-Object PSObject -Property @{
    DomainName      = "europe.contoso.com"
    ShortDomainName = "";
    CustomFilter    = "";
    IsForest        = $false
    LoginName       = "contoso\user01"
}



<############    YOU SHOULD NOT HAVE TO MODIFY ANYTHING BELOW THIS POINT    ############>
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

if($disableNetbiosDCResolution)
{
    $farm = Get-SPFarm
    $farm.Properties["disable-netbios-dc-resolve"] = $true
    $farm.Update()
}

$webApplication = Get-SPWebApplication $webApplicationUrl 

if($webApplication)
{
    # clear out all current values   
    $activeDirectoryDomains = $webApplication.PeoplePickerSettings.SearchActiveDirectoryDomains
    $origActiveDirectoryDomains = $activeDirectoryDomains
    $activeDirectoryDomains.Clear()
            
    foreach($domain in $searchDomains)
    {
        $peoplePickerSearchActiveDirectoryDomain = New-Object Microsoft.SharePoint.Administration.SPPeoplePickerSearchActiveDirectoryDomain
        $peoplePickerSearchActiveDirectoryDomain.DomainName      = $domain.DomainName
        $peoplePickerSearchActiveDirectoryDomain.IsForest        = $domain.IsForest
        $peoplePickerSearchActiveDirectoryDomain.ShortDomainName = $domain.ShortDomainName
        $peoplePickerSearchActiveDirectoryDomain.CustomFilter    = $domain.CustomFilter
                
        if( -not $domain.LoginName )
        {
            # make sure the AppCredentialKey is set, we need this to encrypt the domain password
            if(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Shared Tools\Web Server Extensions\$((Get-SPFarm).BuildVersion.Major).0\Secure" -Name "AppCredentialKey" -ErrorAction SilentlyContinue)
            {
                $domainCred = Get-Credential -UserName $domain.LoginName -Message "Please enter the credentials for the account to query the domain $($domain.DomainName)"
                $peoplePickerSearchActiveDirectoryDomain.LoginName = $domainCred.UserName
                $peoplePickerSearchActiveDirectoryDomain.SetPassword($domainCred.Password)
            }
            else
            {
                Write-Host "`nWhen a password is provided for the first time, you must first set the Application Credential key on every WFE in the farm."
                Write-Host "`nSTSADM.exe Syntax:" -ForegroundColor Green

                Write-Host "`n`tstsadm –o setapppassword -password <password>"

                Write-Host "`nPowerShell Syntax:" -ForegroundColor Green
                Write-Host "`t[Microsoft.SharePoint.SPSecurity]::SetApplicationCredentialKey((ConvertTo-SecureString `"<password>`" -AsPlainText -Force))`n"

                return
            }
        }
                
        $activeDirectoryDomains.Add($peoplePickerSearchActiveDirectoryDomain)
    }
            
    $webApplication.Update()
}
