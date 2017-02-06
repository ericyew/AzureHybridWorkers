<#
    .DESCRIPTION
        This runbook will connect to a trusted remote domain with a credential asset for the domain. 
        To ensure it connects to the nearest domain,
        it will look for a domain controller in Active Directory Site1 first,
        if non found, it will look for one in Active Directory Site2,
        if also non found, it will store the first DC found. 
        This runbook should be used with other runbook to perform scheduled admin tasks on the domain.

    .NOTES
        AUTHOR: Eric Yew
        LASTEDIT: Feb 3, 2017
#>

param (
	[parameter(Mandatory=$false)] 
    [String] $Domain = "mydomain.local",
	
	[parameter(Mandatory=$false)] 
    [String] $Site1 = "ADSite1",
	
	[parameter(Mandatory=$false)] 
    [String] $Site2 = "ADSite2",

    [parameter(Mandatory=$false)] 
    [String] $DomainCredential = "CredAsset"

) 

# Find all DC in domain and add to array
    $AD = new-object 'System.DirectoryServices.ActiveDirectory.DirectoryContext'("domain", $Domain )
    $ADDC = [System.DirectoryServices.ActiveDirectory.DomainController]::FindAll($AD)

# Find DC in AzureASE site
    foreach ($element in $ADDC) {
	    if ($element.SiteName -eq $Site1 ){
            $DC = $element.Name
            break
        } elseif ($element.SiteName -eq $Site2 ){
            $DC = $element.Name
        } elseif (!$DC) {
            $DC = $element.Name
        }
    }

    Import-Module ActiveDirectory

    New-PSDrive -Name DC -PSProvider ActiveDirectory -Server $DC -Scope Global -credential (Get-AutomationPSCredential -Name $DomainCredential) -root "//RootDSE/"
    cd DC:
