<#
    .DESCRIPTION
        This is an Azure Hybrid Worker Runbook.
        This is a template for other tasks, utilising Connect-RemoteDomain.ps1 runbook to:
            Connect to a trusted domain with a credential for the domain. 
            It will look for a domain controller in AD Site1 first,
            if non found, it will look for one in AD Site2,
            if also non found, it will store the first DC found. 
        Tasks can now be run against the trusted domain.

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

.\Connect-RemoteDomain.ps1 `
		-Domain $Domain `
		-Site1 $Site1 `
		-Site2 $Site2 `
		-DomainCredential $DomainCredential

# Other commands goes here
