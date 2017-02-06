<# 
	.SYNOPSIS
		Set a custom attribute to an AD Object utilising the Office 365 ImmutableID
	
	.DESCRIPTION
		This runbook is designed to run within a domain and set a custom attribute utilising the 
		Office 365 ImmutableID instead of the default ObjectGUID attribute.
	
	.PARAMETER
	
	.REQUIREMENTS
		This runbook requires the following to be installed on the hybrid worker server
			- Windows Azure Active Directory Module - http://go.microsoft.com/fwlink/p/?linkid=236297
			- Windows Active Directory RSAT Feature (need Active Directory Module)
		This runbook requires the following permissions
			- Credential asset that acts as a Service Account in On-Premise Active Directory
				- Read/Write Permissions on accounts within scope of the runbook (or domain admins :) )
				- Password that does not expire
		Installation
		1. Update Configuration Section of this script
		2. Run this script as onprem service account to test if outputting desired outcome
		3. If all works as desired create a scheduled task for this script to run as onprem service account
		4. Secure Script run directory permissions so powershell file cannot be modified unauthorised people

	.NOTES
		Author: Jean-Pierre Simonis
		Modified By: Eric Yew
		LASTEDIT: Feb 6, 2017

	.CHANGE LOG
		16/06/2016 v1.0 - Initial Release
			Features
			- Create Unique O365 ImmutableID based on ObjectGUID attribute and store it in a custom attribute
			- Logging Engine to screen, file, event log with custom eventlog source name
			- Full Error trapping and handing
			- Customisable to OU, User Filter, Source and Target attributes
		6/02/2017 - Modified to work with Azure Hybrid Worker
#>

#########################
#     Configuration     #
#########################

# Connect to Remote Domain
    param (
        [parameter(Mandatory=$false)] 
        [String] $Domain = "slatergordon.com.au",
        
        [parameter(Mandatory=$false)] 
        [String] $Site1 = "AzureASE",
        
        [parameter(Mandatory=$false)] 
        [String] $Site2 = "MelbDataCentre",

        [parameter(Mandatory=$false)] 
        [String] $DomainCredential = "SGAU serviceadmin",

        [parameter(Mandatory=$false)] 
    	[String] $ADSearchBase = "OU=SlaterGordon,DC=slatergordon,DC=com,DC=au"
    ) 

    .\Connect-RemoteDomain.ps1 `
            -Domain $Domain `
            -Site1 $Site1 `
            -Site2 $Site2 `
            -DomainCredential $DomainCredential 

# On-Premise Active Directory Provisioning Configuration
    	#Base DN to start user search from
    		$global:ADDSSearchBase = $ADSearchBase
	    #Search Scope 0 = Base, 1 = One Level, 2 = SubTree
    		$global:ADDSSearchScope = "2"
	    #Search ResultSet Size some number eg 100 or All = All Results
    		$global:ADDSResultSetSize = "All"
	    #User Filter
		    $global:ADDSUserFilter = "*"
	    #Attribute to Store Office 365 ImmutableID
		    $global:ADDSUserO365ImmutableTargetAttribute = "msDS-cloudExtensionAttribute1"

#########################
#       Functions       #
#########################

#Function to Import Modules with error handling
Function Import-MyModule 
{ 
Param([string]$name) 
    if(-not(Get-Module -name $name)) 
    { 
        if(Get-Module -ListAvailable $name) 
        { 
            Import-Module -Name $name 
            Write-Output "Imported module $name"
            #LogWrite "Imported module $name" -type i -v $true
        } 
        else 
        { 
            Write-Error "Module $name is not installed. Exiting..."
            Throw "Module $name is not installed. Exiting..." 
        }
    } 
    else 
    { 
        #LogWrite "Module $name is already loaded..." -type w -v $true
        Write-Output "Module $name is already loaded..."
    } 
}  

Function Collect-O365Users
{
param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('ADSFilter')]
    [string]
    $DSFilter,
    [Parameter(
        Position=1, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('TargetProperty')]
    [string]
    $DSTargetProperty,
    [Parameter(
        Position=2, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('SearchBase')]
    [string]
	$DSSrchBase,
	[Parameter(
        Position=3, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('SearchScope')]
    [string]
	$DSSrchScope,
	[Parameter(
        Position=4, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('ResultSetSize')]
    [string]
	$DSResultSize
)
   try 
    {	 
		#Write to Log
		Write-Output "Collecting Office 365 User List"
		#LogWrite "Collecting Office 365 User List" -type i
		#Determine Result Size
		if ($DSResultSize -eq "All") {
			#Clear ResultSetSize Variable
			$DSResultSize = $null
			#Collect AD Users All Results
			$CollectUsers = Get-ADUser -filter $DSFilter -Properties $DSTargetProperty -SearchBase $DSSrchBase -SearchScope $DSSrchScope | Select-Object Name,UserPrincipalName,SamAccountName,$DSTargetProperty   
			#Write to Log
			#LogWrite "Collecting Users from $DSSrchBase filtered by $DSFilter" -type i
			Write-Output "Collecting Users from $DSSrchBase filtered by $DSFilter"
			Write-Output "Collecting Office 365 User List complete"
			#LogWrite "Collecting Office 365 User List complete" -type i
			#Return Values on completion
			Return $CollectUsers
			
			} else {
			#Do Nothing and use configured result size
			#Collect AD Users All Results
			$CollectUsers = Get-ADUser -filter $DSFilter -Properties $DSTargetProperty -SearchBase $DSSrchBase -SearchScope $DSSrchScope -ResultSetSize $DSResultSize | Select-Object Name,UserPrincipalName,SamAccountName,$DSTargetProperty   
			#Write to Log
			#LogWrite "Collecting Users from $DSSrchBase filtered by $DSFilter" -type i
			Write-Output "Collecting Users from $DSSrchBase filtered by $DSFilter"
			Write-Output "Collecting Office 365 User List complete"
			#LogWrite "Collecting Office 365 User List complete" -type i
			#Return Values on completion
			Return $CollectUsers
		}
    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		#LogWrite "Could not Collect Users from $DSSrchBase filtered by $DSFilter, `r`nError: $err" -type e
		Write-Error "Could not Collect Users from $DSSrchBase filtered by $DSFilter, `r`nError: $err"
    }
}

Function Determine-UsersToUpdate
{
param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('UserList')]
    [Array]
    $DSUserList,
    [Parameter(
        Position=1, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('TargetProperty')]
    [string]
    $DSTargetProperty    
)
   try 
    {	 
		#Write to Log
		Write-Output "Determining users requiring updates"
		#LogWrite "Determining users requiring updates" -type i
		#Create User to Update Array Variable
		$UsersToUpdate = @()
		#Begin checking supplied users to determine if TargetProperty has already been set
		ForEach ($user in $DSUserList)
		{
		#Routine to check if TargetProperty is not set for the user 
		$checkCloudAttrib = $user.$DSTargetProperty
        #Friendly user name variable
		$DSuser = $user.'Name'
		Write-Output "Checking $DSuser if $DSTargetProperty is set"
		#LogWrite "Checking $DSuser if $DSTargetProperty is set" -type i -v $true
		if ($checkCloudAttrib -eq $null) {
			Write-Output " - The user $DSuser does not have $DSTargetProperty set."
			#LogWrite "The user $DSuser does not have $DSTargetProperty set." -type i
			#Add User to the UsersToUpdate Array Variable
			$UsersToUpdate = $UsersToUpdate + $user
			
			}
		else {
			Write-Output " - The user $DSuser already has $DSTargetProperty set."
			#LogWrite "The user $DSuser already has $DSTargetProperty set." -type i -v $true
				#Routine to determine is possibly set with an value other than expected
				if ($checkCloudAttrib.length -ne 24) {
					Write-Error "   - The user $DSuser may have an unexpected value in $DSTargetProperty set."
					#LogWrite "The user $DSuser may have an unexpected value in $DSTargetProperty set." -type e
			 	}		    
			}
		}
		
		#Write to Log
		Write-Output "Determining users requiring updates complete"
		#LogWrite "Determining users requiring updates complete" -type i
	
		#Return Values on completion
		Return $UsersToUpdate

    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		#LogWrite "Could not determine the users required to be updated, `r`nError: $err" -type e
		Write-Error "Could not determine the users required to be updated, `r`nError: $err"
    }
}

Function Apply-O365ImmutableID
{
param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('UserList')]
    [Array]
    $DSUserList,
    [Parameter(
        Position=1, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('TargetProperty')]
    [string]
    $DSTargetProperty    
)
   try 
    {	 
		#Write to Log
		Write-Output "Applying O365 Immutable ID to required users..."
		#LogWrite "Applying O365 Immutable ID to required users..." -type i

		#Routine to apply O365 ImmutableID attribute
		ForEach ($user in $DSUserList)
		{
		#Friendly user name variable
		$DSuser = $user.'Name'
		#AD User account name variable
		$aduser = $user.'SamAccountName'
		Write-Output "Setting $DSuser $DSTargetProperty to the same value as ObjectGUID"

		#convert the objectguid to string and set this value into the TargetProperty attribute as the immutable id
		$objectguidstring = Get-ADUser -Identity $aduser -Properties ObjectGUID | select ObjectGUID | foreach {[system.convert]::ToBase64String(([GUID]($_.ObjectGUID)).tobytearray())}
		#LogWrite "Clearing $DSuser $DSTargetProperty" -type i -v $true
		Set-ADUser -Identity $aduser -Clear $DSTargetProperty
		#LogWrite "Setting $DSuser $DSTargetProperty to the same value as ObjectGUID" -type i
		Set-ADUser -Identity $aduser -Replace @{$DSTargetProperty=$objectguidstring}

		}
		
		#Write to Log
		Write-Output "Applying O365 Immutable ID to required users complete"
		#LogWrite "Applying O365 Immutable ID to required users complete" -type i
	
    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		#LogWrite "Could not apply O365 Immutable ID to required users, `r`nError: $err" -type e
		Write-Error "Could not apply O365 Immutable ID to required users, `r`nError: $err"
    }
}

#########################
#        Modules        #
#########################

#LogWrite "Loading Required PowerShell Modules" -type i
Write-Output "Loading Required PowerShell Modules"

#Active Directory
Import-MyModule ActiveDirectory

#########################
#       Execution       #
#########################

##Testing
#Write-Output $global:ADDSUserFilter $global:ADDSUserO365ImmutableTargetAttribute $global:ADDSSearchBase $global:ADDSSearchScope $global:ADDSResultSetSize

##Collect Office 365 User List
$UserList = Collect-O365Users $global:ADDSUserFilter $global:ADDSUserO365ImmutableTargetAttribute $global:ADDSSearchBase $global:ADDSSearchScope $global:ADDSResultSetSize
##Determine users that require O365 ImmutableID

##Testing
#Write-Output $UserList

#Check if userlist is not empty otherwise exit script
If ($UserList -ne $null){ 
	$UsersToUpdate = Determine-UsersToUpdate $UserList $global:ADDSUserO365ImmutableTargetAttribute
} else {
	Write-Output "No O365 Users were found. Exiting.."
	#LogWrite "No O365 Users were found" -type w
	#LogWrite "Office 365 ImmutableID Assignment Script Completed - $TimeDate" -type end
	#Quit Script
	Exit
}

##Apply New Office 365 ImmutableID to required users
#Check if Users requiring update list is not empty otherwise exit script
If ($UsersToUpdate -ne $null){ 
	$UpdateUsers = Apply-O365ImmutableID $UsersToUpdate $global:ADDSUserO365ImmutableTargetAttribute
} else {
	Write-Output "No O365 Users required updating"
	#LogWrite "No O365 Users required updating" -type i
	Write-Output "Office 365 Immutable ID Assignment Script Completed"
	#LogWrite "Office 365 Immutable ID Assignment Script Completed - $TimeDate" -type end
	#Quit Script
	Exit
}


#End of Script
Write-Output "Office 365 Immutable ID Assignment Script Completed"
#LogWrite "Office 365 Immutable ID Assignment Script Completed - $TimeDate" -type end


