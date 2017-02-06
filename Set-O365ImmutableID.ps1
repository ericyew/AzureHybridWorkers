# Jean-Pierre Simonis
# Version 1.0
# Purpose: This script is designed to run within a domain and set a custom attribute for use the Office 365 ImmutableID instead of the default ObjectGUID attribute

#########################
#      Change Log       #
#########################

# 16/06/2016 v1.0 - Initial Release
#  Features
#  - Create Unique O365 ImmutableID based on ObjectGUID attribute and store it in a custom attribute
#  - Logging Engine to screen, file, event log with custom eventlog source name
#  - Full Error trapping and handing
#  - Customisable to OU, User Filter, Source and Target attributes


#########################
#         Notes         #
#########################

# This Script Requires the following to be installed
#  - Windows Azure Active Directory Module - http://go.microsoft.com/fwlink/p/?linkid=236297
#  - Windows Active Directory RSAT Feature (need Active Directory Module)
# This Script requires the following permissions
#  - Service Account in On-Premise Active Directory
#    - Read/Write Permissions on accounts within scope of the script (or domain admins :) )
#    - Password that does not expire
#    - Local admin on server that this script runs from (to be able to run schedule task etc)

# Installation
# 1. Install All above pre-requisites
# 2. Create and configure service accounts
# 3. Login to server you wish run this script from with the On-Premise Service Account
# 3.3 Set Execution Policy to Unrestricted on server running this script
# 4. Update Configuration Section of this script
# 5. Run this script as onprem service account to test if outputting desired outcome
# 6. If all works as desired create a scheduled task for this script to run as onprem service account
# 7. Secure Script run directory permissions so powershell file cannot be modified unauthorised people

#########################
#   Pre-req Functions   #
#########################

#Function to log to file and write output to host
Function LogWrite
{
    Param (
    [Parameter(Mandatory=$True)]
    [string]$Logstring,
    [Parameter(Mandatory=$True)]
    [string]$type,
    [Parameter(Mandatory=$False)]
    [string]$v    
    )
    #Check if Logging is wanted
    if ($global:logging -eq $true) {
        
        #Determine Log Entry Type        
        Switch ($type){
            start { 
                $logType = "[Start]"
                $elogtype = "Information"
            }
            i { 
                $logType = "[Info]"
                $elogtype = "Information"
            }
            w { 
                $logType = "[Warning]"
                $elogtype = "Warning"
            }
            e { 
                $logType = "[Error]"
                $elogtype = "Error"
            }
            end { 
                $logType = "[End]"
                $elogtype = "Information"
            }
        }
        #Log Time Date for each log entry
        if ($global:logTimeDate -eq $true) {
            $TimeStamp = Get-Date -Format "yyyy-MM-dd-HH:mm"
            $TimeStamp = "[$TimeStamp]"
        } else {
            $TimeStamp = $Null
        }
        #Create Eventlog Source
        if ($global:logtoEventlog -eq $true) {
            #Check if log source exists
            $checkLogSourceExists = [System.Diagnostics.EventLog]::SourceExists("$global:eventlogSource")
            if ($checkLogSourceExists -eq $False) {
                New-EventLog -LogName Application -Source $global:eventlogSource -ErrorAction SilentlyContinue
            }
        }

        #Check if Verbose logging enabled and of log entry is marked as verbose
        if ($global:verboselog -eq $true -and $v -eq $true){
            #Check if Log Entry is marked as verbose       
                if ($global:logtoFile -eq $true) {Add-Content $Logfile -value "$TimeStamp[Verbose]$logType $logstring" -ErrorAction Stop}
                if ($global:logtoEventlog -eq $true) {Write-EventLog –LogName Application –Source global:eventlogSource –EntryType $elogtype –EventID 1 –Message $logstring -ErrorAction Stop}
                Write-Host "[Verbose]$logType $logstring"

        } else {
            #Check if log is verbose if it is dont log it otherwise log a standard entry
            if ($v -eq $true) {        
                #Do Nothing
            } else {
                #Write Standard Log Entry
                if ($global:logtoFile -eq $true) {Add-Content $Logfile -value "$TimeStamp$logType $logstring" -ErrorAction Stop}
                if ($global:logtoEventlog -eq $true) {Write-EventLog –LogName Application –Source $global:eventlogSource –EntryType $elogtype –EventID 1 –Message $logstring -ErrorAction Stop}
            }

            #Write Standard Log Entries to Screen if verbose logging is enabled
            if ($global:verboselog -eq $true){
                Write-Host "$logType $logstring"
            }
        }

    }
}

#########################
#     Configuration     #
#########################
 
# General
    $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $TimeDate = Get-Date -Format "yyyy-MM-dd-HH:mm"

# Logging
    $Logfile = ($PSScriptRoot + "\O365UniqueIDAssignment.log")
    $global:logging = $true
    $global:verboselog = $false
    $global:logtoEventlog = $true
    $global:logtoFile = $false
    $global:logTimeDate = $true
    $global:eventlogSource = "O365 ImmutableID Assignment"

    #Start logging and check logfile access
	try 
	{
		Write-Host "Office 365 Immutable ID Assignment Script started" -ForegroundColor Green
		LogWrite "Office 365 Immutable ID Assignment Script started - $TimeDate" -type start
	} 
	catch 
	{
        Throw "You don't have write permissions to $logfile, please start an elevated PowerShell prompt or change NTFS permissions"
	}
  
# On-Premise Active Directory Provisioning Configuration
    #Base DN to start user search from
    $global:ADDSSearchBase = "OU=SlaterGordon,DC=slatergordon,DC=com,DC=au"
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
            Write-Host "Imported module $name" -ForegroundColor Yellow
            LogWrite "Imported module $name" -type i -v $true
        } 
        else 
        { 
            LogWrite "Module $name is not installed. Exiting..." -type e
            Throw "Module $name is not installed. Exiting..." 
        }
    } 
    else 
    { 
        LogWrite "Module $name is already loaded..." -type w -v $true
        Write-Host "Module $name is already loaded..." -ForegroundColor Yellow
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
		Write-Host "Collecting Office 365 User List" -ForegroundColor Green
		LogWrite "Collecting Office 365 User List" -type i
		#Determine Result Size
		if ($DSResultSize -eq "All") {
			#Clear ResultSetSize Variable
			$DSResultSize = $null
			#Collect AD Users All Results
			$CollectUsers = Get-ADUser -filter $DSFilter -Properties $DSTargetProperty -SearchBase $DSSrchBase -SearchScope $DSSrchScope | Select-Object Name,UserPrincipalName,SamAccountName,$DSTargetProperty   
			#Write to Log
			LogWrite "Collecting Users from $DSSrchBase filtered by $DSFilter" -type i
			Write-Host "Collecting Users from $DSSrchBase filtered by $DSFilter" -ForegroundColor Cyan
			Write-Host "Collecting Office 365 User List complete" -ForegroundColor Green
			LogWrite "Collecting Office 365 User List complete" -type i
			#Return Values on completion
			Return $CollectUsers
			
			} else {
			#Do Nothing and use configured result size
			#Collect AD Users All Results
			$CollectUsers = Get-ADUser -filter $DSFilter -Properties $DSTargetProperty -SearchBase $DSSrchBase -SearchScope $DSSrchScope -ResultSetSize $DSResultSize | Select-Object Name,UserPrincipalName,SamAccountName,$DSTargetProperty   
			#Write to Log
			LogWrite "Collecting Users from $DSSrchBase filtered by $DSFilter" -type i
			Write-Host "Collecting Users from $DSSrchBase filtered by $DSFilter" -ForegroundColor Cyan
			Write-Host "Collecting Office 365 User List complete" -ForegroundColor Green
			LogWrite "Collecting Office 365 User List complete" -type i
			#Return Values on completion
			Return $CollectUsers
		}
    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		LogWrite "Could not Collect Users from $DSSrchBase filtered by $DSFilter, `r`nError: $err" -type e
		Write-Host "Could not Collect Users from $DSSrchBase filtered by $DSFilter, `r`nError: $err" -ForegroundColor Red
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
		Write-Host "Determining users requiring updates" -ForegroundColor Green
		LogWrite "Determining users requiring updates" -type i
		#Create User to Update Array Variable
		$UsersToUpdate = @()
		#Begin checking supplied users to determine if TargetProperty has already been set
		ForEach ($user in $DSUserList)
		{
		#Routine to check if TargetProperty is not set for the user 
		$checkCloudAttrib = $user.$DSTargetProperty
        #Friendly user name variable
		$DSuser = $user.'Name'
		Write-Host "Checking $DSuser if $DSTargetProperty is set" -ForegroundColor White
		LogWrite "Checking $DSuser if $DSTargetProperty is set" -type i -v $true
		if ($checkCloudAttrib -eq $null) {
			Write-Host " - The user $DSuser does not have $DSTargetProperty set." -ForegroundColor Yellow
			LogWrite "The user $DSuser does not have $DSTargetProperty set." -type i
			#Add User to the UsersToUpdate Array Variable
			$UsersToUpdate = $UsersToUpdate + $user
			
			}
		else {
			Write-Host " - The user $DSuser already has $DSTargetProperty set." -ForegroundColor Magenta
			LogWrite "The user $DSuser already has $DSTargetProperty set." -type i -v $true
				#Routine to determine is possibly set with an value other than expected
				if ($checkCloudAttrib.length -ne 24) {
					Write-Host "   - The user $DSuser may have an unexpected value in $DSTargetProperty set." -ForegroundColor Red
					LogWrite "The user $DSuser may have an unexpected value in $DSTargetProperty set." -type e
			 	}		    
			}
		}
		
		#Write to Log
		Write-Host "Determining users requiring updates complete" -ForegroundColor Green
		LogWrite "Determining users requiring updates complete" -type i
	
		#Return Values on completion
		Return $UsersToUpdate

    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		LogWrite "Could not determine the users required to be updated, `r`nError: $err" -type e
		Write-Host "Could not determine the users required to be updated, `r`nError: $err" -ForegroundColor Red
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
		Write-Host "Applying O365 Immutable ID to required users..." -ForegroundColor Green
		LogWrite "Applying O365 Immutable ID to required users..." -type i

		#Routine to apply O365 ImmutableID attribute
		ForEach ($user in $DSUserList)
		{
		#Friendly user name variable
		$DSuser = $user.'Name'
		#AD User account name variable
		$aduser = $user.'SamAccountName'
		Write-Host "Setting $DSuser $DSTargetProperty to the same value as ObjectGUID" -ForegroundColor White

		#convert the objectguid to string and set this value into the TargetProperty attribute as the immutable id
		$objectguidstring = Get-ADUser -Identity $aduser -Properties ObjectGUID | select ObjectGUID | foreach {[system.convert]::ToBase64String(([GUID]($_.ObjectGUID)).tobytearray())}
		LogWrite "Clearing $DSuser $DSTargetProperty" -type i -v $true
		Set-ADUser -Identity $aduser -Clear $DSTargetProperty
		LogWrite "Setting $DSuser $DSTargetProperty to the same value as ObjectGUID" -type i
		Set-ADUser -Identity $aduser -Replace @{$DSTargetProperty=$objectguidstring}

		}
		
		#Write to Log
		Write-Host "Applying O365 Immutable ID to required users complete" -ForegroundColor Green
		LogWrite "Applying O365 Immutable ID to required users complete" -type i
	
    } 
    catch [system.exception]
    {
		$err = $_.Exception.Message
		LogWrite "Could not apply O365 Immutable ID to required users, `r`nError: $err" -type e
		Write-Host "Could not apply O365 Immutable ID to required users, `r`nError: $err" -ForegroundColor Red
    }
}

#########################
#        Modules        #
#########################

LogWrite "Loading Required PowerShell Modules" -type i
Write-Host "Loading Required PowerShell Modules" -ForegroundColor Green

#Active Directory
Import-MyModule ActiveDirectory

#########################
#       Execution       #
#########################

##Collect Office 365 User List
$UserList = Collect-O365Users $global:ADDSUserFilter $global:ADDSUserO365ImmutableTargetAttribute $global:ADDSSearchBase $global:ADDSSearchScope $global:ADDSResultSetSize
##Determine users that require O365 ImmutableID

#Check if userlist is not empty otherwise exit script
If ($UserList -ne $null){ 
	$UsersToUpdate = Determine-UsersToUpdate $UserList $global:ADDSUserO365ImmutableTargetAttribute
} else {
	Write-Host "No O365 Users were found. Exiting.." -ForegroundColor Yellow
	LogWrite "No O365 Users were found" -type w
	LogWrite "Office 365 ImmutableID Assignment Script Completed - $TimeDate" -type end
	#Quit Script
	Exit
}

##Apply New Office 365 ImmutableID to required users
#Check if Users requiring update list is not empty otherwise exit script
If ($UsersToUpdate -ne $null){ 
	$UpdateUsers = Apply-O365ImmutableID $UsersToUpdate $global:ADDSUserO365ImmutableTargetAttribute
} else {
	Write-Host "No O365 Users required updating" -ForegroundColor White
	LogWrite "No O365 Users required updating" -type i
	Write-Host "Office 365 Immutable ID Assignment Script Completed" -ForegroundColor Green
	LogWrite "Office 365 Immutable ID Assignment Script Completed - $TimeDate" -type end
	#Quit Script
	Exit
}
#End of Script
Write-Host "Office 365 Immutable ID Assignment Script Completed" -ForegroundColor Green
LogWrite "Office 365 Immutable ID Assignment Script Completed - $TimeDate" -type end
