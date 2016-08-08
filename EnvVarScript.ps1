<#
	.Synopsis 
		Gets the environment variable details of a server (local or remote)
		
	.Description
		Gets the environment variable details of a server (local or remote)
		
	.Parameter ComputerName
		Name of the computer from which you want to query environment variables
		
	.Parameter Name
		Name of the environment variable which you want to query
		
	.Example
		Get-EnvironmentVariable.ps1 -ComputerName TESTPC1

		Returns all environment variables from TESTPC1 computer.
		
	.Example
		Get-EnvironmentVariable.ps1 -ComputerName TESTPC1 -Name TEMP
		Returns all environment variables from TESTPC1 computer that has the name TEMP.
		
	.Notes
		NAME:      Get-EnvironmentVariable.ps1
		AUTHOR:    Sitaram Pamarthi
		Website:   www.techibee.com
		
#>

[cmdletbinding()]
param(
	[string]$Name
)

[string[]]$ComputerName = Get-ADComputer -Filter { OperatingSystem -Like '*Server*'} -Properties OperatingSystem | Select -Expand Name

foreach($Computer in $ComputerName) {
	Write-Verbose "Working on $Computer"
	if(!(Test-Connection -ComputerName $Computer -Count 1 -quiet)) {
		Write-Verbose "$Computer is not online"
		Continue
	}
	
	try {
		$EnvObj = @(Get-WMIObject -Class Win32_Environment -ComputerName $Computer -EA Stop)
		if(!$EnvObj) {
			Write-Verbose "$Computer returned empty list of environment variables"
			Continue
		}
		Write-Verbose "Successfully queried $Computer"
		
		if($Name) {
			Write-Verbose "Looking for environment variable with the name $name"
			$Env = $EnvObj | Where-Object {$_.Name -eq $Name}
			if(!$Env) {
				Write-Verbose "$Computer has no environment variable with name $Name"
				Continue
			}
            $Computer
            $IP = Get-WmiObject win32_networkadapterconfiguration -ComputerName $Computer | where { $_.ipaddress -like "1*" } | select -ExpandProperty ipaddress | select -First 1
            $IP
            $Value = $Env.VariableValue       
            $Value
            $Computer | Out-File -Append output.txt
            $IP | Out-File -Append output.txt
            $Value | Out-File -Append output.txt
		} else {
			Write-Verbose "No environment variable specified. Listing all"
			$EnvObj
		}
		
	} catch {
		Write-Verbose "Error occurred while querying $Computer. $_"
		Continue
	}

}