

[cmdletbinding()]
param(
	[string]$Name
)

[string[]]$ComputerName = Get-ADComputer -Filter {operatingsystem -notlike "*server*"} | Select -Expand Name

$Output= $PSScriptRoot + "\results.csv"
foreach($Computer in $ComputerName) {
	Write-Verbose "Working on $Computer"
	if(!(Test-Connection -ComputerName $Computer -Count 1 -quiet)) {
		Write-Verbose "$Computer is not online"
		Continue
	}
	
	try {
            $results = Get-WmiObject Win32_Computersystem -ComputerName $Computer
            $results | Select-Object PrimaryOwnerName,Name,Model | Export-CSV -Path $Output -NoTypeInformation
	} catch {
		Write-Verbose "Error occurred while querying $Computer. $_"
		Continue
	}

}