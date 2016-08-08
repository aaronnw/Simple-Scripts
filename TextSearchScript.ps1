[cmdletbinding()]
param(
	[string]$Directory,
    [string[]]$Filetype,
    [string[]]$Search,
    [string]$Output
)
if($Search -eq $null){
    Write-Host "No search key specified. Exiting."
    break
}
if(!$Directory){
    Write-Host "No directory specified. Using script directory."
    $Directory= $PSScriptRoot
}
if(!$Output){
    Write-Host "No output parameter. CSV output set to script directory."
    $Output= $PSScriptRoot
}
if($Filetype -eq $null){
    $Filetype = "txt"
    Write-Host "No filetype specified. Using txt by default."
}
$results = New-Object System.Collections.Generic.List[System.Object]
foreach($Type in $Filetype){
    foreach($Key in $Search){
        Get-ChildItem $Directory -Include ('*.' + $Type) -Recurse | Select-String -Pattern $Key | %{$results.Add($_)}
    }
}

$Output = $Output + "\results.csv"
$results | Select-Object @{expression={$_.LineNumber}; label = 'Number'},Filename,Path,@{expression={$_.Pattern}; label = 'Search'} | Export-CSV -Path $Output -NoTypeInformation
Write-Host $Output " written. "
