param(
	[string]$Path,
    [string]$Output
)
if(!$Path){
    Write-Host "No file specified. Use -Path [File Name]"
    exit
}
if(!$Output){
    Write-Host "No output parameter. Output set to script directory."
    $Output= $PSScriptRoot
}
$Output = $Output + "\results.txt"
if(Test-Path $Output){
    Remove-Item $Output
}
[string[]]$rows = Select-String -Path $Path -Pattern '\d+\d+\d+-\d+\d+-\d+\d+\d+\d+'
foreach($row in $rows){
    Out-File -FilePath $Output -Append -InputObject $row 
}
Write-Host "Output written to $Output"