$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'Description',
    'State',    
    'WhenChanged'
)
Get-TransportRule | Select-Object $Properties |Export-Csv $MainPath\Output\TransportRule.csv -Encoding utf8 -NoTypeInformation