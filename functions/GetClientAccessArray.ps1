$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Fqdn',
    'Site',
    'Site',
    @{n='Members';e={$_.members -join ";"}},
    'Name'
)
Get-ClientAccessArray | Select-Object $Properties |Export-Csv "$MainPath\Output\ClientAccessArray.csv" -Encoding utf8 -NoTypeInformation 