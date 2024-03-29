$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'OutlookAnywhereEnabled',
    'AutoDiscoverServiceInternalUri'
)
Get-ClientAccessServer | Select-Object $Properties |Export-Csv "$MainPath\Output\ClientAccessServer.csv" -Encoding utf8 -NoTypeInformation 