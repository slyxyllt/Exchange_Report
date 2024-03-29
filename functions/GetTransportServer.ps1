$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'MessageTrackingLogEnabled',
    'MessageTrackingLogMaxDirectorySize',
    'MessageTrackingLogMaxAge',
    'MessageTrackingLogPath'
)
Get-TransportServer | Select-Object $Properties |Export-Csv "$MainPath\Output\TransportServer.csv" -Encoding utf8 -NoTypeInformation 