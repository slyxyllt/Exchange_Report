$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'InternalAuthenticationMethods',
    'ExternalAuthenticationMethods',
    'InternalUrl',
    'ExternalUrl',
    'Server'
)
Get-ActiveSyncVirtualDirectory | Select-Object $Properties |Export-Csv "$MainPath\Output\ActiveSyncVirtualDirectory.csv" -Encoding utf8 -NoTypeInformation 