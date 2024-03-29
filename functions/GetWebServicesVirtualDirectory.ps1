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
Get-WebServicesVirtualDirectory | Select-Object $Properties |Export-Csv "$MainPath\Output\WebServicesVirtualDirectory.csv" -Encoding utf8 -NoTypeInformation 