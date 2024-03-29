$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'InternalAuthenticationMethods',
    'Server',
    'Name',
    'DefaultDomain',
    'InternalUrl',
    'ExternalUrl'
)
Get-EcpVirtualDirectory | Select-Object $Properties |Export-Csv "$MainPath\Output\EcpVirtualDirectory.csv" -Encoding utf8 -NoTypeInformation 