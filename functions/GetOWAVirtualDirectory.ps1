$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'LogonFormat',
    'ServerName',
    'Name',
    'DefaultDomain',
    'InternalUrl',
    'ExternalUrl'
)
Get-OwaVirtualDirectory | Select-Object $Properties |Export-Csv "$MainPath\Output\OwaVirtualDirectory.csv" -Encoding utf8 -NoTypeInformation 