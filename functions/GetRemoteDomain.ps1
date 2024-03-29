$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'DomainName',
    'DomainTypes',
    'WhenCreated',
    'WhenChanged',
    'Default'  
)
Get-RemoteDomain | Select-Object $Properties |Export-Csv $MainPath\Output\RemoteDomain.csv -Encoding utf8 -NoTypeInformation