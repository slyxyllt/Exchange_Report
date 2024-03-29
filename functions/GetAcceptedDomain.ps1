$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
$OutputFolderPath = Join-Path -Parent $MainPath -Path Output
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'DomainName',
    'DomainTypes',
    'WhenCreated',
    'WhenChanged',
    'Default'  
)
$OutputCSVPath = Join-Path 
Get-AcceptedDomain | Select-Object $Properties |Export-Csv $MainPath\Output\AcceptedDomain.csv -Encoding utf8 -NoTypeInformation