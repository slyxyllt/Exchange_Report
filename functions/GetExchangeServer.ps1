$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'Edition',
    'OrganizationalUnit',
    'Site',
    'ServerRole',
    'RemainingTrialPeriod',
    'WhenCreated',
    'WhenChanged'
)
Get-ExchangeServer | Select-Object $Properties |Export-Csv "$MainPath\Output\ExchangeServer.csv" -Encoding utf8 -NoTypeInformation 