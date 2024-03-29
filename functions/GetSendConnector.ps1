$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    @{n='AddressSpaces';e={[string]($_.AddressSpaces)}},
    'Enabled',
    'Name',
    'MaxMessageSize',
    'Port',
    @{n='SmartHosts';e={[string]($_.SmartHosts)}},
    @{n='SourceTransportServers';e={[string]($_.SourceTransportServers)}}
)
Get-SendConnector | Select-Object $Properties |Export-Csv $MainPath\Output\SendConnector.csv -Encoding utf8 -NoTypeInformation