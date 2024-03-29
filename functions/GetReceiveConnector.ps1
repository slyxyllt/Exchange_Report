$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    @{n='Bindings';e={[string]($_.Bindings)}},
    'AuthMechanism',
    'RequireTLS',
    'MaxMessageSize',
    'PermissionGroups',
    'Server',
    'Name',
    @{n='RemoteIPRanges';e={[string]($_.RemoteIPRanges)}}
)
Get-ReceiveConnector | Select-Object $Properties |Export-Csv $MainPath\Output\ReceiveConnector.csv -Encoding utf8 -NoTypeInformation