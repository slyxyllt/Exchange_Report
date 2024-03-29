$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'DeviceOS',
    'DeviceType',
    'FirstSyncTime',
    'UserDisplayName',
    'DeviceOSLanguage'
)
Get-ActiveSyncDevice -ResultSize unlimited | Select-Object $Properties |Export-Csv $MainPath\Output\ActiveSyncDevice.csv -Encoding utf8 -NoTypeInformation