$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'DeletedItemCount',
    'DisplayName',
    'DisconnectDate',
    'DisconnectReason',
    'ItemCount',
    'LastLoggedOnUserAccount',
    'LegacyDN',
    'MailboxGuid',
    @{n='TotalDeletedItemSize';e={$value = 0;if(-not $_.TotalDeletedItemSize.IsUnlimited){$value = $_.TotalDeletedItemSize.Value.toMB()};$Value}},
    @{n='TotalItemSize';e={$value = 0;if(-not $_.TotalItemSize.IsUnlimited){$value = $_.TotalItemSize.Value.toMB()};$Value}},
    'Database',
    'ServerName',
    'StorageLimitStatus'
)
Get-MailboxDatabase | Get-MailboxStatistics | Select-Object $Properties |Export-Csv $MainPath\Output\MailboxStatistics.csv -Encoding utf8 -NoTypeInformation