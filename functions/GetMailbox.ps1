$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'SamAccountName',
    'UseDatabaseQuotaDefaults',
    'Database',
    'ServerName',
    @{n='IssueWarningQuota';e={$value = 'unlimited';if(-not $_.IssueWarningQuota.IsUnlimited){$value = $_.IssueWarningQuota.Value.toMB()};$Value}},
    @{n='ProhibitSendQuota';e={$value = 'unlimited';if(-not $_.ProhibitSendQuota.IsUnlimited){$value = $_.ProhibitSendQuota.Value.toMB()};$Value}},
    @{n='ProhibitSendReceiveQuota';e={$value = 'unlimited';if(-not $_.ProhibitSendReceiveQuota.IsUnlimited){$value = $_.ProhibitSendReceiveQuota.Value.toMB()};$Value}},
    @{n='RecipientLimits';e={$value = 'unlimited';if(-not $_.RecipientLimits.IsUnlimited){$value = $_.RecipientLimits.Value};$Value}},
    @{n='MaxSendSize';e={$value = 'unlimited';if(-not $_.MaxSendSize.IsUnlimited){$value = $_.MaxSendSize.Value.toMB()};$Value}},
    @{n='MaxReceiveSizes';e={$value = 'unlimited';if(-not $_.MaxReceiveSizes.IsUnlimited){$value = $_.MaxReceiveSizes.Value.toMB()};$Value}},
    'RulesQuota',
    'DisplayName',
    'PrimarySmtpAddress',
    'DistinguishedName',
    'Identity',
    'Name',
    'ExchangeGuid'
    'Guid',
    'Alias'
    'WhenChanged',
    'WhenCreated'
)
Get-Mailbox -ResultSize unlimited | Select-Object $Properties |Export-Csv "$MainPath\Output\Mailbox.csv" -Encoding utf8 -NoTypeInformation #Unlimited | Select-Object $Propertiess |Export-Csv "$MainPaths\Data\Mailbox.csv" -Encoding utf8 -NoTypeInformation