$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
$Properties = @(
    'Name',
    'Recipient',
    'JournalEmailAddress',    
    'Scope',
    'Enabled'
)
Get-JournalRule | Select-Object $Properties |Export-Csv $MainPath\Output\JournalRule.csv -Encoding utf8 -NoTypeInformation