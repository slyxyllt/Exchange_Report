[CmdletBinding()]
param (
    
    [Parameter(Mandatory = $false)]
    [DateTime]$Start=$(Get-Date $(Get-Date).ToString('yyyy-MM-dd')).AddDays(-8),
    #[DateTime]$Start='2019-03-08 00:00:00',

    [Parameter(Mandatory = $false)]
    [DateTime]$End=$(Get-Date $(Get-Date).ToString('yyyy-MM-dd')).AddDays(-1),
    #[DateTime]$End='2019-03-14 00:00:00',

    [Parameter(Mandatory = $false)]
    [string[]]$Servers = (Get-TransportServer | Select-Object -Expand Name),

    #Not Support PS 2.0,if Exchange has been Upgraded to above 2013,you can remove the # below
    #[ValidateRange(1,[int]::MaxValue)]
    [Parameter(Mandatory = $false)]
    [Int]$StepLength = 1,

    [ValidateSet('Day','Hour','Minute')]
    [Parameter(Mandatory = $false)]
    [string]$TimeUnit = "Day",

    #Not Support PS 2.0,if Exchange has been Upgraded to above 2013,you can remove the # below
    #[ValidateRange(1,[int]::MaxValue)]
    [Parameter(Mandatory = $false)]
    [Int]$Concurrency = 3
)
$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
$FlagName = "MessageTrackingLogStatistics"
Set-Location $MainPath 
Set-ADServerSettings -ViewEntireForest:$true
Function Split-Date{
    param(
        [Parameter(Mandatory = $true)]
        [DateTime]$Start,
        [Parameter(Mandatory = $true)]
        [DateTime]$End,
        #Not Support PS 2.0,if Exchange has been Upgraded to above 2013,you can remove the # below
        #[ValidateRange(1,[int]::MaxValue)]
        [Parameter(Mandatory = $false)]
        [Int]$StepLength = 1,
        [ValidateSet('Day','Hour','Minute')]
        [Parameter(Mandatory = $false)]
        [string]$TimeUnit = "Hour",
        
        [Parameter(Mandatory = $false)]
        [Int]$Concurrency = 3
        
    )
    if($Start -ge $End){
            Write-Error -Message "Start time ($Start) must be less then end time ($End)."
            return
    }
    $ExitFlag = $false
    $Objects = @()
    $Object = "" | Select-Object Start,End
    $Front = $Start
    do{
        switch($TimeUnit){
            Hour{
                $After = $(Get-Date $($Front).ToString('yyyy-MM-dd HH:00:00')).AddHours($StepLength)
            }
            Minute{
                $After = $(Get-Date $($Front).ToString('yyyy-MM-dd HH:mm:00')).AddMinutes($StepLength)
            }
            Day{
                $After = $(Get-Date $($Front).ToString('yyyy-MM-dd 00:00:00')).AddDays($StepLength)
            }
        }
        if($After -ge $End){
            Write-Debug "After ($After) has reached the end ($End)"
            $ExitFlag = $true
            $After = $End
        }
        $Object.Start = $Front
        $Object.End = $After
        $Objects += $Object
        $Object = "" | Select-Object Start,End
        $Front = $After
    } until($ExitFlag -eq $true)
    return $Objects
}
Function Group-Range{
    param(
        [Parameter( Mandatory = $true )]
        [System.Object]$InputObject,
        #Not Support PS 2.0,if Exchange has been Upgraded to above 2013,you can remove the # below
        #[ValidateRange(1,[int]::MaxValue)] 
        [Parameter( Mandatory = $true )]
        [int]$RangeStep = 12
    )
    $ExitFlag = $False
    [int]$LowRange = 0
    [int]$HighRange = $LowRange + $RangeStep
    $Total = $InputObject | Measure-Object | Select-Object -ExpandProperty Count
    Write-Verbose "There is $Total elements totally"
    $Count = 1
    $Object = "" | Select-Object SN,Group
    do{       
        if($HighRange -gt ($Total - 1)){
            $ExitFlag = $true
            $HighRange = $Total - 1
        }
        $Range = $LowRange..$HighRange 
        #Write-Host "$LowRange..$HighRange"
        $Object.SN = $Count
        $Object.Group = $InputObject[$Range]
        Write-Output $Object
        $LowRange = $HighRange + 1
        $HighRange = $LowRange + $RangeStep
        $Object = "" | Select-Object SN,Group
        $Count = $Count + 1
    
    } until($ExitFlag -eq $true)
}


$ScriptBlock ={
    param(
        [Parameter(Mandatory = $false)]
        [DateTime]$Start,

        [Parameter(Mandatory = $false)]
        [DateTime]$End,

        [Parameter(Mandatory = $false)]
        [string[]]$Servers,
        
        [Parameter(Mandatory = $true)]
        [string[]]$OutputPath
        
    )
    Add-PSSnapin Microsoft.Exchange.Management* 
    #$MainPath = Split-Path (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
    #Set-Location $MainPath 
    Set-ADServerSettings -ViewEntireForest:$true
    $FlagName = 'MessageTrackingLogStatistics'
    $VerbosePreference = 'continue'
    $Properties = "Server","Start","End","UserReceiveMessageCount","UserSendMessageCount","SMTPReceiveMessageCount","SMTPSendMessageCount","AvgUserReceiveMessageSize","AvgUserSendMessageSize","AvgSMTPReceiveMessageSize","AvgSMTPSendMessageSize","TotalInternalMessageCount","AvgInternalMessageSize","TotalSMTPMessageCount","AvgSMTPMessageSize","TotalMessageCount","RecipientCount"
    if($Start -ge $End){
        Write-Error -Message "Start time ($Start) must be less then the end time ($End)."
        return
    }
    $TransportServers = $Servers | Foreach-Object { Get-TransportServer $_ }
    if(-not $TransportServers){
        Write-Warning "There is no vaild Transport server in {0}" -f ($Servers -join ";")
        exit
    }
    $MessageReport = @()
    
    foreach($TransportServer in $TransportServers){        
        $MessageObj = "" | Select-Object $Properties
        $MessageObj.Server = $TransportServer.name
        $MessageObj.Start = $Start
        $MessageObj.End = $End
        $RecipientCount = 0
        $CountDeliver = 0
        $CountSubmit = 0
        $CountSMTPReceive = 0
        $CountSMTPSend = 0
        $DELIVERTotalBytes = 0
        $SUBMITTotalBytes = 0
        $ReceiveSMTPTotalBytes = 0
        $SentSMTPTotalBytes = 0
        Write-Verbose "Get Message tracking log from $Start to $End on Server $($TransportServer.name)"
        Get-MessageTrackingLog -Start $Start -End $End -ResultSize unlimited -Server $TransportServer.name | ? { $_.eventid -like "Deliver" -or $_.eventid -like "Submit" -or $_.eventid -like "Send" -or $_.eventid -like "Receive" } | % {

            if ($_.EventId -eq "DELIVER" -and $_.Source -eq "STOREDRIVER")
            {
                $CountDeliver++
                $DELIVERTotalBytes += $_.TotalBytes
            }
            
            if ($_.EventId -eq "Receive" -and $_.Source -eq "STOREDRIVER")
            {
                $CountSubmit++
                $SUBMITTotalBytes += $_.TotalBytes
            }
            if ($_.EventId -eq "Receive" -and $_.Source -eq "SMTP")
            {
                $CountSMTPReceive++
                $ReceiveSMTPTotalBytes += $_.TotalBytes
            }
            if ($_.EventId -eq "Send" -and $_.Source -eq "SMTP")
            {
                $CountSMTPSend++
                $SentSMTPTotalBytes += $_.TotalBytes
            }
            $RecipientCount += $_.RecipientCount
        }
        
        $MessageObj.UserReceiveMessageCount =  $CountDeliver
        $MessageObj.UserSendMessageCount =  $CountSubmit
        $MessageObj.SMTPReceiveMessageCount =  $CountSMTPReceive
        $MessageObj.SMTPSendMessageCount =  $CountSMTPSend
        
        if ($CountDeliver -gt 0)
        {
            $MessageObj.AvgUserReceiveMessageSize =  [int](("{0:n2}" -f ($DELIVERTotalBytes/($CountDeliver * 1024))))
        }
        else
        {
            $MessageObj.AvgUserReceiveMessageSize = 0
        }
        
        if ($CountSubmit -gt 0)
        {
            $MessageObj.AvgUserSendMessageSize =  [int](("{0:n2}" -f ($SUBMITTotalBytes/($CountSubmit * 1024))))
        }
        else
        {
            $MessageObj.AvgUserSendMessageSize =  0
        }
        
        if ($CountSMTPReceive -gt 0)
        {
            $MessageObj.AvgSMTPReceiveMessageSize =  [int](("{0:n2}" -f ($ReceiveSMTPTotalBytes/($CountSMTPReceive * 1024))))
        }
        else
        {
            $MessageObj.AvgSMTPReceiveMessageSize = 0
        }
        
        if ($CountSMTPSend -gt 0)
        {
            $MessageObj.AvgSMTPSendMessageSize =  [int](("{0:n2}" -f ($SentSMTPTotalBytes/($CountSMTPSend * 1024))))
        }
        else
        {
            $MessageObj.AvgSMTPSendMessageSize = 0
        }
        
        $MessageObj.TotalInternalMessageCount =  ($CountDeliver + $CountSubmit)
        
        if (($CountDeliver + $CountSubmit) -gt 0)
        {
            $MessageObj.AvgInternalMessageSize =  [int](("{0:n2}" -f (($DELIVERTotalBytes + $SUBMITTotalBytes)/(($CountDeliver + $CountSubmit) * 1024))))
        }
        else
        {
            $MessageObj.AvgInternalMessageSize = 0
        }
        
        $MessageObj.TotalSMTPMessageCount =  ($CountSMTPReceive + $CountSMTPSend)
        
        if (($CountSMTPReceive + $CountSMTPSend) -gt 0)
        {
            $MessageObj.AvgSMTPMessageSize =  [int](("{0:n2}" -f (($ReceiveSMTPTotalBytes + $SentSMTPTotalBytes)/(($CountSMTPReceive + $CountSMTPSend) * 1024))))
        }
        else
        {
            $MessageObj.AvgSMTPMessageSize = 0
        }
        
        $MessageObj.TotalMessageCount =  ($CountDeliver + $CountSubmit + $CountSMTPReceive + $CountSMTPSend)
        $MessageObj.RecipientCount =  $RecipientCount
        Write-Output $MessageObj
        $MessageReport = @($MessageReport + $MessageObj)
     
    }
    $MessageReportPath = "$OutputPath\{0}{1}{2}.csv" -f $FlagName,$Start.tostring("-yyyyMMddHHmmss"),$End.tostring("-yyyyMMddHHmmss")
    $MessageReport | Export-Csv $MessageReportPath -NoTypeInformation -Encoding UTF8 -Force 
}
$VerbosePreference = 'continue'
Write-Verbose "Split date time from $Start to $End"
$DateArrays = Split-Date -Start $Start -End $End -TimeUnit $TimeUnit -StepLength $StepLength
Write-Verbose "Group date time array by step length $($Concurrency-1)"
$DateArraysGroups = Group-Range $DateArrays -RangeStep ($Concurrency-1)
$AggreateReport = @()
foreach($DateArraysGroup in $DateArraysGroups){
    foreach($DateArray in $DateArraysGroup.Group){
        Start-Job -ScriptBlock $ScriptBlock -ArgumentList ($DateArray.Start,$DateArray.End,$Servers,"$MainPath\Output")
    }
    Get-Job | Wait-Job
    $AggreateReport += Get-Job | Receive-Job
}
$AggreateReportPath = "$MainPath\Output\{0}{1}{2}.csv" -f $FlagName,$Start.tostring("-yyyyMMddHHmmss"),$End.tostring("-yyyyMMddHHmmss")
$AggreateReport | Export-Csv $AggreateReportPath -NoTypeInformation -Encoding UTF8 -Force 
