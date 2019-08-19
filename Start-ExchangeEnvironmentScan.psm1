$script:MainPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$script:FuntionPath = Join-Path -Parent $script:MainPath -Path functions 
$Functions = @( Get-ChildItem -Path $script:FuntionPath -Include *.ps1 -ErrorAction SilentlyContinue )
Function Start-ExchangeEnvironmentScan {
    param(
        
    )

}