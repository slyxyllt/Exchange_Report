@echo off
set ScriptName=. '%~dp0\GetExchangeServer.ps1'
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -command ". '%ExchangeInstallPath%\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; %ScriptName% "