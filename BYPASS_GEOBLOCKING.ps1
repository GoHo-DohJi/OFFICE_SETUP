### SELF-ELEVATE ###
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process -FilePath "powershell" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}



### VARIABLES ###
$PATH = "HKCU:\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs"
$NAME = "CountryCode"
$VALUE = "std::wstring|US"



### MAIN ###
if (-not (Test-Path -Path "$PATH")) { New-Item -Path "$PATH" -Force | Out-Null }
Set-ItemProperty -Path "$PATH" -Name "$NAME" -Value "$VALUE" -Type "String" -Force