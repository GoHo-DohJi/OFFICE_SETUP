# GET ADMIN #
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process "powershell" -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# VARIABLES #
$WORK_DIR     = "$env:TEMP\ohook"
$ZIP_PATH     = "$WORK_DIR\ohook.zip"
$DOWNLOAD_URL = "https://github.com/asdcorp/ohook/releases/download/0.5/ohook_0.5.zip"
$OFFICE_DIR   = "$env:PROGRAMFILES\Microsoft Office\root\vfs\System"
$SYSTEM_DIR   = "$env:SYSTEMROOT\System32"
$PRODUCT_KEY  = "VWCNX-7FKBD-FHJYG-XBR4B-88KC6" # ProPlus2024Retail ( https://massgrave.dev/manual_ohook_activation#office-2024 )

# STOP OFFICE #
Get-CimInstance -ClassName "Win32_Process" | Where-Object { $_.ExecutablePath -match "Microsoft Office|ClickToRun" } | Invoke-CimMethod -MethodName "Terminate" -ErrorAction "SilentlyContinue"
Start-Sleep -Seconds 2

# PREPARE WORKDIR #
if (Test-Path -Path "$WORK_DIR") { Remove-Item -Path "$WORK_DIR" -Recurse -Force }
New-Item -ItemType "Directory" -Path "$WORK_DIR" -Force | Out-Null

# GET OHOOK #
Invoke-RestMethod -Uri "$DOWNLOAD_URL" -OutFile "$ZIP_PATH"
Expand-Archive -Path "$ZIP_PATH" -DestinationPath "$WORK_DIR" -Force

# PATCH OFFICE #
New-Item -ItemType "SymbolicLink" -Path "$OFFICE_DIR\sppcs.dll" -Target "$SYSTEM_DIR\sppc.dll" -Force | Out-Null
Copy-Item -Path "$WORK_DIR\sppc64.dll" -Destination "$OFFICE_DIR\sppc.dll" -Force

# ACTIVATE OFFICE #
cscript //nologo "$SYSTEM_DIR\slmgr.vbs" /ipk "$PRODUCT_KEY"

# RESUME OFFICE SERVICE #
Start-Service -Name "ClickToRunSvc" -ErrorAction "SilentlyContinue" 

# CLEANUP #
if (Test-Path -Path "$WORK_DIR") { Remove-Item -Path "$WORK_DIR" -Recurse -Force }