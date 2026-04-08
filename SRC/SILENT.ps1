<#
.SYNOPSIS
    Silent Office installation and activation script
.DESCRIPTION
    Installs Office from IMG file with custom product selection and activates using Ohook
.PARAMETER ImagePath
    Full path to Office IMG file
.PARAMETER Products
    Comma-separated list of products to install (Word, Excel, PowerPoint, Access)
.EXAMPLE
    .\InstallOffice.ps1 -ImagePath "C:\Office.img" -Products "Word, Excel, PowerPoint, Access"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ImagePath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Word", "Excel", "PowerPoint", "Access")]
    [string[]]$Products = @("Word", "Excel", "PowerPoint", "Access")
)

#region CONSTANTS
$CONFIG_TEMPLATE_PATH = Join-Path $PSScriptRoot "CONFIGURATION_TEMPLATE.xml"
$CONFIG_PATH = Join-Path $PSScriptRoot "CONFIGURATION.xml"
$SETUP_EXECUTABLE = "Office\Setup64.exe"
$ACTIVATION_SCRIPT_URL = "https://get.activated.win"

$PRODUCT_MAP = @{
    "Word"       = "Word"
    "Excel"      = "Excel"
    "PowerPoint" = "PowerPoint"
    "Access"     = "Access"
}

$ALL_EXCLUDABLE_APPS = @("Word", "Excel", "PowerPoint", "Access", "Outlook", "OneNote", "Publisher", "Lync", "OneDrive")
#endregion

#region FUNCTIONS
function Show-Notification {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Error")]
        [string]$Type = "Info"
    )
    
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $Icon = if ($Type -eq "Error") { 
            [System.Windows.Forms.ToolTipIcon]::Error 
        } else { 
            [System.Windows.Forms.ToolTipIcon]::Info 
        }
        
        $Notification = New-Object System.Windows.Forms.NotifyIcon
        $Notification.Icon = [System.Drawing.SystemIcons]::Information
        $Notification.BalloonTipIcon = $Icon
        $Notification.BalloonTipTitle = $Title
        $Notification.BalloonTipText = $Message
        $Notification.Visible = $true
        $Notification.ShowBalloonTip(10000)
        
        Start-Sleep -Seconds 2
        $Notification.Dispose()
    } catch {
        # Fallback to console if notification fails
        if ($Type -eq "Error") {
            [Console]::Error.WriteLine("${Title}: ${Message}")
        }
    }
}

function Get-ExcludedApps {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$SelectedProducts
    )
    
    $NormalizedProducts = $SelectedProducts | ForEach-Object { $_.Trim() }
    $ExcludedApps = $ALL_EXCLUDABLE_APPS | Where-Object { $_ -notin $NormalizedProducts }
    
    return $ExcludedApps
}

function Update-Configuration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $true)]
        [string[]]$ExcludedApps
    )
    
    if (-not (Test-Path $TemplatePath)) {
        throw "Configuration template not found at: $TemplatePath"
    }
    
    [xml]$ConfigXml = Get-Content $TemplatePath -Encoding UTF8 -ErrorAction Stop
    
    $ProductNode = $ConfigXml.Configuration.Add.Product
    if ($null -eq $ProductNode) {
        throw "Invalid configuration template: Product node not found"
    }
    
    $ExistingExclusions = $ProductNode.SelectNodes("ExcludeApp") | ForEach-Object { $_.ID }
    
    foreach ($App in $ExcludedApps) {
        if ($App -notin $ExistingExclusions) {
            $ExcludeNode = $ConfigXml.CreateElement("ExcludeApp")
            $ExcludeNode.SetAttribute("ID", $App) | Out-Null
            $ProductNode.AppendChild($ExcludeNode) | Out-Null
        }
    }
    
    $ConfigXml.Save($OutputPath)
}

function Mount-OfficeImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )
    
    $MountedImage = Mount-DiskImage -ImagePath $ImagePath -PassThru -ErrorAction Stop
    $DriveLetter = ($MountedImage | Get-Volume).DriveLetter
    
    if ([string]::IsNullOrEmpty($DriveLetter)) {
        throw "Failed to retrieve drive letter for mounted image"
    }
    
    return "${DriveLetter}:"
}

function Install-Office {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MountPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    $SetupPath = Join-Path $MountPath $SETUP_EXECUTABLE
    
    if (-not (Test-Path $SetupPath)) {
        throw "Setup executable not found at: $SetupPath"
    }
    
    $ProcessArgs = "/configure `"$ConfigPath`""
    $Process = Start-Process -FilePath $SetupPath -ArgumentList $ProcessArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
    
    if ($Process.ExitCode -ne 0) {
        throw "Office installation failed with exit code: $($Process.ExitCode)"
    }
}

function Invoke-Activation {
    try {
        $ActivationScript = Invoke-RestMethod -Uri $ACTIVATION_SCRIPT_URL -UseBasicParsing -ErrorAction Stop
        $ScriptBlock = [ScriptBlock]::Create($ActivationScript)
        & $ScriptBlock /Ohook
    } catch {
        throw "Activation failed: $($_.Exception.Message)"
    }
}

function Dismount-OfficeImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )
    
    Dismount-DiskImage -ImagePath $ImagePath -ErrorAction SilentlyContinue | Out-Null
}
#endregion

#region MAIN
$ErrorActionPreference = "Stop"
$MountedDrive = $null

try {
    # Calculate excluded apps
    $ExcludedApps = Get-ExcludedApps -SelectedProducts $Products
    
    # Update configuration file
    Update-Configuration -TemplatePath $CONFIG_TEMPLATE_PATH -OutputPath $CONFIG_PATH -ExcludedApps $ExcludedApps
    
    # Mount image
    $MountedDrive = Mount-OfficeImage -ImagePath $ImagePath
    
    # Install Office
    Install-Office -MountPath $MountedDrive -ConfigPath $CONFIG_PATH
    
    # Activate Office
    Invoke-Activation
    
    # Dismount image
    Dismount-OfficeImage -ImagePath $ImagePath
    
    # Success notification
    Show-Notification -Title "Office Installation" -Message "Office installed and activated successfully!" -Type "Info"
    
    exit 0
    
} catch {
    # Cleanup on error
    if ($null -ne $MountedDrive) {
        Dismount-OfficeImage -ImagePath $ImagePath
    }
    
    # Error notification
    $ErrorMessage = "Installation failed: $($_.Exception.Message)`n`nDetails: $($_.ScriptStackTrace)"
    Show-Notification -Title "Office Installation Error" -Message $ErrorMessage -Type "Error"
    
    exit 1
}
#endregion