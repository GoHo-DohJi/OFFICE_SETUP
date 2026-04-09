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
$CONFIG_PATH = "Office\CONFIGURATION.xml"
$SETUP_EXECUTABLE = "Office\Setup64.exe"
$ACTIVATION_SCRIPT_URL = "https://get.activated.win"

$XML_TEMPLATE = @"
<Configuration>

  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="ru-ru" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="OneNote" />
    </Product>
  </Add>

  <RemoveMSI />

  <Display Level="None" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />

  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />

  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="Ruler" Value="1" Type="REG_DWORD" App="word16" Id="L_WordRuler" /> <!-- RULER /> -->
    <User Key="software\microsoft\office\16.0\word\options" Name="MeasurementUnits" Value="2" Type="REG_DWORD" App="word16" Id="L_WordMeasurementUnits" /> <!-- MEASUREMENT UNITS /> -->
  </AppSettings>

</Configuration>
"@
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
        if ($Type -eq "Error") {
            [Console]::Error.WriteLine("${Title}: ${Message}")
        }
    }
}

function New-ConfigurationFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MountPath,
        
        [Parameter(Mandatory = $true)]
        [string[]]$SelectedProducts
    )
    
    [xml]$ConfigXml = $XML_TEMPLATE
    $ProductNode = $ConfigXml.Configuration.Add.Product
    
    # Добавляем все приложения кроме выбранных в ExcludeApp
    $AllApps = @("Word", "Excel", "PowerPoint", "Access")
    $AppsToExclude = $AllApps | Where-Object { $_ -notin $SelectedProducts }
    
    foreach ($App in $AppsToExclude) {
        $ExcludeNode = $ConfigXml.CreateElement("ExcludeApp")
        $ExcludeNode.SetAttribute("ID", $App) | Out-Null
        $ProductNode.AppendChild($ExcludeNode) | Out-Null
    }
    
    # Сохраняем конфигурацию
    $ConfigFilePath = Join-Path $MountPath $CONFIG_PATH
    $ConfigXml.Save($ConfigFilePath)
    
    return $ConfigFilePath
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
    # Mount image
    $MountedDrive = Mount-OfficeImage -ImagePath $ImagePath
    
    # Create configuration file
    $ConfigFilePath = New-ConfigurationFile -MountPath $MountedDrive -SelectedProducts $Products
    
    # Install Office
    Install-Office -MountPath $MountedDrive -ConfigPath $ConfigFilePath
    
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
