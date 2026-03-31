### IMPORT ### 
. "$PSScriptRoot\utils.ps1"
. "$PSScriptRoot\VARIABLES.ps1"



### MAIN ###
## OFFICE CHECK ##
if (Test-Path "$OFFICE_DIR") {
    if (-not (Write-Centered -Text "$BANNER_REINSTALL" -TextColor "DarkYellow" -Prompt_No)) { exit 0 }
}


## GET IMAGE PATH ##
while ($true) {
    Write-Centered -Text "$BANNER_GET_IMAGE_PATH" -DownloadPrompt

    $IMAGE_PATH = Get-Image-Path
    
    if ($IMAGE_PATH) { break }
}


## PRODUCT SELECTOR ##
$script:ExcludedProducts = Write-Centered -Text "$BANNER_PRODUCT_SELECTOR" -ProductSelector
## CONFIGURATION ## 
[xml]$xml = Get-Content "$CONFIGURATION_TEMPLATE_PATH" -Encoding "UTF8"
$ExcludedProducts.ForEach({ $xml.Configuration.Add.Product.InnerXml += "<ExcludeApp ID=`"$($_.ID)`"/>" })
$xml.Save($CONFIGURATION_PATH)


## SETUP ##
Write-Centered -Text "$BANNER_SETUP_OFFICE"

try {
    $DRIVE_LETTER = (Mount-DiskImage -ImagePath "$IMAGE_PATH" | Get-Volume).DriveLetter                          # IMAGE MOUNT #
    # if (-not (Test-Path "${DriveLetter}:\Office\Setup64.exe")) {
    #     Write-Centered -Text "❌ ⚠ NOT A VALID OFFICE IMAGE!"
    #     Dismount-DiskImage -ImagePath $IMAGE_PATH
    #     continue
    # }
    Start-Process "${DRIVE_LETTER}:\Office\Setup64.exe" -ArgumentList "/configure `"$CONFIGURATION_PATH`"" -Wait # SETUP #
    & ([ScriptBlock]::Create((irm https://get.activated.win))) /para /Ohook                                      # ACTIVATION #
    Dismount-DiskImage -ImagePath "$IMAGE_PATH"                                                                  # IMAGE DISMOUNT #
    
    Write-Centered -Text "$BANNER_SETUP_SUCCESS"
} catch {
    Dismount-DiskImage -ImagePath "$IMAGE_PATH" 

    $errorText = "$BANNER_SETUP_ERROR" + "$($_.Exception.Message)"
    Write-Centered -Text "$errorText"
}

$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null