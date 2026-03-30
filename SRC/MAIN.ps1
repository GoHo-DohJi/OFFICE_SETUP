### IMPORT ### 
. "$PSScriptRoot\utils.ps1"
. "$PSScriptRoot\UI.ps1"



### VARIABLES ###
## CONST ## 
$OFFICE_DIR = "$env:PROGRAMFILES\Microsoft Office"
$CONFIGURATION_TEMPLATE_PATH = "$PSScriptRoot\CONFIGURATION_TEMPLATE.xml"
$CONFIGURATION_PATH = "$PSScriptRoot\CONFIGURATION.xml"


### MAIN ###
## OFFICE CHECK ##
if (Test-Path "$OFFICE_DIR") {
    if (-not (Write-Centered -Text "$BANNER_REINSTALL" -TextColor "DarkYellow" -Prompt_No)) { exit 0 }
}


## GET IMAGE PATH ##
while ($true) {
    Write-Centered -Text "$BANNER_GET_IMAGE_PATH"
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null

    $IMAGE_PATH = Get-Image-Path
    
    if ($IMAGE_PATH) { break }
    
    if (-not (Write-Centered -Text "$BANNER_IMAGE_PATH_ERROR" -Prompt_Yes)) {
        ## EXIT ## 
        Write-Centered -Text "$BANNER_EXIT"
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
        return
    }
}


## PRODUCT SELECTOR ##
$script:ExcludedProducts = Write-Centered -Text "$BANNER_PRODUCT_SELECTOR" -ProductSelector


## SETUP ##
Write-Centered -Text "$BANNER_SETUP_OFFICE"

try {
    ## CONFIGURATION ## 
    [xml]$xml = Get-Content "$CONFIGURATION_TEMPLATE_PATH" -Encoding "UTF8"
    $ExcludedProducts.ForEach({
        $xml.Configuration.Add.Product.InnerXml += "<ExcludeApp ID=`"$($_.ID)`"/>"
    })
    $xml.Save($CONFIGURATION_PATH)

    ## OFFICE ##
    $DRIVE_LETTER = (Mount-DiskImage -ImagePath "$IMAGE_PATH" | Get-Volume).DriveLetter                          # IMAGE MOUNT #
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