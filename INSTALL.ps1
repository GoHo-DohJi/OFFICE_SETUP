### VARIABLES ###
## LOCAL ##
$REPOSITORY_NAME = "OFFICE_SETUP"
$EXTRACTION_DIR = "$env:TEMP"


## CONST ##
$REPOSITORY_URL = "https://github.com/GoHo-DohJi/$REPOSITORY_NAME/archive/refs/heads/main.zip"
$ZIP_PATH = "$EXTRACTION_DIR\$REPOSITORY_NAME.zip"
$REPOSITORY_PATH = "$EXTRACTION_DIR\${REPOSITORY_NAME}-main\MAIN.bat"



### MAIN ###
Clear-Host
try {
    ## DOWNLOAD ##
    Write-Host "[*] DOWNLOADING: $REPOSITORY_URL" -ForegroundColor "DarkGray"
    curl.exe --silent --show-error --location "$REPOSITORY_URL" --output "$ZIP_PATH"
    

    ## EXTRACT ##
    Write-Host "[*] EXTRACTING TO: $EXTRACTION_DIR" -ForegroundColor "Cyan"
    Expand-Archive -Path "$ZIP_PATH" -DestinationPath "$EXTRACTION_DIR" -Force
    Remove-Item "$ZIP_PATH" -Force
    

    ## RUN SCRIPT ##
    Write-Host "[✓] SUCCESS" -ForegroundColor "Green"
    & "$REPOSITORY_PATH"
    exit 0
} catch {
    Write-Host "[x] ERROR: $($_.Exception.Message)" -ForegroundColor "Red"
    Write-Host "$_" -ForegroundColor "DarkRed"
    exit 1
}
















