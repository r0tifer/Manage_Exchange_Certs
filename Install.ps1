# install.ps1 - Setup Exchange Certificate Manager in Exchange scripts folder

Write-Host "Starting installation of Exchange Certificate Manager..."

# Step 1: Validate $exscripts is available
if (-not (Test-Path env:exscripts)) {
    Write-Error "The `$exscripts environment variable is not defined. Please run this from the Exchange Management Shell."
    exit 1
}

$scriptUrl = "https://raw.githubusercontent.com/YourGitHubUsername/YourRepoName/main/Manage-ExchangeCert.ps1"
$targetFolder = $env:exscripts
$targetFile = Join-Path $targetFolder "Manage-ExchangeCert.ps1"

# Step 2: Confirm destination
Write-Host "Target Exchange scripts folder: $targetFolder"

# Step 3: Download script
Write-Host "Downloading Manage-ExchangeCert.ps1 from $scriptUrl..."
try {
    Invoke-WebRequest -Uri $scriptUrl -OutFile $targetFile -UseBasicParsing -ErrorAction Stop
    Write-Host "Script downloaded to $targetFile"
} catch {
    Write-Error "Download failed: $_"
    exit 1
}

# Step 4: Import script
Write-Host "Importing script..."
try {
    . $targetFile
    Write-Host "Script imported successfully."
} catch {
    Write-Error "Failed to import the script: $_"
    exit 1
}

Write-Host "Installation complete. You may now run Start-CertManager."
