# install.ps1 - Deploy Manage-ExchangeCert.ps1 from GitHub

function Show-InstallOptions {
    Write-Host "Choose install destination:"
    Write-Host "1. System-wide PowerShell Module Folder (requires admin)"
    Write-Host "2. User PowerShell Module Folder"
    Write-Host "3. Exchange Script Library (`$exscripts) - WARNING: May be overwritten by Exchange updates"
    Write-Host "4. Custom Folder"
    do {
        $choice = Read-Host "Enter your choice [1-4]"
    } while ($choice -notmatch '^[1-4]$')
    return $choice
}

function Get-InstallPath($choice) {
    switch ($choice) {
        '1' { $path = "C:\Program Files\WindowsPowerShell\Modules\ManageExchangeCert" }
        '2' { $path = Join-Path $HOME "Documents\WindowsPowerShell\Modules\ManageExchangeCert" }
        '3' {
            if (-not (Test-Path env:exscripts)) {
                throw "Exchange script path (`$exscripts) not available. Please run from the Exchange Management Shell."
            }
            Write-Warning "WARNING: Installing to Exchange Script Library may be dangerous and overwritten by updates."
            $confirm = Read-Host "Type 'YES' to confirm this choice"
            if ($confirm -ne 'YES') {
                throw "Installation cancelled."
            }
            $path = Join-Path $env:exscripts "ManageExchangeCert"
        }
        '4' {
            $path = Read-Host "Enter full custom path"
        }
    }

    if (-not (Test-Path $path)) {
        Write-Host "`nThe following folder does not exist:"
        Write-Host "  $path"
        $confirm = Read-Host "Do you want to create this folder? (y/n)"
        if ($confirm -notmatch '^(y|yes)$') {
            throw "Installation cancelled by user."
        }
        Write-Host "Creating folder: $path"
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }

    return $path
}

# --- Main Installer Logic ---
Clear-Host
Write-Host "Starting Manage-ExchangeCert module installation..."

$repoZipUrl = "https://github.com/r0tifer/Manage_Exchange_Certs/archive/refs/heads/main.zip"
$tempZipPath = Join-Path $env:TEMP "ManageExchangeCert-main.zip"
$tempExtractPath = Join-Path $env:TEMP "Manage_Exchange_Certs-main"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "You are not running PowerShell as Administrator. Some install locations may fail."
}

try {
    $choice = Show-InstallOptions
    $destination = Get-InstallPath -choice $choice
    $finalModulePath = Join-Path $destination "ManageExchangeCert"

    # Clean up any prior download
    if (Test-Path $tempZipPath) { Remove-Item $tempZipPath -Force }
    if (Test-Path $tempExtractPath) { Remove-Item $tempExtractPath -Recurse -Force }

    Write-Host "`nDownloading full module from GitHub..."
    Invoke-WebRequest -Uri $repoZipUrl -OutFile $tempZipPath -UseBasicParsing -ErrorAction Stop

    Write-Host "Extracting archive..."
    # Check if Expand-Archive is available, fallback if not
    if (Get-Command Expand-Archive -ErrorAction SilentlyContinue) {
        Expand-Archive -Path $tempZipPath -DestinationPath $env:TEMP -Force
    } else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZipPath, $env:TEMP)
    }

    Clear-Host
    Write-Host "Installing to: $finalModulePath"
    if (Test-Path $finalModulePath) {
        Remove-Item -Recurse -Force $finalModulePath
    }
    Move-Item -Path $tempExtractPath -Destination $finalModulePath

    Remove-Item $tempZipPath -Force

    $psm1Path = Join-Path $finalModulePath "ManageExchangeCert.psm1"
    if (-not (Test-Path $psm1Path)) {
        throw "Module file not found: $psm1Path"
    }

    Write-Host "Importing module..."
    Import-Module $psm1Path -Force
    Write-Host "Module successfully imported."
    Start-Sleep -Seconds 3
    Clear-Host

    Write-Host "`nLaunching Manage-ExchangeCert module in 5 seconds..."
    for ($i = 5; $i -ge 1; $i--) {
        Write-Host "$i..."
        Start-Sleep -Seconds 1
    }

    Clear-Host
    Start-CertManager

} catch {
    Write-Error "Installation failed: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
