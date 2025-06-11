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
Write-Host "Starting Manage-ExchangeCert.ps1 installation..."

$scriptUrl = "https://raw.githubusercontent.com/r0tifer/Manage_Exchange_Certs/main/Manage-ExchangeCert.ps1"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "You are not running PowerShell as Administrator. Some install locations may fail."
}

try {
    $choice = Show-InstallOptions
    $destination = Get-InstallPath -choice $choice
    $targetFile = Join-Path $destination "Manage-ExchangeCert.ps1"

    Write-Host "`nDownloading Manage-ExchangeCert.ps1 from GitHub..."
    try {
        Invoke-WebRequest -Uri $scriptUrl -OutFile $targetFile -UseBasicParsing -ErrorAction Stop
        Write-Host "Script downloaded to: $targetFile"
    } catch {
        if ($_.Exception.Message -match 'Access.*denied') {
            Write-Error "Access denied writing to: $targetFile"
            Write-Host "Please re-run this script in an elevated PowerShell session (Run as Administrator)."
            Read-Host "Press Enter to exit"
            exit 1
        } else {
            Write-Error "Download failed: $_"
            Read-Host "Press Enter to exit"
            exit 1
        }
    }

    Write-Host "Importing script..."
    try {
        . $targetFile
        Write-Host "Script successfully imported."
    } catch {
        Write-Error "Failed to import the script: $_"
        Read-Host "Press Enter to exit"
        exit 1
    }

    # Final Countdown Launch
    Write-Host "`nLaunching Manage-ExchangeCert.ps1 in 5 seconds..."
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