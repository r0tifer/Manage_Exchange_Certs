# install.ps1 - Deploy ManageExchangeCert.ps1 to user-selected location

function Show-InstallOptions {
    Write-Host "Choose install destination:"
    Write-Host "1. System-wide PowerShell Module Folder (requires admin)"
    Write-Host "2. User PowerShell Module Folder"
    Write-Host "3. Exchange Script Library (`$exscripts)` - WARNING: may be overwritten by updates"
    Write-Host "4. Custom Folder"
    do {
        $choice = Read-Host "Enter your choice [1-4]"
    } while ($choice -notmatch '^[1-4]$')
    return $choice
}

function Get-InstallPath($choice) {
    switch ($choice) {
        '1' {
            $path = "C:\Program Files\WindowsPowerShell\Modules\ManageExchangeCert"
        }
        '2' {
            $path = Join-Path -Path $HOME -ChildPath "Documents\WindowsPowerShell\Modules\ManageExchangeCert"
        }
        '3' {
            if (-not (Test-Path env:exscripts)) {
                throw "Exchange script path (`$exscripts) not available. Please run this from Exchange Management Shell."
            }
            Write-Warning "WARNING: Installing to Exchange Script Library may be dangerous and overwritten by updates."
            $confirm = Read-Host "Type 'YES' to confirm this choice"
            if ($confirm -ne 'YES') {
                throw "Installation cancelled."
            }
            $path = Join-Path $env:exscripts "ManageExchangeCert"
        }
        '4' {
            $custom = Read-Host "Enter full custom path"
            $path = $custom
        }
    }

    if (-not (Test-Path $path)) {
        Write-Host "The following folder does not exist:"
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

# Main install logic
Write-Host "Starting ManageExchangeCert.ps1 installer..."
$scriptUrl = "https://raw.githubusercontent.com/YourGitHubUser/YourRepoName/main/Manage-ExchangeCert.ps1"

try {
    $choice = Show-InstallOptions
    $destination = Get-InstallPath -choice $choice
    $targetFile = Join-Path $destination "Manage-ExchangeCert.ps1"

    Write-Host "Downloading script from $scriptUrl..."
    Invoke-WebRequest -Uri $scriptUrl -OutFile $targetFile -UseBasicParsing -ErrorAction Stop
    Write-Host "Script downloaded to: $targetFile"

    Write-Host "Importing script..."
    . $targetFile
    Write-Host "Script successfully imported. You can now run: Start-CertManager"
} catch {
    Write-Error "Installation failed: $_"
    exit 1
}
