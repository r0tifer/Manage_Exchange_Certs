<#
.SYNOPSIS
    Generates a Certificate Signing Request (CSR) from an existing Exchange certificate.

.DESCRIPTION
    Uses the provided thumbprint to retrieve the Exchange certificate,
    then generates a CSR, writes it to disk, and opens it for review.

.PARAMETER Thumbprint
    The thumbprint of the existing Exchange certificate.

.PARAMETER OutputPath
    The full path to write the CSR file (e.g. C:\cert\2025.req). Defaults to C:\cert\ExchangeCSR.req

.PARAMETER Server
    The Exchange server to generate the request on. Defaults to local server if omitted.
#>
param (
    [string]$Thumbprint,
    [string]$OutputPath,
    [string]$Server
)

function Show-MainMenu {
    Write-Host "`nMain Menu"
    Write-Host "1. Renew Certificate"
    Write-Host "2. Replace Certificate"
    Write-Host "3. Revoke Certificate"
    Write-Host "4. Delete Certificate"
    Write-Host "5. Export Certificate (.pfx)"
    Write-Host "6. Select Different Certificate"
    Write-Host "7. Resume Pending Renewal"
    Write-Host "8. Switch Exchange Server"
    Write-Host "9. Import New Certificate"
    Write-Host "10. Exit"
}

function Start-CertManager {
    Clear-Host
    $global:Server = $Server
    $global:AllExchangeServers = $AllExchangeServers
    $global:cert = $null
    do {
        try {
            if (-not $cert) {
                $cert = Prompt-CertSelection
            }
            Show-MainMenu
            $mainAction = Read-Host "Enter your choice"

            switch ($mainAction) {
                '1' {
                    Clear-Host 
                    Renew-Certificate -cert $cert -server $Server -allServers $AllExchangeServers 
                }
                '2' {
                    Clear-Host 
                    Replace-Certificate -cert $cert -server $Server -allServers $AllExchangeServers 
                }
                '3' {
                    Clear-Host 
                    Revoke-Certificate -cert $cert; $cert = $null 
                }
                '4' {
                    Clear-Host 
                    Delete-Certificate -cert $cert; $cert = $null 
                }
                '5' {
                    Clear-Host 
                    Export-Certificate -cert $cert 
                }
                '6' {
                    Clear-Host
                    $cert = Prompt-CertSelection
                }
                '7' {
                    Clear-Host 
                    Resume-PendingCertificate 
                }
                '8' {
                    Clear-Host
                    $serverSelection = Select-ExchangeServer
                    $Server = $serverSelection.Primary
                    $AllExchangeServers = $serverSelection.All
                    $cert = $null
                }
                '9' {
                    Clear-Host
                    $newCertPath = Read-Host "Enter full path to new certificate file"
                    if (-not (Test-Path $newCertPath)) {
                        Write-Warning "File not found."
                        pause; return
                    }
                    $cert = Prompt-CertSelection
                    Handle-CertificateImport -cert $cert -server $Server -newCertPath $newCertPath -allServers $AllExchangeServers
                }
                '10' {
                    Write-Host "Exiting..."; break
                }
                default {
                    Write-Warning "Invalid option."
                }
            }
        } catch {
            Write-Warning $_
            Read-Host "Press Enter to return to the menu"
        }
    } while ($true)
}

function Select-ExchangeServer {
    $servers = Get-ExchangeServer | Sort-Object Name
    if (-not $servers) {
        throw "No Exchange servers found in this environment."
    }

    Write-Host "`nAvailable Exchange Servers:`n"

    for ($i = 0; $i -lt $servers.Count; $i++) {
        $name = $servers[$i].Name
        $fqdn = $servers[$i].Fqdn
        Write-Host "$($i + 1). $name ($fqdn)"
    }

    $index = Read-Host "Select the primary server to operate on (1-$($servers.Count))"
    $index = $index.Trim()

    if ($index -notmatch '^\d+$' -or [int]$index -lt 1 -or [int]$index -gt $servers.Count) {
        throw "Invalid server selection."
    }

    $primary = $servers[$index - 1].Name
    $all = $servers | ForEach-Object { $_.Name }

    Clear-Host

    return @{ Primary = $primary; All = $all }
}

function Show-CertList {
    Write-Host "`nAvailable Exchange Certificates:`n"
    $global:certs = Get-ExchangeCertificate | Sort-Object NotAfter
    $i = 1
    foreach ($cert in $certs) {
        Write-Host "$i. Thumbprint: $($cert.Thumbprint.Substring(0, 8))... | Subject: $($cert.Subject) | Expires: $($cert.NotAfter.ToShortDateString())"
        $i++
    }
}

function Prompt-CertSelection {
    Show-CertList
    $choice = Read-Host "`nSelect a certificate (enter number)"
    if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $certs.Count) {
        Write-Error "Invalid selection."
        exit 1
    }
    Clear-Host
    return $certs[[int]$choice - 1]
}

function Prompt-Action {
    Write-Host "`nWhat do you want to do with this certificate?"
    Write-Host "1. Renew"
    Write-Host "2. Replace"
    Write-Host "3. Revoke"
    Write-Host "4. Delete"
    $action = Read-Host "Enter your choice [1-4]"
    return $action
}

function Handle-CertificateImport($cert, $server, $newCertPath, $allServers) {
    $ext = [System.IO.Path]::GetExtension($newCertPath).ToLower()
    $originalPath = $newCertPath

    switch ($ext) {
        '.p7b' {
            $converted = "$env:TEMP\converted_from_p7b.cer"
            certutil -dump $newCertPath | Out-Null
            certutil -encode $newCertPath $converted | Out-Null
            $newCertPath = $converted
            Write-Host "Converted .p7b to .cer at $converted"
        }
        '.pem' {
            $converted = "$env:TEMP\converted_from_pem.cer"
            certutil -encode $newCertPath $converted | Out-Null
            $newCertPath = $converted
            Write-Host "Converted .pem to .cer at $converted"
        }
        '.crt' {
            Write-Host "Using .crt file directly. Attempting import..."
        }
        '.cer' {
            Write-Host "Using .cer file directly."
        }
        default {
            Write-Warning "Unrecognized extension ($ext). Attempting to import anyway..."
        }
    }

    try {
        $newCert = Import-ExchangeCertificate -FileData ([Byte[]]$(Get-Content $newCertPath -Encoding byte -ReadCount 0)) -Server $server
        if (-not $newCert) {
            Write-Warning "Failed to import certificate from $originalPath"
            return
        }
    } catch {
        Write-Warning "Import error: $_"
        return
    }

    $services = ($cert.Services -join ",").ToUpper()
    Write-Host "Enabling new certificate on $server for services: $services"
    Enable-ExchangeCertificate -Thumbprint $newCert.Thumbprint -Services $services -Force

    $del = Read-Host "Do you want to DELETE the old certificate (Thumbprint $($cert.Thumbprint))? (y/n)"
    if ($del -match '^(y|yes)$') {
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false
        Write-Host "Old certificate deleted."
    } else {
        Write-Host "Old certificate retained."
    }

    $applyAll = Read-Host "Do you want to apply this certificate to ALL Exchange servers? (y/n)"
    if ($applyAll -match '^(y|yes)$') {
        $securePass = Read-Host "Enter export password for .pfx" -AsSecureString
        $tempPfx = "$env:TEMP\$($newCert.Thumbprint).pfx"
        Export-ExchangeCertificate -Thumbprint $newCert.Thumbprint -BinaryEncoded:$true -Password $securePass -Path $tempPfx

        foreach ($srv in $allServers) {
            if ($srv -eq $server) { continue }

            Write-Host "Installing certificate on $srv..."
            $remoteBytes = [System.IO.File]::ReadAllBytes($tempPfx)
            $session = New-PSSession -ComputerName $srv

            Invoke-Command -Session $session -ScriptBlock {
                param($bytes, $pw, $svc)
                $thumb = (Import-ExchangeCertificate -FileData $bytes -Password $pw).Thumbprint
                Enable-ExchangeCertificate -Thumbprint $thumb -Services $svc -Force
            } -ArgumentList $remoteBytes, $securePass, $services

            Remove-PSSession $session
        }

        Remove-Item $tempPfx -Force
    }
}

function Export-Certificate($cert) {
    $path = Read-Host "Enter the full output path for the exported certificate (.pfx)"
    if (-not $path.ToLower().EndsWith(".pfx")) {
        Write-Error "Export path must end with .pfx"
        return
    }

    $folder = Split-Path $path -Parent
    if (-not (Test-Path $folder)) {
        Write-Host "Creating folder: $folder"
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $securePass = Read-Host "Enter password to protect the .pfx" -AsSecureString
    Export-ExchangeCertificate -Thumbprint $cert.Thumbprint -BinaryEncoded:$true -Password $securePass -Path $path
    Write-Host "Certificate exported to $path"
}

function Renew-Certificate($cert, $server, $allServers) {
    # Check for existing pending certificates
    $pending = Get-ExchangeCertificate -Server $server | Where-Object { $_.Status -eq "Pending" }
    if ($pending) {
        Write-Host "Warning: There are existing pending certificates on $server."
        $cleanup = Read-Host "Do you want to list and optionally delete them? (y/n)"
        if ($cleanup -match '^(y|yes)$') {
            $pending | Format-Table Thumbprint, Subject, NotBefore, NotAfter
            $del = Read-Host "Delete all pending certs? (y/n)"
            if ($del -match '^(y|yes)$') {
                $pending | Remove-ExchangeCertificate -Confirm:$false
                Write-Host "Pending certs deleted."
            }
        }

        $resume = Read-Host "Do you want to RESUME a pending request with a signed cert file? (y/n)"
        if ($resume -match '^(y|yes)$') {
            $newCertPath = Read-Host "Enter the full path to the signed certificate (.cer, .crt, .pem, .p7b)"
            if (-not (Test-Path $newCertPath)) {
                Write-Warning "File not found: $newCertPath"
                return
            }
            return Handle-CertificateImport $cert $server $newCertPath $allServers
        }
    }

    # Prompt for CSR output path
    $reqPath = Read-Host "Enter the FULL path and file name where the CSR (.req) should be saved (e.g. C:\certs\<servername>_renew.req or \\server\share\<servername>.req)"
    if (-not $reqPath.ToLower().EndsWith(".req")) {
        Write-Error "The output file must have a .req extension."
        return
    }

    $reqFolder = Split-Path $reqPath -Parent
    if (-not (Test-Path $reqFolder)) {
        Write-Host "Creating folder: $reqFolder"
        New-Item -ItemType Directory -Path $reqFolder -Force | Out-Null
    }

    # Generate CSR
    Write-Host "Generating CSR for thumbprint $($cert.Thumbprint)..."
    $csrText = $cert | New-ExchangeCertificate -GenerateRequest -KeySize 2048 -PrivateKeyExportable $true -Server $server
    [System.IO.File]::WriteAllBytes($reqPath, [System.Text.Encoding]::Unicode.GetBytes($csrText))
    Write-Host "CSR saved to $reqPath"
    Start-Process notepad.exe $reqPath

    $newCertPath = Read-Host "Once the certificate is signed, enter the full path to the new certificate file (.cer, .crt, .pem, or .p7b)"
    if (-not (Test-Path $newCertPath)) {
        Write-Warning "File not found: $newCertPath"
        return
    }

    return Handle-CertificateImport $cert $server $newCertPath $allServers
}

function Resume-PendingCertificate {
    $pending = Get-ExchangeCertificate | Where-Object { $_.Status -eq "Pending" }
    if (-not $pending) {
        Write-Host "No pending certificate renewals found."
        return
    }

    $pending | Format-Table Thumbprint, Subject, NotBefore, NotAfter
    $thumb = Read-Host "Enter the thumbprint of the pending cert to complete"
    $cert = Get-ExchangeCertificate -Thumbprint $thumb
    if (-not $cert) {
        Write-Error "No matching pending certificate found."
        return
    }

    $serverSelection = Select-ExchangeServer
    Renew-Certificate -cert $cert -server $serverSelection.Primary -allServers $serverSelection.All
}

function Replace-Certificate($cert, $server, $allServers) {
    $newCertPath = Read-Host "Enter path to new certificate file (.cer)"
    if (-not (Test-Path $newCertPath)) {
        Write-Error "File not found."
        exit 2
    }

    $newCert = Import-ExchangeCertificate -FileData ([Byte[]]$(Get-Content $newCertPath -Encoding byte -ReadCount 0)) -Server $server
    Enable-ExchangeCertificate -Thumbprint $newCert.Thumbprint -Services "SMTP, IIS" -Force
    Write-Host "Certificate replaced and enabled for SMTP/IIS on $server."

    $applyAll = Read-Host "Do you want to apply this certificate to ALL Exchange servers? (y/n)"
    if ($applyAll -match '^(y|yes)$') {
        $securePass = Read-Host "Enter export password for .pfx" -AsSecureString
        $tempPfx = "$env:TEMP\$($newCert.Thumbprint).pfx"
        Export-ExchangeCertificate -Thumbprint $newCert.Thumbprint -BinaryEncoded:$true -Password $securePass -Path $tempPfx

        foreach ($srv in $allServers) {
            if ($srv -eq $server) { continue }

            Write-Host "Installing certificate on $srv..."
            $remoteBytes = [System.IO.File]::ReadAllBytes($tempPfx)
            $session = New-PSSession -ComputerName $srv

            Invoke-Command -Session $session -ScriptBlock {
                param($bytes, $pw)
                $thumb = (Import-ExchangeCertificate -FileData $bytes -Password $pw).Thumbprint
                Enable-ExchangeCertificate -Thumbprint $thumb -Services "SMTP, IIS" -Force
            } -ArgumentList $remoteBytes, $securePass

            Remove-PSSession $session
        }

        Remove-Item $tempPfx -Force
    }
}

function Revoke-Certificate($cert) {
    $confirm = Read-Host "Are you sure you want to revoke this cert? (y/n)"
    if ($confirm -eq 'y') {
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false
        Write-Host "Certificate revoked."
    } else {
        Write-Host "Revocation cancelled."
    }
}

function Delete-Certificate($cert) {
    $confirm = Read-Host "Are you sure you want to delete this cert? This cannot be undone. (y/n)"
    if ($confirm -eq 'y') {
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false
        Write-Host "Certificate deleted."
    } else {
        Write-Host "Deletion cancelled."
    }
}

# --- Entrypoint ---
if (-not $Server) {
    try {
        $selection = Select-ExchangeServer
        $Server = $selection.Primary
        $AllExchangeServers = $selection.All
    } catch {
        Write-Error "Failed to select a valid Exchange server. $_"
        exit 1
    }
}

Start-CertManager