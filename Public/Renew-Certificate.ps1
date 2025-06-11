function Renew-Certificate {
    param (
        [Parameter(Mandatory)] $cert,
        [Parameter(Mandatory)] $server,
        [Parameter(Mandatory)] $allServers
    )

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
            return Import-Certificate -cert $cert -server $server -newCertPath $newCertPath -allServers $allServers
        }
    }

    $reqPath = Read-Host "Enter the FULL path and file name where the CSR (.req) should be saved"
    if (-not $reqPath.ToLower().EndsWith(".req")) {
        Write-Error "The output file must have a .req extension."
        return
    }

    $reqFolder = Split-Path $reqPath -Parent
    if (-not (Test-Path $reqFolder)) {
        Write-Host "Creating folder: $reqFolder"
        New-Item -ItemType Directory -Path $reqFolder -Force | Out-Null
    }

    Write-Host "Generating CSR for thumbprint $($cert.Thumbprint)..."
    $csrText = $cert | New-ExchangeCertificate -GenerateRequest -KeySize 2048 -PrivateKeyExportable $true -Server $server
    [System.IO.File]::WriteAllBytes($reqPath, [System.Text.Encoding]::Unicode.GetBytes($csrText))
    Write-Host "CSR saved to $reqPath"
    Start-Process notepad.exe $reqPath

    $newCertPath = Read-Host "Once the certificate is signed, enter the full path to the new certificate file"
    if (-not (Test-Path $newCertPath)) {
        Write-Warning "File not found: $newCertPath"
        return
    }

    return Import-Certificate -cert $cert -server $server -newCertPath $newCertPath -allServers $allServers
}