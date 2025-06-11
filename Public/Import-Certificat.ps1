function Import-Certificate {
    param (
        [Parameter(Mandatory)] $cert,
        [Parameter(Mandatory)] $server,
        [Parameter(Mandatory)] $newCertPath,
        [Parameter(Mandatory)] $allServers
    )

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
