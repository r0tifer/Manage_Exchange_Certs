function Replace-Certificate {
    param (
        [Parameter(Mandatory)] $cert,
        [Parameter(Mandatory)] $server,
        [Parameter(Mandatory)] $allServers
    )

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
