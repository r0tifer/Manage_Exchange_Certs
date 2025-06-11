function Export-Certificate {
    param (
        [Parameter(Mandatory)] $cert
    )

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