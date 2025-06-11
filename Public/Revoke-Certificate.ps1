function Revoke-Certificate {
    param (
        [Parameter(Mandatory)] $cert
    )

    $confirm = Read-Host "Are you sure you want to revoke this cert? (y/n)"
    if ($confirm -eq 'y') {
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false
        Write-Host "Certificate revoked."
    } else {
        Write-Host "Revocation cancelled."
    }
}