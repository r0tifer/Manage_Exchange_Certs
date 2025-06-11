function Delete-Certificate {
    param (
        [Parameter(Mandatory)] $cert
    )

    $confirm = Read-Host "Are you sure you want to delete this cert? This cannot be undone. (y/n)"
    if ($confirm -eq 'y') {
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false
        Write-Host "Certificate deleted."
    } else {
        Write-Host "Deletion cancelled."
    }
}
