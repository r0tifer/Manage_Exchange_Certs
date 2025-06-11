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