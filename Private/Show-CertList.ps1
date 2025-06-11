function Show-CertList {
    Write-Host "`nAvailable Exchange Certificates:`n"
    $global:certs = Get-ExchangeCertificate | Sort-Object NotAfter
    $i = 1
    foreach ($cert in $certs) {
        Write-Host "$i. Thumbprint: $($cert.Thumbprint.Substring(0, 8))... | Subject: $($cert.Subject) | Expires: $($cert.NotAfter.ToShortDateString())"
        $i++
    }
}