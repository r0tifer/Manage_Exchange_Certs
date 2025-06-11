function Prompt-CertSelection {
    Write-Host "`nWelcome to the Manage-ExchangeCert CLI.`n"
    Show-CertList
    $choice = Read-Host "`nSelect a certificate (enter number)"
    if ($choice -notmatch '^\d+$' -or [int]$choice -lt 1 -or [int]$choice -gt $certs.Count) {
        Write-Error "Invalid selection."
        exit 1
    }
    Clear-Host
    return $certs[[int]$choice - 1]
}