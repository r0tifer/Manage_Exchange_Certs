# Import Public Functions
Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" | ForEach-Object { . $_.FullName }

# Import Private Functions
Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" | ForEach-Object { . $_.FullName }

Export-ModuleMember -Function Start-CertManager, Renew-Certificate, Replace-Certificate, Export-Certificate, Import-Certificate, Select-ExchangeServer, Revoke-Certificate, Delete-Certificate

# Select Exchange server before launching CLI
try {
    $selection = Select-ExchangeServer
    $global:Server = $selection.Primary
    $global:AllExchangeServers = $selection.All
} catch {
    Write-Error "Failed to select a valid Exchange server. $_"
    exit 1
}

Start-CertManager