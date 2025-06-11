# ManageExchangeCert.psm1

# Import Public Functions
Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" | ForEach-Object { . $_.FullName }

# Import Private Functions
Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" | ForEach-Object { . $_.FullName }

# Export all user-facing entry points
Export-ModuleMember -Function Start-CertManager, Renew-Certificate, Replace-Certificate, Export-Certificate, Import-Certificate, Select-ExchangeServer, Revoke-Certificate, Delete-Certificate
