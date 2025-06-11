function Start-CertManager {
    Clear-Host
    do {
        try {
            if (-not $global:cert) {
                $global:cert = Prompt-CertSelection
            }
            Show-MainMenu
            $mainAction = Read-Host "Enter your choice"

            switch ($mainAction) {
                '1' { Clear-Host; Renew-Certificate -cert $cert -server $Server -allServers $AllExchangeServers }
                '2' { Clear-Host; Replace-Certificate -cert $cert -server $Server -allServers $AllExchangeServers }
                '3' { Clear-Host; Revoke-Certificate -cert $cert; $cert = $null }
                '4' { Clear-Host; Delete-Certificate -cert $cert; $cert = $null }
                '5' { Clear-Host; Export-Certificate -cert $cert }
                '6' { Clear-Host; $cert = Prompt-CertSelection }
                '7' { Clear-Host; Resume-PendingCertificate }
                '8' {
                    Clear-Host
                    $selection = Select-ExchangeServer
                    $global:Server = $selection.Primary
                    $global:AllExchangeServers = $selection.All
                    $cert = $null
                }
                '9' {
                    Clear-Host
                    $newCertPath = Read-Host "Enter full path to new certificate file"
                    if (-not (Test-Path $newCertPath)) {
                        Write-Warning "File not found."
                        pause
                        return
                    }
                    $cert = Prompt-CertSelection
                    Import-Certificate -cert $cert -server $Server -newCertPath $newCertPath -allServers $AllExchangeServers
                }
                '10' {
                        Clear-Host
                        Write-Host "Exiting..."
                        return
                    }

                default { Write-Warning "Invalid option." }
            }
        } catch {
            Write-Warning $_
            Read-Host "Press Enter to return to the menu"
        }
    } while ($true)
}
