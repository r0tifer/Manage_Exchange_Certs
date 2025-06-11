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
                '1' { Clear-Host; Renew-Certificate -cert $global:cert -server $global:Server -allServers $global:AllExchangeServers }
                '2' { Clear-Host; Replace-Certificate -cert $global:cert -server $global:Server -allServers $global:AllExchangeServers }
                '3' { Clear-Host; Revoke-Certificate -cert $global:cert; $global:cert = $null }
                '4' { Clear-Host; Delete-Certificate -cert $global:cert; $global:cert = $null }
                '5' { Clear-Host; Export-Certificate -cert $global:cert }
                '6' {
                    Clear-Host
                    $selection = Select-ExchangeServer
                    $global:Server = $selection.Primary
                    $global:AllExchangeServers = $selection.All
                    $global:cert = $null
                    $global:cert = Prompt-CertSelection
                }
                '7' { Clear-Host; Resume-PendingCertificate }
                '8' {
                    Clear-Host
                    $selection = Select-ExchangeServer
                    $global:Server = $selection.Primary
                    $global:AllExchangeServers = $selection.All
                    $global:cert = $null
                }
                '9' {
                    Clear-Host
                    $newCertPath = Read-Host "Enter full path to new certificate file"
                    if (-not (Test-Path $newCertPath)) {
                        Write-Warning "File not found."
                        pause
                        return
                    }
                    $global:cert = Prompt-CertSelection
                    Import-Certificate -cert $global:cert -server $global:Server -newCertPath $newCertPath -allServers $global:AllExchangeServers
                }
                '10' {
                    Clear-Host
                    Write-Host "Exiting..."
                    return
                }
                default {
                    Write-Warning "Invalid option."
                }
            }
        } catch {
            Write-Warning $_
            Read-Host "Press Enter to return to the menu"
        }
    } while ($true)
}
