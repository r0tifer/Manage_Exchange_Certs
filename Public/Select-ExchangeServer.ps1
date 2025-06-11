function Select-ExchangeServer {
    Write-Host "`nWelcome to the Manage-ExchangeCert CLI.`n"
    $servers = Get-ExchangeServer | Sort-Object Name
    if (-not $servers) {
        throw "No Exchange servers found in this environment."
    }

    Write-Host "`nAvailable Exchange Servers:`n"
    for ($i = 0; $i -lt $servers.Count; $i++) {
        $name = $servers[$i].Name
        $fqdn = $servers[$i].Fqdn
        Write-Host "$($i + 1). $name ($fqdn)"
    }

    $index = Read-Host "Select the primary server to operate on (1-$($servers.Count))"
    $index = $index.Trim()

    if ($index -notmatch '^\d+$' -or [int]$index -lt 1 -or [int]$index -gt $servers.Count) {
        throw "Invalid server selection."
    }

    $primary = $servers[$index - 1].Name
    $all = $servers | ForEach-Object { $_.Name }

    Clear-Host
    return @{ Primary = $primary; All = $all }
}
